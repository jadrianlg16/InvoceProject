# pdf_processor.py
import fitz  # PyMuPDF
import re
import os
import pandas as pd
import time
import logging
import requests # Ensure this is imported
import json     # Ensure this is imported

# --- Configuration ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
LM_STUDIO_API_URL = "http://localhost:1234/v1/chat/completions" # Default LM Studio endpoint

# --- Regular Expressions (Keep as they are) ---
FOLIO_PATTERNS = [
    re.compile(r'\bFOLIO\s*:\s*([A-Za-z0-9/-]{2,})\b', re.IGNORECASE),
    re.compile(r'\bFolio\s*Interno\s*:?\s+([A-F0-9]{8}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{12})\b', re.IGNORECASE),
    re.compile(r'\b(?:Folio|Factura)\s*(?:N[oO]\.?|#)?\s*:?\s+([A-Za-z0-9/-]{2,})\b', re.IGNORECASE),
    # Added: Look for Folio Interno specifically
    re.compile(r'\bFolio\s+interno\s*:\s*(\d+)\b', re.IGNORECASE),
]
def parse_potentially_multiple_ids(id_str):
    ids_found = set()
    # Normalize separators (commas, 'y', 'a')
    text = re.sub(r'\s+', ' ', id_str).strip()
    text = re.sub(r'\s+\b(Y|A)\b\s+', ',', text, flags=re.IGNORECASE) # 'y' or 'a' as separators
    text = re.sub(r'\s*,\s*', ',', text)
    text = re.sub(r'\s+-\s+', '-', text) # Handle ' a ' or ' - ' ranges like 11428 - 11429
    parts = text.split(',')

    for part in parts:
        part = part.strip().upper()
        # Handle ranges like '11428-11429' or '11428 A 11429' (already converted 'A' to ',')
        range_match = re.match(r'(\d+)\s*-\s*(\d+)', part)
        if range_match:
            try:
                start = int(range_match.group(1))
                end = int(range_match.group(2))
                if start <= end:
                    for i in range(start, end + 1):
                         # Basic check to avoid excessively long generated numbers
                         if len(str(i)) < 10:
                            ids_found.add(str(i))
                continue # Skip further processing for this part
            except ValueError:
                pass # Ignore if conversion fails

        # Find individual numbers possibly mixed with text
        potential_ids_in_part = re.findall(r'[\d/-]{2,}', part) # Find sequences of digits, /, -
        for pid in potential_ids_in_part:
             # Basic validation: contains a digit, not excessively long
             if pid and re.search(r'\d', pid) and len(pid) < 30:
                # Avoid adding parts of UUIDs or dates if possible (heuristic)
                if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', pid): # Avoid YYYY-MM-DD
                   ids_found.add(pid)

    # Fallback: If no IDs found using separators/ranges, check the original cleaned string
    if not ids_found:
        original_cleaned = id_str.strip().upper()
        # Check if the original string itself looks like a plausible ID
        if re.search(r'\d', original_cleaned) and re.fullmatch(r'[\dA-Za-z/-]{2,}', original_cleaned) and len(original_cleaned) < 30:
             ids_found.add(original_cleaned)

    # Final filtering: remove empty strings just in case
    ids_found = {id_val for id_val in ids_found if id_val}

    return sorted(list(ids_found))

ESCRITURA_REGEX = re.compile(
    # Look for 'Escritura', 'ESC', 'Esc. Púb.' etc. followed by number indicators
    r'\b(?:Escritura(?:\s+P[úu]blica)?|ESC(?:\.|\s|\b)|ESCRITURA PUBLICA NUMERO)\s+'
    r'(?:(?:N[oO]\.|Nro\.|Num\.|N[úu]mero|No)s?)?\s*' # Optional number indicators like No., Nro., Numero(s)
    # Capture the ID part: must start with digit/slash, can contain digits, letters, spaces, A, Y, comma, slash, hyphen
    # Make it less greedy and look for reasonable length. Stop before common currency words.
    r'([\d/][\dA-Za-z\sAaYy,/ -]{1,60}?[\dA-Za-z/-])' # Capture group 1: The ID string(s)
    # Negative lookahead: ensure it's not followed by currency units or common unrelated words
    r'(?!\s*(?:pesos|eur|usd|mxn|fecha|cliente|rfc)\b)',
    re.IGNORECASE | re.DOTALL
)

ACTA_REGEX = re.compile(
    # Look for 'Acta(s) Fuera (de) Protocolo'
    r'\b(?:Actas?\s+Fuera\s+(?:de\s+)?Protocolo)\s+'
    r'(?:(?:N[oO]\.|Nro\.|Num\.|N[úu]mero|No)s?)?\s*' # Optional number indicators
    # Capture the ID part, similar logic to Escritura
    r'([\d/][\d\sAaYy,/ -]{1,60}?[\dA-Za-z/-])' # Capture group 1: The ID string(s)
    # Negative lookahead
    r'(?!\s*(?:pesos|eur|usd|mxn|fecha|cliente|rfc)\b)',
    re.IGNORECASE | re.DOTALL
)


# --- Updated LLM Query Function ---
def query_llm_for_details(pdf_text, filename_hint=""):
    """
    Sends PDF text to the local LLM via LM Studio API and attempts to extract
    folio number, document type (Escritura/Acta F.P.), and associated ID number(s).

    Args:
        pdf_text (str): The full text extracted from the PDF.
        filename_hint (str): The original filename, potentially useful context for the LLM.

    Returns:
        dict: A dictionary with keys 'folio_number', 'document_type', 'document_id'.
              Values will be strings, using "N/A" if not found.
              'document_id' can contain comma-separated values if multiple are found.
              Returns None if the API call fails or the response is unusable.
    """
    max_chars = 15000 # Keep truncation for very large files
    truncated_text = pdf_text[:max_chars]
    if len(pdf_text) > max_chars:
        logging.warning(f"LLM Query ({filename_hint}): Text truncated to {max_chars} characters for LLM.")

    # --- REVISED & STRENGTHENED PROMPT ---
    system_prompt = (
        "You are an expert automated assistant. Your *only* task is to extract specific data points from Mexican invoice text (Facturas) and return them *exclusively* in JSON format. "
        "Analyze the provided text and identify these three fields:\n"
        "1.  **`folio_number`**: The main invoice identifier. Look for 'Folio:', 'Factura No.:', 'Folio interno:'. If multiple are present, prioritize the Folio Interno . If none are clearly labeled, look for potential standalone identifiers like '123' or just numbers like '35' if they appear as a primary identifier. Use 'N/A' if none found.\n"
        "2.  **`document_type`**: Determine if the text mentions 'Escritura' (or abbreviations like 'ESC', 'Esc. Púb.') or 'Acta Fuera de Protocolo'. Respond with 'Escritura', 'Acta Fuera de Protocolo', or 'Both' if both are clearly mentioned *with associated numbers*. If neither is mentioned, respond with 'N/A'.\n"
        "3.  **`document_id`**: Find the specific number(s) associated *directly* with the `document_type` identified. These are often numeric but can contain '/' or '-'.\n"
        "    - If `document_type` is 'Escritura' or 'Acta Fuera de Protocolo', extract *all* associated numbers. If multiple numbers are listed for that single type (e.g., 'Escrituras 123 y 456', 'Actas 789, 790'), return them as a single comma-separated string (e.g., '123, 456' or '789, 790').\n"
        "    - If `document_type` is 'Both', list Escritura numbers first (comma-separated), then a semicolon ';', then Acta numbers (comma-separated). Example: '123, 456; 789, 790'.\n"
        "    - If `document_type` is 'N/A', or if a type is found but *no* specific number is associated with it, use 'N/A' for this field.\n\n"
        "**IMPORTANT:** Your response MUST be *only* a single, valid JSON object. Do NOT include any introduction, explanation, apologies, or any text before or after the JSON object. Start your response *immediately* with `{` and end it with `}`.\n\n"
        "Example of the required JSON format:\n"
        "```json\n"
        "{\n"
        "  \"folio_number\": \"<extracted_folio or N/A>\",\n"
        "  \"document_type\": \"<Escritura|Acta Fuera de Protocolo|Both|N/A>\",\n"
        "  \"document_id\": \"<extracted_id(s) or N/A>\"\n"
        "}\n"
        "```"
    )

    user_prompt = (
        f"Analyze the following text extracted from an invoice PDF (filename hint: {filename_hint}). Extract the Folio Number, Document Type, and associated Document ID(s) according to the rules provided in the system prompt. Remember to respond ONLY with the JSON object.\n\n"
        f"PDF Text:\n```\n{truncated_text}\n```\n\n"
        f"JSON Output:" # Cue for the model to start the JSON
    )
    # --- END REVISED PROMPT ---

    payload = {
        "model": "local-model", # LM Studio ignores this, but good practice
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 0.05, # Slightly lower temperature for more deterministic output
        "max_tokens": 300,  # Increased slightly just in case IDs are long or numerous
        "stream": False,
        # Consider adding stop sequences if your model supports it, e.g., stop=["\n"]
        # to prevent rambling after the JSON, but this can be tricky.
    }

    headers = {"Content-Type": "application/json"}

    try:
        logging.info(f"LLM Query ({filename_hint}): Sending request to {LM_STUDIO_API_URL}")
        response = requests.post(LM_STUDIO_API_URL, headers=headers, json=payload, timeout=120) # Increased timeout further
        response.raise_for_status() # Check for HTTP errors like 4xx, 5xx

        response_data = response.json()
        logging.debug(f"LLM Raw Response ({filename_hint}): {json.dumps(response_data, indent=2)}") # Log the full raw response for debugging

        if response_data.get("choices") and len(response_data["choices"]) > 0:
            message = response_data["choices"][0].get("message", {})
            message_content = message.get("content")

            # Ensure message_content is a string before proceeding
            if not isinstance(message_content, str):
                logging.error(f"LLM Query ({filename_hint}): Response content is not a string: {type(message_content)}")
                return None

            message_content = message_content.strip() # Remove leading/trailing whitespace

            if message_content:
                logging.debug(f"LLM Raw Content ({filename_hint}):\n{message_content}")
                # Attempt to extract JSON robustly
                extracted_data = None
                json_string_found = None

                try:
                    # Ideal case: The entire response is the JSON object
                    extracted_data = json.loads(message_content)
                    json_string_found = message_content
                    logging.debug(f"LLM Query ({filename_hint}): Successfully parsed entire response as JSON.")

                except json.JSONDecodeError:
                    logging.warning(f"LLM Query ({filename_hint}): Direct JSON parsing failed. Searching for JSON within content.")
                    # Try finding JSON within ```json ... ``` block
                    json_match = re.search(r'```json\s*(\{.*?\})\s*```', message_content, re.DOTALL | re.IGNORECASE)
                    if json_match:
                        json_string_found = json_match.group(1)
                        logging.debug(f"LLM Query ({filename_hint}): Found JSON within ```json block.")
                    else:
                        # Fallback: Find the first standalone JSON object {}
                        json_match = re.search(r'(\{.*?\})', message_content, re.DOTALL)
                        if json_match:
                             json_string_found = json_match.group(1)
                             logging.debug(f"LLM Query ({filename_hint}): Found JSON using general {{...}} search.")

                    if json_string_found:
                        try:
                            extracted_data = json.loads(json_string_found)
                        except json.JSONDecodeError as json_e:
                            logging.error(f"LLM Query ({filename_hint}): Failed to parse extracted JSON string: {json_e}\nExtracted String: {json_string_found}\nOriginal Content:\n{message_content}")
                            return None
                    else:
                        logging.error(f"LLM Query ({filename_hint}): Could not find any valid JSON object within the response content:\n{message_content}")
                        return None

                # --- Validation and Normalization (Keep this robust logic) ---
                if extracted_data:
                    validated_data = {}
                    required_keys = ['folio_number', 'document_type', 'document_id']
                    # Added 'Both' to allowed types
                    allowed_types = ['Escritura', 'Acta Fuera de Protocolo', 'Both', 'N/A']

                    # Check keys (case-insensitive), assign value or "N/A"
                    extracted_keys_lower = {k.lower(): k for k in extracted_data.keys()}
                    for req_key in required_keys:
                        original_key = extracted_keys_lower.get(req_key)
                        # Check if key exists and value is not None
                        if original_key and extracted_data.get(original_key) is not None:
                            value = str(extracted_data[original_key]).strip() # Ensure string, strip whitespace
                            validated_data[req_key] = value if value else "N/A" # Use N/A if value is empty string
                        else:
                            logging.warning(f"LLM Query ({filename_hint}): Missing or null key '{req_key}' in response JSON. Setting to N/A.")
                            validated_data[req_key] = "N/A"

                    # Validate document_type value (case-sensitive after potential extraction)
                    doc_type = validated_data.get('document_type', 'N/A')
                    if doc_type not in allowed_types:
                         logging.warning(f"LLM Query ({filename_hint}): Invalid document_type '{doc_type}' received from LLM. Allowed: {allowed_types}. Setting to N/A.")
                         validated_data['document_type'] = 'N/A'

                    # If type is N/A, ID should also be N/A
                    if validated_data['document_type'] == 'N/A':
                        if validated_data['document_id'] != 'N/A':
                             logging.warning(f"LLM Query ({filename_hint}): Document type is N/A, but document_id is '{validated_data['document_id']}'. Setting ID to N/A.")
                             validated_data['document_id'] = 'N/A'
                    # If type is NOT N/A, but ID *is* N/A, log a warning but allow it (as per prompt instructions)
                    elif validated_data['document_id'] == 'N/A':
                        logging.warning(f"LLM Query ({filename_hint}): Document type is '{validated_data['document_type']}' but document_id is N/A.")

                    logging.info(f"LLM Query ({filename_hint}): Successfully parsed and validated JSON: {validated_data}")
                    return validated_data
                else:
                    # This case should theoretically not be reached if parsing logic is correct
                    logging.error(f"LLM Query ({filename_hint}): Reached end of parsing without valid extracted_data.")
                    return None
                # --- End Validation ---

            else:
                 logging.error(f"LLM Query ({filename_hint}): 'content' field is empty in response choice.")
                 return None
        else:
            logging.error(f"LLM Query ({filename_hint}): 'choices' field missing or empty in API response.")
            return None

    except requests.exceptions.Timeout:
        logging.error(f"LLM Query ({filename_hint}): API request timed out after {payload.get('timeout', 'default')} seconds.")
        return None
    except requests.exceptions.ConnectionError as e:
        logging.error(f"LLM Query ({filename_hint}): API connection failed. Is LM Studio server running at {LM_STUDIO_API_URL}? Error: {e}")
        return None
    except requests.exceptions.RequestException as e:
        logging.error(f"LLM Query ({filename_hint}): API request failed: {e}")
        return None
    except Exception as e:
        # Log the full traceback for unexpected errors
        logging.exception(f"LLM Query ({filename_hint}): Unexpected error during LLM query.")
        return None


# --- Main Analysis Function (Generator - Modified to use improved LLM function) ---
def analyze_pdfs(pdf_directory, output_excel_path):
    yield "PROCESS_START"
    yield f"Starting analysis..."
    yield f"Input directory: {pdf_directory}"
    yield f"Output file: {output_excel_path}"
    logging.info(f"Starting analysis. Input: {pdf_directory}, Output: {output_excel_path}")

    # --- Directory Checks (Keep as is) ---
    if not os.path.isdir(pdf_directory):
        yield f"ERROR: Input directory not found: {pdf_directory}"
        logging.error(f"Input directory not found: {pdf_directory}")
        yield "PROCESS_END"
        return

    output_dir = os.path.dirname(output_excel_path)
    if output_dir and not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
            yield f"INFO: Created output directory: {output_dir}"
            logging.info(f"Created output directory: {output_dir}")
        except Exception as e:
            yield f"ERROR: Could not create output directory '{output_dir}': {e}"
            logging.error(f"Could not create output directory '{output_dir}': {e}")
            yield "PROCESS_END"
            return

    all_extracted_data = []
    pdf_files = [f for f in os.listdir(pdf_directory) if f.lower().endswith('.pdf')]
    total_files = len(pdf_files)
    yield f"Found {total_files} PDF files."
    logging.info(f"Found {total_files} PDF files.")

    if total_files == 0:
        yield "WARNING: No PDF files found in the directory."
        logging.warning("No PDF files found.")
        yield "PROCESS_END"
        return

    processed_files = 0
    errors_occurred = 0
    llm_fallbacks = 0
    llm_successes = 0

    # --- PDF Processing Loop ---
    for filename in pdf_files:
        processed_files += 1
        pdf_path = os.path.join(pdf_directory, filename)
        yield f"PROGRESS:{processed_files}/{total_files}"
        yield f"INFO: Processing file {processed_files}/{total_files}: {filename}"
        logging.info(f"Processing file {processed_files}/{total_files}: {filename}")

        full_text = ""
        clean_text = ""
        doc = None
        regex_folio_number = "Not Found"
        regex_file_findings = set() # Store tuples of (Type, ID)
        llm_processed = False # Flag to indicate if LLM successfully provided data

        try:
            # --- PDF Extraction ---
            doc = fitz.open(pdf_path)
            if doc.is_encrypted:
                yield f"WARNING: File '{filename}' is encrypted. Attempting empty password..."
                logging.warning(f"File '{filename}' is encrypted.")
                if not doc.authenticate(""):
                    yield f"ERROR: File '{filename}' is password protected and could not be opened. Skipping."
                    logging.error(f"Failed to decrypt '{filename}'.")
                    errors_occurred += 1
                    if doc: doc.close()
                    continue

            for page_num in range(len(doc)):
                try:
                    page = doc.load_page(page_num)
                    # Extract text preserving layout slightly better for regex if needed
                    page_text = page.get_text("text", sort=True)
                    full_text += page_text + "\n"
                except Exception as page_e:
                    yield f"WARNING: Could not extract text from page {page_num + 1} in '{filename}': {page_e}"
                    logging.warning(f"Text extraction error on page {page_num + 1} in '{filename}': {page_e}")

            # Basic cleaning for regex and LLM
            clean_text = re.sub(r'[ \t]+', ' ', full_text) # Consolidate spaces/tabs
            clean_text = re.sub(r'\n\s*\n', '\n', clean_text) # Consolidate multiple newlines
            clean_text = clean_text.strip()

            if not clean_text:
                yield f"WARNING: No text extracted from '{filename}'. Skipping."
                logging.warning(f"No text extracted from '{filename}'.")
                if doc: doc.close()
                continue

            # --- Attempt Regex Extraction ---
            logging.debug(f"Regex Attempt ({filename}): Starting extraction.")
            # 1. Folio Number
            found_folio = False
            for i, pattern in enumerate(FOLIO_PATTERNS):
                # Search the first ~3000 chars first for performance, then full text if needed
                search_area = clean_text[:3000]
                match = pattern.search(search_area)
                if not match:
                     match = pattern.search(clean_text) # Search full text if not in first part

                if match:
                    potential_folio = match.group(1).strip()
                    # Basic validation: Avoid overly long strings unless it's a UUID
                    is_uuid = re.fullmatch(r'[A-F0-9]{8}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{12}', potential_folio, re.IGNORECASE)
                    if potential_folio and (len(potential_folio) < 20 or is_uuid):
                        regex_folio_number = potential_folio
                        logging.info(f"Regex Found Folio '{regex_folio_number}' in '{filename}' (pattern {i+1}).")
                        found_folio = True
                        break # Stop after first valid folio match
            if not found_folio:
                 logging.info(f"Regex: Folio number not found in '{filename}'.")

            # 2. Escritura and Acta IDs
            # Find Escrituras
            for match in ESCRITURA_REGEX.finditer(clean_text):
                 captured_string = match.group(1)
                 parsed_ids = parse_potentially_multiple_ids(captured_string)
                 logging.debug(f"Regex Escritura raw: '{captured_string}', Parsed: {parsed_ids} in '{filename}'")
                 for esc_num in parsed_ids:
                     regex_file_findings.add(("Escritura", esc_num))

            # Find Actas
            for match in ACTA_REGEX.finditer(clean_text):
                 captured_string = match.group(1)
                 parsed_ids = parse_potentially_multiple_ids(captured_string)
                 logging.debug(f"Regex Acta F.P. raw: '{captured_string}', Parsed: {parsed_ids} in '{filename}'")
                 for act_num in parsed_ids:
                    regex_file_findings.add(("Acta Fuera de Protocolo", act_num))

            logging.debug(f"Regex Attempt ({filename}): Found {len(regex_file_findings)} Type/ID pairs.")

            # --- Check Regex Success & LLM Fallback ---
            # Trigger LLM if Folio OR Type/ID is missing
            is_incomplete = (regex_folio_number == "Not Found" or not regex_file_findings)

            if is_incomplete and clean_text:
                llm_fallbacks += 1
                yield f"INFO: Regex extraction incomplete for '{filename}'. Attempting LLM fallback..."
                logging.info(f"Attempting LLM fallback for '{filename}'. Regex Folio='{regex_folio_number}', Regex IDs Found={len(regex_file_findings)>0}")

                # Call the improved LLM function
                llm_result = query_llm_for_details(clean_text, filename_hint=filename)

                if llm_result:
                    llm_successes += 1
                    # Use LLM results, potentially overwriting regex findings if LLM is more complete
                    llm_folio = llm_result.get('folio_number', 'N/A')
                    llm_type = llm_result.get('document_type', 'N/A')
                    llm_id = llm_result.get('document_id', 'N/A')

                    yield f"SUCCESS: LLM fallback for '{filename}'. Found: Folio='{llm_folio}', Type='{llm_type}', ID='{llm_id}'"
                    logging.info(f"LLM fallback success for '{filename}': Folio='{llm_folio}', Type='{llm_type}', ID='{llm_id}'")

                    llm_processed = True # Mark that LLM data should be used

                    # --- Add LLM data to results ---
                    # Decide which folio to use: LLM's if found, otherwise stick to Regex's if it found one
                    final_folio = llm_folio if llm_folio != 'N/A' else regex_folio_number

                    # Handle the potentially complex ID field based on type
                    if llm_type == "Both" and llm_id != 'N/A' and ';' in llm_id:
                         # Split IDs if type is Both and semicolon is present
                         esc_ids_str, act_ids_str = map(str.strip, llm_id.split(';', 1))
                         # Add row for Escritura(s) if IDs are present
                         if esc_ids_str and esc_ids_str != 'N/A':
                             all_extracted_data.append({
                                "Filename": filename, "Folio_Number": final_folio,
                                "Type": "Escritura", "ID_Number": esc_ids_str
                             })
                         # Add row for Acta(s) if IDs are present
                         if act_ids_str and act_ids_str != 'N/A':
                             all_extracted_data.append({
                                "Filename": filename, "Folio_Number": final_folio,
                                "Type": "Acta Fuera de Protocolo", "ID_Number": act_ids_str
                             })
                         # If split results in empty strings or N/A for both, add a single N/A row
                         elif not (esc_ids_str and esc_ids_str != 'N/A') and not (act_ids_str and act_ids_str != 'N/A'):
                              all_extracted_data.append({
                                "Filename": filename, "Folio_Number": final_folio,
                                "Type": "Both", "ID_Number": "N/A" # Indicate Both type was found but no specific IDs
                             })

                    elif llm_type in ["Escritura", "Acta Fuera de Protocolo"]:
                         # Add single row, ID field might contain comma-separated values from LLM
                         all_extracted_data.append({
                            "Filename": filename, "Folio_Number": final_folio,
                            "Type": llm_type, "ID_Number": llm_id # Keep potentially comma-separated IDs
                         })
                    else: # Type is N/A or Both without clear separation/IDs
                        # If LLM found a folio but no type/id, record that folio
                        all_extracted_data.append({
                            "Filename": filename, "Folio_Number": final_folio,
                            # Report 'Both' if LLM said so but IDs were N/A or badly formatted
                            "Type": "Both" if llm_type == "Both" else "N/A",
                            "ID_Number": "N/A" # ID is N/A if type is N/A or if Both had no valid IDs
                        })
                    # --- End LLM data adding ---

                else:
                    yield f"WARNING: LLM fallback failed for '{filename}'. Using partial regex results if available."
                    logging.warning(f"LLM fallback failed for '{filename}'. Will use regex results.")
                    # No changes needed here, the code will proceed to add regex results below
                    # because llm_processed is still False

            # --- Add Regex findings (ONLY if LLM wasn't successfully used) ---
            if not llm_processed:
                if regex_file_findings:
                    # Add a row for each unique Type/ID pair found by regex
                    for type_found, id_found in sorted(list(regex_file_findings)):
                        all_extracted_data.append({
                            "Filename": filename,
                            "Folio_Number": regex_folio_number, # Use the folio found by regex
                            "Type": type_found,
                            "ID_Number": id_found # Regex parser gives individual IDs
                        })
                    # If regex found IDs but no folio, the Folio_Number will be "Not Found"
                    if regex_folio_number == "Not Found":
                         logging.warning(f"Regex found Type/ID(s) but no Folio for '{filename}'. Folio set to 'Not Found'.")

                else:
                    # No type/ID found by either method (Regex failed, and LLM either failed or wasn't triggered)
                    # Add a single row with N/A for Type/ID, using the regex folio result
                    logging.info(f"No specific Type/IDs found via Regex or successful LLM for '{filename}'. Adding N/A row for Type/ID.")
                    all_extracted_data.append({
                        "Filename": filename,
                        "Folio_Number": regex_folio_number, # Could be "Not Found" or an actual folio
                        "Type": "N/A",
                        "ID_Number": "N/A"
                    })

        # --- Error Handling and Cleanup ---
        except fitz.fitz.FileNotFoundError:
            yield f"ERROR: File not found (skipped): {pdf_path}"
            logging.error(f"File not found: {pdf_path}")
            errors_occurred += 1
        except fitz.fitz.FileDataError as fd_err:
             yield f"ERROR: File '{filename}' is corrupted or not a valid PDF: {fd_err}. Skipping."
             logging.error(f"Corrupted or invalid PDF '{filename}': {fd_err}")
             errors_occurred += 1
        except Exception as e:
            yield f"ERROR: Unexpected error processing file '{filename}': {e}"
            logging.exception(f"Unexpected error processing file '{filename}':") # Logs stack trace
            errors_occurred += 1
        finally:
            if doc:
                try:
                    doc.close()
                except Exception as close_e:
                     # This is minor, just log it
                     logging.warning(f"Minor error closing file '{filename}': {close_e}")


    # --- Export to Excel/CSV ---
    yield f"INFO: Analysis complete. Processed {processed_files} files with {errors_occurred} errors. LLM fallback attempted for {llm_fallbacks} files ({llm_successes} successful)."
    logging.info(f"Analysis complete. Processed={processed_files}, Errors={errors_occurred}, LLM Fallbacks={llm_fallbacks}, LLM Successes={llm_successes}")

    if all_extracted_data:
        yield "INFO: Creating output file..."
        try:
            df = pd.DataFrame(all_extracted_data)
            # Ensure correct column order
            df = df[["Filename", "Folio_Number", "Type", "ID_Number"]]
            # Sort and remove duplicates - crucial if LLM and Regex find the same info,
            # or if 'Both' logic accidentally creates duplicates.
            df = df.sort_values(by=["Filename", "Folio_Number", "Type", "ID_Number"]).drop_duplicates()

            # Replace empty strings or None with "N/A" for consistency before saving
            df.fillna("N/A", inplace=True)
            df.replace("", "N/A", inplace=True)


            if output_excel_path.lower().endswith('.xlsx'):
                df.to_excel(output_excel_path, index=False, engine='openpyxl')
                yield f"SUCCESS: Successfully created Excel file: {output_excel_path}"
                logging.info(f"Successfully created Excel file: {output_excel_path}")
            elif output_excel_path.lower().endswith('.csv'):
                 df.to_csv(output_excel_path, index=False, encoding='utf-8-sig') # Use utf-8-sig for better Excel compatibility with CSV
                 yield f"SUCCESS: Successfully created CSV file: {output_excel_path}"
                 logging.info(f"Successfully created CSV file: {output_excel_path}")
            else:
                 yield f"ERROR: Output file path '{output_excel_path}' must end with .xlsx or .csv"
                 logging.error(f"Invalid output file extension: {output_excel_path}")

        except Exception as e:
            yield f"ERROR: Failed to write output file '{output_excel_path}': {e}"
            logging.exception(f"Failed to write output file '{output_excel_path}':")
    elif errors_occurred == processed_files and processed_files > 0:
         yield f"ERROR: No data extracted and all files encountered errors. Output file not created."
         logging.error("No data extracted due to errors in all files.")
    elif processed_files > 0:
         yield f"INFO: No relevant data found (even after LLM attempts) in the processed files. Output file not created."
         logging.info("No relevant data found matching patterns or via LLM.")
    else:
        # No PDFs found case already handled
        pass

    yield "PROCESS_END"

# --- Example Usage (if running standalone) ---
# if __name__ == "__main__":
#     input_dir = "path/to/your/input/pdfs"
#     output_file = "path/to/your/output/results.xlsx"
#     # Example of iterating through the generator
#     for message in analyze_pdfs(input_dir, output_file):
#         print(message) # Or update a GUI, etc.