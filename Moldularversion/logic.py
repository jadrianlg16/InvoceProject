# logic.py
# Contains data extraction, processing logic, and helper functions.

import os
import re
import fitz  # PyMuPDF
import pandas as pd
import time
import copy # Keep copy here if used by logic, though it seems mostly GUI-related for clipboard
import traceback # For detailed error logging

# --- Regex Patterns ---
REGEX_ESCRITURA_RANGE = r'(?:Escritura|Esc)\s*(?:P[úu]blica)?\s*(?:N[uú]meros?|Nos.?|N°)\s+(\d+)\s+A\s+(\d+)\b'
REGEX_ACTA_RANGE = r'Acta(?:s)?\s+Fuera\s+de\s+Protocolo\s+(?:N[uú]meros?|Nos.?|N°)\s+(\d+)\s+A\s+(\d+)\b'
REGEX_ACTA_RANGE_AL = r'Acta(?:s)?\s+Fuera\s+de\s+Protocolo\s+N[uú]meros?\s+(\d+)\s+AL\s+(\d+)\b' # Handles "NUMEROS start AL end"
REGEX_ESCRITURA_LIST_Y = r'(?:Escritura|Esc)\s*(?:P[úu]blica)?\s*(?:N[uú]meros?|Nos.?|N°)\s+(\d+)\s+Y\s+(\d+)\b'
REGEX_ACTA_LIST_Y = r'Acta(?:s)?\s+Fuera\s+de\s+Protocolo\s+(?:N[uú]meros?|Nos.?|N°)\s+(\d+)\s+Y\s+(\d+)\b'
# Updated to capture all numbers in formats like 016/18428/18
REGEX_ACTA_SPECIAL = r'Acta\s+Fuera\s+de\s+Protocolo\s+N[uú]mero\s+(\d+)(?:/(\d+))?(?:/(\d+))?(?:/(\d+))?(?:/(\d+))?(?:/(\d+))?(?:/(\d+))?(?:/(\d+))?(?:/(\d+))?(?:/(\d+))?\b'
REGEX_ESCRITURA_SINGLE = r'(?:Escritura|Esc)\s*(?:P[úu]blica)?\s*(?:N[uú]mero|No.?|N°)?\s*[-:\s]?\s*(\d+)\b(?!\s+A\s+\d+)(?!/)' # Added negative lookahead for / to avoid dates
REGEX_ACTA_SINGLE = r'Acta\s+Fuera\s+de\s+Protocolo\s+(?:N[uú]mero|No.?|N°)?\s*[-:\s]?\s*(\d+)\b(?!\s+(?:A|AL)\s+\d+)'
# New patterns for DBA Instrumentos format
# For Escrituras
REGEX_INSTRUMENTOS_ESC_SINGLE = r'Instrumentos\s*-\s*Esc\s+(\d+)\b(?!\s*,)'
REGEX_INSTRUMENTOS_ESC_COMMA = r'Instrumentos\s*-\s*Esc\s+(\d+)(?:\s*,\s*Esc\s+(\d+))?(?:\s*,\s*(\d+))?(?:\s*,\s*(\d+))?(?:\s*,\s*(\d+))?'
# For Actas
REGEX_INSTRUMENTOS_ACT_SINGLE = r'Instrumentos\s*-\s*Act(?:a)?\s+(\d+)\b(?!\s*,)'
REGEX_INSTRUMENTOS_ACT_COMMA = r'Instrumentos\s*-\s*Act(?:a)?\s+(\d+)(?:\s*,\s*Act(?:a)?\s+(\d+))?(?:\s*,\s*(\d+))?(?:\s*,\s*(\d+))?(?:\s*,\s*(\d+))?'
# For combined Esc and Act in the same Instrumentos entry
REGEX_INSTRUMENTOS_COMBINED = r'Instrumentos\s*-\s*(?:Esc\s+(\d+)(?:\s*,\s*(?:Esc\s+)?(\d+))?(?:\s*,\s*(?:Esc\s+)?(\d+))?)?(?:\s*,\s*Act(?:a)?\s+(\d+)(?:\s*,\s*(?:Act(?:a)?\s+)?(\d+))?(?:\s*,\s*(?:Act(?:a)?\s+)?(\d+))?)?'
REGEX_FOLIO_DBA = r'(?i)\bSerie\s(?:RP)?\sFolio\s(\d+)\b' # Text search fallback
REGEX_FOLIO_DBA_ALT = r'(?i)DATOS\s+CFDI.?Folio:\s(\d+)' # Text search fallback
REGEX_FOLIO_TOTALNOT = r'(?i)Folio\s+interno:\s*(\w+)\b'
REGEX_FOLIO_DBA_FILENAME = r'_(\d{1,5})_' # Filename extraction

# --- Helper Functions ---
def find_unique_output_filename(base_name="Extracted_Invoices.xlsx"):
    """Finds a unique filename in the current directory, adding _N if needed."""
    directory = os.getcwd()
    output_path = os.path.join(directory, base_name)
    counter = 1
    name, ext = os.path.splitext(base_name)
    while os.path.exists(output_path):
        output_path = os.path.join(directory, f"{name}_{counter}{ext}")
        counter += 1
    return output_path

def extract_text_from_pdf(pdf_path):
    """Extracts text from all pages of a PDF."""
    try:
        doc = fitz.open(pdf_path)
        full_text = ""
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            # Added flags for better text extraction, esp. ligatures and spacing
            full_text += page.get_text("text", flags=fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE) + "\n"
        doc.close()
        # Normalize whitespace: replace multiple spaces/newlines with single space
        full_text = re.sub(r'\s+', ' ', full_text).strip()
        return full_text
    except Exception as e:
        print(f"Error opening or reading PDF {pdf_path}: {e}")
        return None

# --- Extraction Logic ---
def find_folio(text, invoice_type, filename=None):
    """
    Finds the folio number based on invoice type.
    Prioritizes filename for DBA, then falls back to text search.
    """
    folio = None
    if not text and invoice_type != 'DBA': # DBA might find folio in filename even if text fails
        print(f"DEBUG: Text is None for non-DBA file: {filename}. Folio set to NOT_FOUND.")
        return "NOT_FOUND"

    # DBA: Prioritize Filename
    if invoice_type == 'DBA':
        if filename:
            filename_match = re.search(REGEX_FOLIO_DBA_FILENAME, filename, re.IGNORECASE)
            if filename_match:
                folio_candidate = filename_match.group(1)
                # Basic validation: is it a number and not excessively long?
                if folio_candidate.isdigit() and len(folio_candidate) < 20:
                    print(f"DEBUG: Folio found in DBA filename '{filename}': {folio_candidate}")
                    folio = folio_candidate
                else:
                     print(f"DEBUG: Folio candidate from DBA filename '{filename}' rejected (not digit or too long): {folio_candidate}")

        # Fallback to Text Search for DBA (only if not found in filename and text exists)
        if not folio and text:
            print(f"DEBUG: Folio not found in DBA filename '{filename}'. Searching text.")
            match = re.search(REGEX_FOLIO_DBA, text, re.IGNORECASE)
            if match:
                folio = match.group(1)
                print(f"DEBUG: Folio found in DBA text (REGEX_FOLIO_DBA): {folio}")
            else:
                match_alt = re.search(REGEX_FOLIO_DBA_ALT, text, re.IGNORECASE | re.DOTALL)
                if match_alt:
                    folio = match_alt.group(1)
                    print(f"DEBUG: Folio found in DBA text (REGEX_FOLIO_DBA_ALT): {folio}")
                else:
                    print(f"DEBUG: Folio not found in DBA text search for: {filename}")
        elif not folio and not text:
             print(f"DEBUG: Text extraction failed for '{filename}', cannot search text for DBA folio.")

    # Other Invoice Types
    elif invoice_type == 'TOTALNOT':
        if text:
            match = re.search(REGEX_FOLIO_TOTALNOT, text, re.IGNORECASE)
            if match:
                folio = match.group(1)
                print(f"DEBUG: Folio found for TOTALNOT: {folio}")
            else:
                print(f"DEBUG: Folio not found in TOTALNOT text search for: {filename}")
        else: # Should not happen due to initial check, but defensive coding
            print(f"DEBUG: Text is None for TOTALNOT, cannot search for folio: {filename}")

    elif invoice_type == 'CONTPAQ':
        if text:
            # Look for "FOLIO: XXX" but try to avoid "Folio fiscal"
            contpaq_simple_pattern = r'\bFOLIO:\s*(\w+)\b'
            candidate_folio = None
            # Search from end to beginning to potentially find the main invoice folio last
            matches = list(re.finditer(contpaq_simple_pattern, text, re.IGNORECASE))
            for match in reversed(matches):
                start_index = match.start()
                # Check text immediately preceding the match
                preceding_text = text[max(0, start_index - 30) : start_index]
                if not re.search(r'\bFolio\s+fiscal\b', preceding_text, re.IGNORECASE):
                    candidate_folio = match.group(1)
                    print(f"DEBUG: Candidate folio found for CONTPAQ: {candidate_folio}")
                    break # Found a likely candidate, stop searching
                else:
                    print(f"DEBUG: Ignoring potential CONTPAQ folio '{match.group(1)}' due to preceding 'Folio fiscal'.")

            if candidate_folio:
                 folio = candidate_folio
            else:
                 print(f"DEBUG: Folio not found in CONTPAQ text search for: {filename}")
        else: # Should not happen due to initial check
            print(f"DEBUG: Text is None for CONTPAQ, cannot search for folio: {filename}")

    # Final Checks & Return
    if folio and len(folio) > 20: # Check if it looks like a UUID (Folio Fiscal)
        print(f"DEBUG: Folio '{folio}' for {filename} looks like Folio Fiscal, rejecting.")
        return "FOLIO_FISCAL_SUSPECTED"
    elif not folio:
        # Specific debug for DBA if it reaches here without finding anything
        if invoice_type == 'DBA':
            print(f"DEBUG: Folio ultimately NOT_FOUND for DBA file: {filename} (Checked filename and text if available)")
        return "NOT_FOUND"

    print(f"DEBUG: Final folio returned for {filename} (Type: {invoice_type}): {folio}")
    return folio


# --- NEW Type-Specific Reference Finders ---
# These are currently identical copies of the original logic.
# They can be customized later for each invoice type's specific format.

def find_references_dba(text):
    """Extracts Escritura and Acta references from text. (DBA specific - with Instrumentos format support)"""
    references = []
    if not text: return []

    # Define patterns to exclude (dates, etc.)
    exclude_patterns = [
        r'Fecha\s+(?:de\s+)?Escritura\s*:\s*\d{1,2}/\d{1,2}/\d{2,4}',  # Fecha Escritura: 29/11/2023
        r'Fecha\s+(?:de\s+)?Acta\s*:\s*\d{1,2}/\d{1,2}/\d{2,4}'        # Fecha Acta: 29/11/2023
    ]

    # Remove or mask text matching exclude patterns
    text_for_processing = text
    for pattern in exclude_patterns:
        text_for_processing = re.sub(pattern, "EXCLUDED_DATE", text_for_processing, flags=re.IGNORECASE)

    # Debug log
    if text != text_for_processing:
        print(f"DEBUG: Excluded date patterns from text for reference extraction")

    found_escritura_numbers = set()
    found_acta_numbers = set()
    flags = re.IGNORECASE | re.UNICODE

    # First check for Instrumentos format (DBA specific)
    instrumentos_patterns = [
        (REGEX_INSTRUMENTOS_ESC_SINGLE, "Escritura", found_escritura_numbers),
        (REGEX_INSTRUMENTOS_ESC_COMMA, "Escritura", found_escritura_numbers),
        (REGEX_INSTRUMENTOS_ACT_SINGLE, "Acta Fuera de Protocolo", found_acta_numbers),
        (REGEX_INSTRUMENTOS_ACT_COMMA, "Acta Fuera de Protocolo", found_acta_numbers)
    ]

    # Process all standard Instrumentos patterns
    for pattern, type_name, found_set in instrumentos_patterns:
        for match in re.finditer(pattern, text_for_processing, flags):
            try:
                # Process all groups that might contain numbers
                for i in range(1, len(match.groups()) + 1):
                    if match.group(i):
                        num_str = match.group(i).strip()
                        if num_str and num_str.isdigit() and num_str not in found_set:
                            print(f"DEBUG: Found {type_name} number {num_str} in Instrumentos format")
                            references.append({"Type": type_name, "Number": num_str})
                            found_set.add(num_str)
            except (ValueError, IndexError) as e:
                print(f"Warning: Could not parse Instrumentos numbers from pattern {pattern}: {match.groups()} - Error: {e}")
                pass

    # Process combined Esc and Act pattern separately
    for match in re.finditer(REGEX_INSTRUMENTOS_COMBINED, text_for_processing, flags):
        try:
            # Groups 1-3 are Escritura numbers
            for i in range(1, 4):
                if match.group(i):
                    num_str = match.group(i).strip()
                    if num_str and num_str.isdigit() and num_str not in found_escritura_numbers:
                        print(f"DEBUG: Found Escritura number {num_str} in combined Instrumentos format")
                        references.append({"Type": "Escritura", "Number": num_str})
                        found_escritura_numbers.add(num_str)

            # Groups 4-6 are Acta numbers
            for i in range(4, 7):
                if match.group(i):
                    num_str = match.group(i).strip()
                    if num_str and num_str.isdigit() and num_str not in found_acta_numbers:
                        print(f"DEBUG: Found Acta Fuera de Protocolo number {num_str} in combined Instrumentos format")
                        references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
                        found_acta_numbers.add(num_str)
        except (ValueError, IndexError) as e:
            print(f"Warning: Could not parse combined Instrumentos numbers: {match.groups()} - Error: {e}")
            pass



    # Ranges and Lists First (standard patterns)
    patterns_and_types = [
        (REGEX_ESCRITURA_RANGE, "Escritura", found_escritura_numbers, True),
        (REGEX_ACTA_RANGE, "Acta Fuera de Protocolo", found_acta_numbers, True),
        (REGEX_ACTA_RANGE_AL, "Acta Fuera de Protocolo", found_acta_numbers, True),
        (REGEX_ESCRITURA_LIST_Y, "Escritura", found_escritura_numbers, False),
        (REGEX_ACTA_LIST_Y, "Acta Fuera de Protocolo", found_acta_numbers, False),
        (REGEX_ACTA_SPECIAL, "Acta Fuera de Protocolo", found_acta_numbers, False, True) # Special: only group 1 matters
    ]

    for pattern, type_name, found_set, is_range, *special in patterns_and_types:
        use_special_grouping = special and special[0]
        for match in re.finditer(pattern, text_for_processing, flags):
            try:
                if is_range:
                    start, end = int(match.group(1)), int(match.group(2))
                    if start <= end:
                        for num in range(start, end + 1):
                            num_str = str(num)
                            if num_str not in found_set:
                                references.append({"Type": type_name, "Number": num_str})
                                found_set.add(num_str)
                elif use_special_grouping:
                    # Process all groups for special patterns (like ACTA_SPECIAL with multiple numbers)
                    for i in range(1, len(match.groups()) + 1):
                        if match.group(i):
                            num_str = match.group(i).strip()
                            if num_str and num_str.isdigit() and num_str not in found_set:
                                print(f"DEBUG: Found {type_name} number {num_str} in special format")
                                references.append({"Type": type_name, "Number": num_str})
                                found_set.add(num_str)
                else: # Is List ("Y")
                    num1_str = match.group(1).strip()
                    if num1_str and num1_str.isdigit() and num1_str not in found_set:
                        references.append({"Type": type_name, "Number": num1_str})
                        found_set.add(num1_str)
                    num2_str = match.group(2).strip()
                    if num2_str and num2_str.isdigit() and num2_str not in found_set:
                         references.append({"Type": type_name, "Number": num2_str})
                         found_set.add(num2_str)
            except (ValueError, IndexError) as e:
                print(f"Warning: Could not parse numbers for {type_name} from pattern {pattern}: {match.groups()} - Error: {e}")
                pass # Continue processing other matches/patterns

    # Singles Last
    single_patterns = [
        (REGEX_ESCRITURA_SINGLE, "Escritura", found_escritura_numbers),
        (REGEX_ACTA_SINGLE, "Acta Fuera de Protocolo", found_acta_numbers)
    ]
    for pattern, type_name, found_set in single_patterns:
         # Use finditer to get all non-overlapping matches
         potential_singles = []
         for m in re.finditer(pattern, text_for_processing, flags):
             num_candidate = m.group(1).strip()
             if num_candidate and num_candidate.isdigit():
                 # Add check: ensure it wasn't part of a range/list already processed
                 # This requires checking the context, which is hard with regex alone.
                 # Current logic relies on singles being less specific and processed last.
                 potential_singles.append(num_candidate)

         # Process collected singles
         for num_str in potential_singles:
             if num_str not in found_set: # Check against numbers already found by ranges/lists/previous singles
                 references.append({"Type": type_name, "Number": num_str})
                 found_set.add(num_str) # Add to the set to avoid duplicates from this pattern

    # Sorting
    def sort_key(item):
        try: num_val = int(item["Number"])
        except ValueError: num_val = float('inf') # Place non-numeric refs last
        return (item["Type"], num_val)
    references.sort(key=sort_key)

    return references


def find_references_totalnot(text):
    """Extracts Escritura and Acta references from text. (TOTALNOT specific - currently generic)"""
    references = []
    if not text: return []

    found_escritura_numbers = set()
    found_acta_numbers = set()
    flags = re.IGNORECASE | re.UNICODE

    # Ranges and Lists First
    patterns_and_types = [
        (REGEX_ESCRITURA_RANGE, "Escritura", found_escritura_numbers, True),
        (REGEX_ACTA_RANGE, "Acta Fuera de Protocolo", found_acta_numbers, True),
        (REGEX_ACTA_RANGE_AL, "Acta Fuera de Protocolo", found_acta_numbers, True),
        (REGEX_ESCRITURA_LIST_Y, "Escritura", found_escritura_numbers, False),
        (REGEX_ACTA_LIST_Y, "Acta Fuera de Protocolo", found_acta_numbers, False),
        (REGEX_ACTA_SPECIAL, "Acta Fuera de Protocolo", found_acta_numbers, False, True) # Special: only group 1 matters
    ]

    for pattern, type_name, found_set, is_range, *special in patterns_and_types:
        use_special_grouping = special and special[0]
        for match in re.finditer(pattern, text, flags):
            try:
                if is_range:
                    start, end = int(match.group(1)), int(match.group(2))
                    if start <= end:
                        for num in range(start, end + 1):
                            num_str = str(num)
                            if num_str not in found_set:
                                references.append({"Type": type_name, "Number": num_str})
                                found_set.add(num_str)
                elif use_special_grouping:
                     num_str = match.group(1).strip()
                     if num_str and num_str.isdigit() and num_str not in found_set:
                         references.append({"Type": type_name, "Number": num_str})
                         found_set.add(num_str)
                else: # Is List ("Y")
                    num1_str = match.group(1).strip()
                    if num1_str and num1_str.isdigit() and num1_str not in found_set:
                        references.append({"Type": type_name, "Number": num1_str})
                        found_set.add(num1_str)
                    num2_str = match.group(2).strip()
                    if num2_str and num2_str.isdigit() and num2_str not in found_set:
                         references.append({"Type": type_name, "Number": num2_str})
                         found_set.add(num2_str)
            except (ValueError, IndexError) as e:
                print(f"Warning: Could not parse numbers for {type_name} from pattern {pattern}: {match.groups()} - Error: {e}")
                pass # Continue processing other matches/patterns

    # Singles Last
    single_patterns = [
        (REGEX_ESCRITURA_SINGLE, "Escritura", found_escritura_numbers),
        (REGEX_ACTA_SINGLE, "Acta Fuera de Protocolo", found_acta_numbers)
    ]
    for pattern, type_name, found_set in single_patterns:
         # Use finditer to get all non-overlapping matches
         potential_singles = []
         for m in re.finditer(pattern, text, flags):
             num_candidate = m.group(1).strip()
             if num_candidate and num_candidate.isdigit():
                 potential_singles.append(num_candidate)

         # Process collected singles
         for num_str in potential_singles:
             if num_str not in found_set: # Check against numbers already found by ranges/lists/previous singles
                 references.append({"Type": type_name, "Number": num_str})
                 found_set.add(num_str) # Add to the set to avoid duplicates from this pattern

    # Sorting
    def sort_key(item):
        try: num_val = int(item["Number"])
        except ValueError: num_val = float('inf') # Place non-numeric refs last
        return (item["Type"], num_val)
    references.sort(key=sort_key)

    return references


def find_references_contpaq(text):
    """Extracts Escritura and Acta references from text. (CONTPAQ specific - currently generic)"""
    references = []
    if not text: return []

    found_escritura_numbers = set()
    found_acta_numbers = set()
    flags = re.IGNORECASE | re.UNICODE

    # Ranges and Lists First
    patterns_and_types = [
        (REGEX_ESCRITURA_RANGE, "Escritura", found_escritura_numbers, True),
        (REGEX_ACTA_RANGE, "Acta Fuera de Protocolo", found_acta_numbers, True),
        (REGEX_ACTA_RANGE_AL, "Acta Fuera de Protocolo", found_acta_numbers, True),
        (REGEX_ESCRITURA_LIST_Y, "Escritura", found_escritura_numbers, False),
        (REGEX_ACTA_LIST_Y, "Acta Fuera de Protocolo", found_acta_numbers, False),
        (REGEX_ACTA_SPECIAL, "Acta Fuera de Protocolo", found_acta_numbers, False, True) # Special: only group 1 matters
    ]

    for pattern, type_name, found_set, is_range, *special in patterns_and_types:
        use_special_grouping = special and special[0]
        for match in re.finditer(pattern, text, flags):
            try:
                if is_range:
                    start, end = int(match.group(1)), int(match.group(2))
                    if start <= end:
                        for num in range(start, end + 1):
                            num_str = str(num)
                            if num_str not in found_set:
                                references.append({"Type": type_name, "Number": num_str})
                                found_set.add(num_str)
                elif use_special_grouping:
                     num_str = match.group(1).strip()
                     if num_str and num_str.isdigit() and num_str not in found_set:
                         references.append({"Type": type_name, "Number": num_str})
                         found_set.add(num_str)
                else: # Is List ("Y")
                    num1_str = match.group(1).strip()
                    if num1_str and num1_str.isdigit() and num1_str not in found_set:
                        references.append({"Type": type_name, "Number": num1_str})
                        found_set.add(num1_str)
                    num2_str = match.group(2).strip()
                    if num2_str and num2_str.isdigit() and num2_str not in found_set:
                         references.append({"Type": type_name, "Number": num2_str})
                         found_set.add(num2_str)
            except (ValueError, IndexError) as e:
                print(f"Warning: Could not parse numbers for {type_name} from pattern {pattern}: {match.groups()} - Error: {e}")
                pass # Continue processing other matches/patterns

    # Singles Last
    single_patterns = [
        (REGEX_ESCRITURA_SINGLE, "Escritura", found_escritura_numbers),
        (REGEX_ACTA_SINGLE, "Acta Fuera de Protocolo", found_acta_numbers)
    ]
    for pattern, type_name, found_set in single_patterns:
         # Use finditer to get all non-overlapping matches
         potential_singles = []
         for m in re.finditer(pattern, text, flags):
             num_candidate = m.group(1).strip()
             if num_candidate and num_candidate.isdigit():
                 potential_singles.append(num_candidate)

         # Process collected singles
         for num_str in potential_singles:
             if num_str not in found_set: # Check against numbers already found by ranges/lists/previous singles
                 references.append({"Type": type_name, "Number": num_str})
                 found_set.add(num_str) # Add to the set to avoid duplicates from this pattern

    # Sorting
    def sort_key(item):
        try: num_val = int(item["Number"])
        except ValueError: num_val = float('inf') # Place non-numeric refs last
        return (item["Type"], num_val)
    references.sort(key=sort_key)

    return references

# --- End of NEW Type-Specific Reference Finders ---


def process_single_pdf(pdf_path, invoice_type):
    """
    Processes a single PDF file, returning a list of result dictionaries.
    Calls the appropriate reference finding function based on invoice_type.
    """
    filename = os.path.basename(pdf_path)
    print(f"\n--- Processing PDF: {filename} (Type: {invoice_type}) ---")
    text = extract_text_from_pdf(pdf_path)

    # Handle immediate text extraction failure (unless DBA, where filename folio is possible)
    if text is None and invoice_type != 'DBA':
         print(f"ERROR: Text extraction failed for {filename}. Cannot proceed.")
         return [{"Document Type": "ERROR", "Reference Number": "Text Extraction Failed",
                  "Invoice Folio": "ERROR", "Source PDF": filename,
                  "Full PDF Path": pdf_path}]

    # Find folio first
    folio = find_folio(text, invoice_type, filename) # Pass text even if None for DBA filename check

    # Find references using the type-specific function
    references = []
    if text: # Only attempt reference finding if text was successfully extracted
        print(f"DEBUG: Calling reference finder for type: {invoice_type}")
        if invoice_type == 'DBA':
            references = find_references_dba(text)
        elif invoice_type == 'TOTALNOT':
            references = find_references_totalnot(text)
        elif invoice_type == 'CONTPAQ':
            references = find_references_contpaq(text)
        else:
            print(f"WARNING: Unknown invoice type '{invoice_type}' passed to process_single_pdf. No specific reference finder called.")
            # Optionally, you could call a default or raise an error here
            # For now, references will remain an empty list
        print(f"DEBUG: Found {len(references)} references for {filename}")
    elif invoice_type == 'DBA' and folio != "NOT_FOUND" and folio != "ERROR":
        print(f"DEBUG: Text extraction failed for DBA file {filename}, but folio was found ({folio}). No references can be extracted.")
    else:
         print(f"DEBUG: Text extraction failed for {filename}, cannot search for references.")


    # --- Assemble Output Rows ---
    output_rows = []

    # Case 1: Text extraction failed, but Folio was found (DBA filename scenario)
    if text is None and invoice_type == 'DBA' and folio not in ["NOT_FOUND", "ERROR", "FOLIO_FISCAL_SUSPECTED"]:
         output_rows.append({
            "Document Type": "N/A", "Reference Number": "TEXT_EXTRACTION_FAILED",
            "Invoice Folio": folio, "Source PDF": filename,
            "Full PDF Path": pdf_path })
         print(f"OUTPUT: Row added for DBA with failed text but found folio ({folio})")

    # Case 2: Text extracted, but NO references found
    elif text is not None and not references:
         output_rows.append({
            "Document Type": "N/A", "Reference Number": "N/A",
            "Invoice Folio": folio, "Source PDF": filename,
            "Full PDF Path": pdf_path })
         print(f"OUTPUT: Row added for file with text but no references found (Folio: {folio})")

    # Case 3: References FOUND
    elif references:
        for ref in references:
            output_rows.append({
                "Document Type": ref["Type"], "Reference Number": ref["Number"],
                "Invoice Folio": folio, "Source PDF": filename,
                "Full PDF Path": pdf_path })
        print(f"OUTPUT: {len(references)} rows added based on found references (Folio: {folio})")

    # Case 4: Catch-all for errors / unexpected states (e.g., text failed AND folio not found)
    # This handles the original explicit error condition for combined failure.
    if not output_rows and text is None and folio in ["NOT_FOUND", "ERROR"]:
         err_ref_num = "Text Extraction Failed" if text is None else "Processing Error"
         err_folio = folio if folio == "NOT_FOUND" else "ERROR"
         output_rows.append({
             "Document Type": "ERROR", "Reference Number": err_ref_num,
             "Invoice Folio": err_folio, "Source PDF": filename,
             "Full PDF Path": pdf_path})
         print(f"OUTPUT: ERROR row added (Text Failed: {text is None}, Folio: {folio})")
    elif not output_rows:
        # This case might occur if text extraction worked, folio finding failed, *and* no refs were found.
        # Let's ensure at least one row is always returned per file.
         output_rows.append({
            "Document Type": "N/A", "Reference Number": "N/A",
            "Invoice Folio": folio, # Keep whatever folio status was found
            "Source PDF": filename,
            "Full PDF Path": pdf_path })
         print(f"OUTPUT: Row added for file processed but resulted in no specific data/refs (Folio: {folio})")


    print(f"--- Finished Processing PDF: {filename} ---")
    return output_rows


# --- Main Processing Function (runs in a separate thread, interacts with GUI instance) ---
def run_processing(folder_path, invoice_type, app_instance):
    """
    Scans folder, processes PDFs, updates GUI via app_instance, and saves results.
    Args:
        folder_path (str): The root folder to scan for PDFs.
        invoice_type (str): The type of invoice ('DBA', 'TOTALNOT', 'CONTPAQ').
        app_instance: The instance of the InvoiceProcessorApp GUI class.
    """
    all_data = []
    pdf_files = []
    output_filename = None
    # Need access to GUI methods (messagebox, status updates, etc.)
    master_gui = app_instance.master # Get the root Tk window for scheduling updates
    messagebox = app_instance.messagebox # Use the messagebox from the gui module

    try:
        master_gui.after(0, app_instance.update_status, f"Scanning folder for PDF files: {folder_path}")
        for root, _, files in os.walk(folder_path):
            for file in files:
                # Ignore hidden files (like .DS_Store on macOS) and ensure it's a PDF
                if not file.startswith('.') and file.lower().endswith('.pdf'):
                     pdf_path = os.path.join(root, file)
                     # Double check if it's actually a file (might be a broken link)
                     if os.path.isfile(pdf_path):
                         pdf_files.append(pdf_path)
                     else:
                         print(f"Warning: Skipped non-file entry: {pdf_path}")

    except Exception as e:
        error_msg = f"Error accessing folder structure in {folder_path}: {e}"
        print(error_msg)
        master_gui.after(0, app_instance.update_status, f"Error accessing folder. Check permissions or path.")
        master_gui.after(0, messagebox.showerror, "Folder Error", error_msg, parent=master_gui)
        master_gui.after(0, app_instance.enable_buttons)
        master_gui.after(10, app_instance.clear_pdf_preview, "Error during folder scan.")
        return

    total_files = len(pdf_files)
    if total_files == 0:
        master_gui.after(0, app_instance.update_status, "No PDF files found in the selected folder or subfolders.")
        master_gui.after(0, app_instance.enable_buttons)
        master_gui.after(0, messagebox.showinfo, "Information", "No PDF files found.", parent=master_gui)
        master_gui.after(10, app_instance.clear_pdf_preview, "No PDFs found to process.")
        return

    start_time = time.time()
    files_processed_count = 0 # Files processed *without* resulting in an ERROR row
    files_with_errors = 0     # Files that resulted in at least one ERROR row
    rows_added_count = 0      # Total rows added to all_data

    for i, pdf_path in enumerate(pdf_files):
        current_filename = os.path.basename(pdf_path)
        # Update status less frequently for performance, but always on first/last
        if i == 0 or i == total_files - 1 or (i + 1) % 10 == 0: # Update every 10 files
            status_message = f"Processing file {i+1}/{total_files}: {current_filename}"
            # Use after(0) for immediate scheduling, but allow GUI to refresh
            master_gui.after(0, app_instance.update_status, status_message)
            # Optionally yield control briefly if GUI becomes unresponsive on large batches
            # master_gui.update_idletasks() # Or time.sleep(0.01)

        try:
            # Process the PDF using the updated function
            results = process_single_pdf(pdf_path, invoice_type)

            # Check if processing generated results and if any are error rows
            if results:
                had_error_in_results = any(row.get("Document Type") == "ERROR" for row in results)
                all_data.extend(results) # Add all rows from the result list
                rows_added_count += len(results)

                if had_error_in_results:
                    files_with_errors += 1
                    print(f"INFO: File {current_filename} processed, but resulted in an ERROR row.")
                else:
                    files_processed_count += 1 # Count as successfully processed if no ERROR rows
            else:
                 # This condition should ideally not be reached if process_single_pdf always returns a list
                 files_with_errors += 1
                 print(f"CRITICAL WARNING: process_single_pdf returned empty list/None for {current_filename}")
                 # Add a generic error row to ensure the file is represented
                 all_data.append({
                     "Document Type": "ERROR", "Reference Number": "Processing Function Returned Nothing",
                     "Invoice Folio": "ERROR", "Source PDF": current_filename,
                     "Full PDF Path": pdf_path })
                 rows_added_count += 1

        except Exception as e:
            files_with_errors += 1 # Count this file as having an error
            error_msg = f"CRITICAL RUNTIME ERROR processing file {current_filename}: {e}"
            print(error_msg)
            traceback.print_exc() # Print full traceback to console
            # Update GUI status bar about the error
            master_gui.after(0, app_instance.update_status, f"ERROR processing {current_filename}. See console.")
            # Add an error row to the data
            all_data.append({
                "Document Type": "ERROR", "Reference Number": f"Runtime Error: {e}",
                "Invoice Folio": "ERROR", "Source PDF": current_filename,
                "Full PDF Path": pdf_path })
            rows_added_count += 1
            # Optionally show a popup for critical errors? Might be too many.
            # master_gui.after(0, messagebox.showerror, "Processing Error", error_msg, parent=master_gui)

    end_time = time.time()
    processing_time = end_time - start_time
    # Refined summary message
    final_summary = f"{files_processed_count + files_with_errors}/{total_files} files attempted."
    if files_processed_count > 0:
        final_summary += f" {files_processed_count} processed without major errors."
    if files_with_errors > 0:
        final_summary += f" {files_with_errors} encountered errors (check ERROR rows/console)."
    final_summary += f"\nGenerated {rows_added_count} total rows in {processing_time:.2f} seconds."


    if not all_data:
        # This case means even error rows weren't generated, implies a deeper issue or no files processed
        final_message = f"Processing complete. {final_summary}\nHowever, NO data rows were generated (all files might have failed critically before data generation)."
        master_gui.after(0, app_instance.update_status, final_message)
        master_gui.after(0, app_instance.enable_buttons)
        master_gui.after(0, messagebox.showwarning, "Processing Warning", final_message, parent=master_gui)
        master_gui.after(10, app_instance.clear_pdf_preview, "Processing complete. No data.")
        return

    # --- DataFrame Creation and Saving ---
    try:
        # Ensure consistent column order, create DataFrame
        all_columns_ordered = ["Source PDF", "Invoice Folio", "Document Type", "Reference Number", "Full PDF Path"]
        df = pd.DataFrame(all_data)

        # Add any missing columns and reorder
        for col in all_columns_ordered:
            if col not in df.columns:
                df[col] = None # Add missing column filled with None
        df = df[all_columns_ordered] # Reorder to desired sequence

        # --- Enhanced Sorting ---
        try:
            # Convert Reference Number to numeric for sorting, handle errors gracefully
            # Use a temporary column to avoid data type issues in the original column
            df['Reference Number Num'] = pd.to_numeric(df['Reference Number'], errors='coerce')

            # Sort by: Filename, Document Type (e.g., Acta before Escritura), then Numeric Reference Number
            df.sort_values(
                by=["Source PDF", "Document Type", "Reference Number Num", "Reference Number"],
                inplace=True,
                na_position='last' # Put rows with non-numeric/missing refs last within their group
            )

            # Remove the temporary numeric column
            df.drop(columns=['Reference Number Num'], inplace=True)
            print("INFO: DataFrame sorted successfully.")

        except Exception as sort_e:
            print(f"Warning: Could not perform detailed sorting on DataFrame: {sort_e}. Falling back to basic sort.")
            # Fallback sort if the numeric conversion/sort fails
            df.sort_values(by=["Source PDF"], inplace=True, na_position='last')


    except Exception as e:
        error_msg = f"Error creating or sorting DataFrame: {e}"
        print(error_msg)
        traceback.print_exc()
        master_gui.after(0, app_instance.update_status, f"Error preparing data for display/saving. See console.")
        master_gui.after(0, messagebox.showerror, "DataFrame Error", error_msg, parent=master_gui)
        master_gui.after(0, app_instance.enable_buttons)
        master_gui.after(10, app_instance.clear_pdf_preview, "Error creating data table.")
        return # Stop processing if DataFrame fails

    # Find unique output filename
    output_filename = find_unique_output_filename("Extracted_Invoices.xlsx")

    # --- Attempt to Save to Excel ---
    try:
        # Select only columns intended for Excel output
        excel_columns = ["Source PDF", "Invoice Folio", "Document Type", "Reference Number"]
        df_to_save = df[excel_columns].copy() # Create a copy to avoid modifying the main df

        # Save to Excel
        df_to_save.to_excel(output_filename, index=False, engine='openpyxl')
        print(f"INFO: Data successfully saved to {output_filename}")

        # --- Update GUI on Success ---
        final_message = f"Processing complete. {final_summary}\nData saved to:\n{output_filename}"
        # Schedule GUI updates using master_gui.after to run them on the main thread
        master_gui.after(0, app_instance.update_status, final_message)
        master_gui.after(0, app_instance.set_data_and_file, df, output_filename) # Pass full df and filename
        master_gui.after(10, app_instance.load_data_to_treeview) # Load data into the GUI table
        master_gui.after(20, messagebox.showinfo, "Success", final_message, parent=master_gui)
        master_gui.after(30, app_instance.clear_pdf_preview, "Select a row above to preview PDF")

    except PermissionError:
        # Specific handling for file-in-use errors
        error_message = (f"Error saving Excel file:\n{output_filename}\n\n"
                         f"PERMISSION DENIED. The file might be open in Excel or another program.\n\n"
                         f"Data WAS processed and is shown below, but COULD NOT be saved to Excel initially.\n"
                         f"Close the file and try 'Save Changes to Excel' later.")
        print(error_message)
        master_gui.after(0, app_instance.update_status, f"Error saving file: Permission Denied. Data NOT saved initially.")
        master_gui.after(0, messagebox.showerror, "Save Error", error_message, parent=master_gui)
        # Set data in GUI but mark filename as None to indicate save failure
        master_gui.after(10, app_instance.set_data_and_file, df, None)
        master_gui.after(20, app_instance.load_data_to_treeview) # Still show data
        master_gui.after(30, app_instance.clear_pdf_preview, "Initial save failed (Permission Error).")


    except Exception as e:
        # Handle other potential saving errors
        error_message = f"An unexpected error occurred while saving the Excel file '{output_filename}':\n{e}"
        print(error_message)
        traceback.print_exc()
        master_gui.after(0, app_instance.update_status, f"Error saving file. See console. Data NOT saved initially.")
        master_gui.after(0, messagebox.showerror, "Save Error", error_message, parent=master_gui)
         # Set data in GUI but mark filename as None to indicate save failure
        master_gui.after(10, app_instance.set_data_and_file, df, None)
        master_gui.after(20, app_instance.load_data_to_treeview) # Still show data
        master_gui.after(30, app_instance.clear_pdf_preview, "Initial save failed.")


    # --- Final Step: Re-enable Buttons ---
    # Ensure buttons are re-enabled regardless of success or failure in saving
    master_gui.after(40, app_instance.enable_buttons)



