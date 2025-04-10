import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import re
import fitz  # PyMuPDF
import pandas as pd
import threading
import time # For status updates

# --- Regex Patterns ---

# -- Reference Patterns --
# Prioritize multi-number formats first
REGEX_ESCRITURA_RANGE = r'(?:Escritura|Esc)\s*(?:P[úu]blica)?\s*(?:N[uú]meros?|Nos\.?|N°)\s+(\d+)\s+A\s+(\d+)\b'
REGEX_ACTA_RANGE = r'Acta(?:s)?\s+Fuera\s+de\s+Protocolo\s+(?:N[uú]meros?|Nos\.?|N°)\s+(\d+)\s+A\s+(\d+)\b'
REGEX_ESCRITURA_LIST_Y = r'(?:Escritura|Esc)\s*(?:P[úu]blica)?\s*(?:N[uú]meros?|Nos\.?|N°)\s+(\d+)\s+Y\s+(\d+)\b'
REGEX_ACTA_LIST_Y = r'Acta(?:s)?\s+Fuera\s+de\s+Protocolo\s+(?:N[uú]meros?|Nos\.?|N°)\s+(\d+)\s+Y\s+(\d+)\b'
# Add variations if needed e.g. with commas: r'(\d+)\s*,\s*(\d+)\s+Y\s+(\d+)' etc.

# Special Acta Format
REGEX_ACTA_SPECIAL = r'Acta\s+Fuera\s+de\s+Protocolo\s+N[uú]mero\s+\d+\/(\d+)\/\d+\b'

# Single References (Use *after* checking multi-patterns)
REGEX_ESCRITURA_SINGLE = r'(?:Escritura|Esc)\s*(?:P[úu]blica)?\s*(?:N[uú]mero|No\.?|N°)?\s*[-:\s]*(\d+)\b'
REGEX_ACTA_SINGLE = r'Acta\s+Fuera\s+de\s+Protocolo\s+(?:N[uú]mero|No\.?|N°)?\s*[-:\s]*(\d+)\b'


# -- Folio Patterns --
# Type-specific Folio patterns (prioritize these)
REGEX_FOLIO_DBA = r'(?i)\bSerie\s*(?:RP)?\s*Folio\s*(\d+)\b'
REGEX_FOLIO_DBA_ALT = r'(?i)DATOS\s+CFDI.*?Folio:\s*(\d+)' # Look for Folio within 'DATOS CFDI' section as fallback for DBA
REGEX_FOLIO_TOTALNOT = r'(?i)Folio\s+interno:\s*(\w+)\b' # \w allows numbers and letters like R00857
# REGEX_FOLIO_CONTPAQ handled in function find_folio directly


# --- Helper Functions ---

def find_unique_output_filename(base_name="Extracted_Invoices.xlsx"):
    """Checks if a file exists and appends a number if it does."""
    directory = os.getcwd() # Save in the current working directory, or specify another
    output_path = os.path.join(directory, base_name)
    counter = 1
    name, ext = os.path.splitext(base_name)
    while os.path.exists(output_path):
        output_path = os.path.join(directory, f"{name}_{counter}{ext}")
        counter += 1
    return output_path

def extract_text_from_pdf(pdf_path):
    """Extracts all text from a PDF file using PyMuPDF."""
    try:
        doc = fitz.open(pdf_path)
        full_text = ""
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            # Extract text, preserving layout slightly better for pattern matching
            full_text += page.get_text("text", flags=fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE) + "\n"
        doc.close()
        # Simple cleanup: replace multiple spaces/newlines which might break regex context
        full_text = re.sub(r'\s{2,}', ' ', full_text)
        full_text = re.sub(r'\n+', '\n', full_text)
        return full_text
    except Exception as e:
        print(f"Error opening or reading PDF {pdf_path}: {e}")
        return None

# --- Extraction Logic ---

# --- find_folio function (no changes needed based on request) ---
def find_folio(text, invoice_type):
    """Finds the folio number based on invoice type."""
    folio = None
    if not text: # Add check for empty text
        return None

    if invoice_type == 'DBA':
        match = re.search(REGEX_FOLIO_DBA, text, re.IGNORECASE)
        if match:
            folio = match.group(1)
        else:
            # Use DOTALL for REGEX_FOLIO_DBA_ALT as Folio might be lines away from DATOS CFDI
            match_alt = re.search(REGEX_FOLIO_DBA_ALT, text, re.IGNORECASE | re.DOTALL)
            if match_alt:
                folio = match_alt.group(1)
    elif invoice_type == 'TOTALNOT':
        match = re.search(REGEX_FOLIO_TOTALNOT, text, re.IGNORECASE)
        if match:
            folio = match.group(1)
    elif invoice_type == 'CONTPAQ':
        # Use a simpler pattern first, then check context
        # This pattern finds "FOLIO:" followed by alphanumeric characters
        contpaq_simple_pattern = r'\bFOLIO:\s*(\w+)\b'
        # Find all potential matches using finditer to get match objects (with positions)
        for match in re.finditer(contpaq_simple_pattern, text, re.IGNORECASE):
            # Get the start position of the potential match
            start_index = match.start()
            # Define how many characters to look back (adjust if needed)
            lookback_chars = 25
            # Extract the text immediately preceding the match
            preceding_text = text[max(0, start_index - lookback_chars) : start_index]
            # Check if "Folio fiscal" (case-insensitive) is in the preceding text
            # Use a more precise check to avoid partial matches like 'Folio fi...'
            if not re.search(r'\bFolio\s+fiscal\b', preceding_text, re.IGNORECASE):
                # If "Folio fiscal" is NOT found before it, this is likely the correct Folio
                folio = match.group(1)
                break # Found the first valid one, stop searching

    # Basic check to avoid unreasonably long 'folios'
    if folio and len(folio) > 20:
         print(f"Warning: Found potentially long folio '{folio}' in {invoice_type}. Might be Folio Fiscal. Skipping.")
         return None

    return folio


# --- Updated find_references function ---
def find_references(text):
    """Finds all Escritura and Acta references, handling single, range, list, and special formats."""
    references = []
    if not text:
        return []

    # Use sets to keep track of numbers already found to avoid duplicates
    found_escritura_numbers = set()
    found_acta_numbers = set()

    # Flags: Ignore Case, Multiline (though less critical with cleaned text)
    flags = re.IGNORECASE

    # --- 1. Process Ranges ("A") ---
    # Escritura Ranges
    for match in re.finditer(REGEX_ESCRITURA_RANGE, text, flags):
        try:
            start_num = int(match.group(1))
            end_num = int(match.group(2))
            if start_num <= end_num:
                print(f"DEBUG: Found Escritura Range: {start_num} A {end_num}")
                for num in range(start_num, end_num + 1):
                    num_str = str(num)
                    if num_str not in found_escritura_numbers:
                        references.append({"Type": "Escritura", "Number": num_str})
                        found_escritura_numbers.add(num_str)
            else:
                 print(f"Warning: Invalid range found {start_num} A {end_num}")
        except ValueError:
            print(f"Warning: Could not parse range numbers in match: {match.groups()}")

    # Acta Ranges
    for match in re.finditer(REGEX_ACTA_RANGE, text, flags):
        try:
            start_num = int(match.group(1))
            end_num = int(match.group(2))
            if start_num <= end_num:
                print(f"DEBUG: Found Acta Range: {start_num} A {end_num}")
                for num in range(start_num, end_num + 1):
                    num_str = str(num)
                    if num_str not in found_acta_numbers:
                        references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
                        found_acta_numbers.add(num_str)
            else:
                 print(f"Warning: Invalid range found {start_num} A {end_num}")
        except ValueError:
            print(f"Warning: Could not parse range numbers in match: {match.groups()}")

    # --- 2. Process Lists ("Y") ---
    # Escritura Lists
    for match in re.finditer(REGEX_ESCRITURA_LIST_Y, text, flags):
        num1_str = match.group(1).strip()
        num2_str = match.group(2).strip()
        print(f"DEBUG: Found Escritura List: {num1_str} Y {num2_str}")
        if num1_str not in found_escritura_numbers:
            references.append({"Type": "Escritura", "Number": num1_str})
            found_escritura_numbers.add(num1_str)
        if num2_str not in found_escritura_numbers:
            references.append({"Type": "Escritura", "Number": num2_str})
            found_escritura_numbers.add(num2_str)

    # Acta Lists
    for match in re.finditer(REGEX_ACTA_LIST_Y, text, flags):
        num1_str = match.group(1).strip()
        num2_str = match.group(2).strip()
        print(f"DEBUG: Found Acta List: {num1_str} Y {num2_str}")
        if num1_str not in found_acta_numbers:
            references.append({"Type": "Acta Fuera de Protocolo", "Number": num1_str})
            found_acta_numbers.add(num1_str)
        if num2_str not in found_acta_numbers:
            references.append({"Type": "Acta Fuera de Protocolo", "Number": num2_str})
            found_acta_numbers.add(num2_str)

    # --- 3. Process Special Acta Format ---
    for match in re.finditer(REGEX_ACTA_SPECIAL, text, flags):
        num_str = match.group(1).strip()
        print(f"DEBUG: Found Special Acta: .../{num_str}/...")
        if num_str not in found_acta_numbers:
            references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
            found_acta_numbers.add(num_str)

    # --- 4. Process Single References (Check against found sets) ---
    # Single Escrituras
    for match in re.finditer(REGEX_ESCRITURA_SINGLE, text, flags):
        num_str = match.group(1).strip()
        # Check if this number was already captured by range/list logic
        if num_str not in found_escritura_numbers:
             # Heuristic check: Ensure this match isn't immediately preceded/followed by keywords
             # that suggest it was part of a multi-match that the specific regex missed.
             start_pos = match.start(0) # Start of the whole match e.g. "Escritura Publica Numero 123"
             end_pos = match.end(0)     # End of the whole match
             context_window_before = text[max(0, start_pos - 10):start_pos]
             context_window_after = text[end_pos:min(len(text), end_pos + 10)]

             # Avoid if it looks like NUMEROS X Y [number] or [number] A Y
             is_part_of_list_y = re.search(r'(?:N[uú]meros?|Nos\.?|N°)\s+\d+\s+Y\s*$', context_window_before, flags) or \
                                 re.search(r'^\s*Y\s+\d+', context_window_after, flags)
             # Avoid if it looks like NUMEROS X A [number] or [number] A X
             is_part_of_range_a = re.search(r'(?:N[uú]meros?|Nos\.?|N°)\s+\d+\s+A\s*$', context_window_before, flags) or \
                                  re.search(r'^\s*A\s+\d+', context_window_after, flags)

             if not is_part_of_list_y and not is_part_of_range_a:
                 print(f"DEBUG: Found Single Escritura: {num_str}")
                 references.append({"Type": "Escritura", "Number": num_str})
                 found_escritura_numbers.add(num_str)
             else:
                # This number might have been part of a slightly misformatted range/list.
                # It's safer to skip adding it again if already found.
                # (The `if num_str not in found_escritura_numbers:` already handles the primary case)
                print(f"DEBUG: Skipping single Escritura {num_str} - already found or context suggests multi-match.")


    # Single Actas
    for match in re.finditer(REGEX_ACTA_SINGLE, text, flags):
        num_str = match.group(1).strip()
        if num_str not in found_acta_numbers:
             start_pos = match.start(0)
             end_pos = match.end(0)
             context_window_before = text[max(0, start_pos - 10):start_pos]
             context_window_after = text[end_pos:min(len(text), end_pos + 10)]

             # Avoid special format like xxx/[number]/yyy
             is_part_of_special = re.search(r'\d+\/\s*$', context_window_before) or \
                                  re.search(r'^\/\d+', context_window_after)
             # Avoid lists Y
             is_part_of_list_y = re.search(r'(?:N[uú]meros?|Nos\.?|N°)\s+\d+\s+Y\s*$', context_window_before, flags) or \
                                 re.search(r'^\s*Y\s+\d+', context_window_after, flags)
             # Avoid ranges A
             is_part_of_range_a = re.search(r'(?:N[uú]meros?|Nos\.?|N°)\s+\d+\s+A\s*$', context_window_before, flags) or \
                                  re.search(r'^\s*A\s+\d+', context_window_after, flags)


             if not is_part_of_special and not is_part_of_list_y and not is_part_of_range_a:
                 print(f"DEBUG: Found Single Acta: {num_str}")
                 references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
                 found_acta_numbers.add(num_str)
             else:
                 print(f"DEBUG: Skipping single Acta {num_str} - already found or context suggests multi/special match.")

    # Sort results for consistency (optional, but helpful)
    # Convert to int for sorting where possible, handle non-digits gracefully
    def sort_key(item):
        try:
            return (item["Type"], int(item["Number"]))
        except ValueError:
            return (item["Type"], float('inf')) # Put non-numeric last

    references.sort(key=sort_key)

    return references


def process_single_pdf(pdf_path, invoice_type):
    """Processes a single PDF to extract folio and references."""
    print(f"Processing: {os.path.basename(pdf_path)}")
    text = extract_text_from_pdf(pdf_path)
    if not text:
        print(f"Warning: Could not extract text from {os.path.basename(pdf_path)}. Skipping.")
        return [] # Skip if text extraction failed

    # Add DEBUG print for extracted text (optional, can be noisy)
    # print(f"--- Text from {os.path.basename(pdf_path)} ---")
    # print(text[:1000]) # Print first 1000 chars
    # print("--- End Text ---")


    folio = find_folio(text, invoice_type)
    if not folio:
        print(f"Warning: Folio number not found or skipped in {os.path.basename(pdf_path)}")
        # Decide how to handle missing folio: "NOT_FOUND", None, or ""
        # Using "NOT_FOUND" makes it explicit in the output.
        folio = "NOT_FOUND"

    references = find_references(text) # Use the updated function

    output_rows = []
    if not references:
        # If no references found, add a row indicating the folio was processed
        # but no references were found for that specific folio.
         output_rows.append({
            "Document Type": "N/A",
            "Reference Number": "N/A",
            "Invoice Folio": folio,
            "Source PDF": os.path.basename(pdf_path) # Add source file for traceability
        })
         print(f"Info: No Escritura or Acta found in {os.path.basename(pdf_path)}")
    else:
        for ref in references:
            output_rows.append({
                "Document Type": ref["Type"],
                "Reference Number": ref["Number"],
                "Invoice Folio": folio,
                "Source PDF": os.path.basename(pdf_path) # Add source file
            })
        print(f"Info: Found {len(references)} reference(s) in {os.path.basename(pdf_path)}")


    return output_rows

# --- Main Processing Function (runs in a separate thread) ---

def run_processing(folder_path, invoice_type, app_instance):
    """Iterates through folder, processes PDFs, and saves to Excel."""
    all_data = []
    pdf_files = []

    # Find all PDF files first to calculate total
    try:
        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith('.pdf'):
                    pdf_files.append(os.path.join(root, file))
    except Exception as e:
        error_msg = f"Error walking through directory {folder_path}: {e}"
        print(error_msg)
        app_instance.master.after(0, app_instance.update_status, f"Error accessing folder. Check permissions.")
        app_instance.master.after(0, messagebox.showerror, "Folder Error", error_msg)
        app_instance.master.after(0, app_instance.enable_buttons)
        return


    total_files = len(pdf_files)
    if total_files == 0:
        # Ensure GUI update happens on the main thread
        app_instance.master.after(0, app_instance.update_status, "No PDF files found in the selected folder.")
        app_instance.master.after(0, app_instance.enable_buttons)
        app_instance.master.after(0, messagebox.showinfo, "Information", "No PDF files were found in the selected directory or its subdirectories.")
        return

    start_time = time.time()
    files_processed_count = 0
    files_with_errors = 0

    for i, pdf_path in enumerate(pdf_files):
        # Update status on the main thread
        status_message = f"Processing file {i+1}/{total_files}: {os.path.basename(pdf_path)}"
        # Use 'after' to schedule the GUI update on the main thread
        app_instance.master.after(0, app_instance.update_status, status_message)

        try:
            results = process_single_pdf(pdf_path, invoice_type)
            if results: # Only extend if process_single_pdf didn't return None or [] due to error
                 all_data.extend(results)
            files_processed_count += 1
        except Exception as e:
            files_with_errors += 1
            error_msg = f"CRITICAL ERROR processing file {os.path.basename(pdf_path)}: {e}"
            print(error_msg)
            # Update status about the error, but continue processing others
            app_instance.master.after(0, app_instance.update_status, f"ERROR processing {os.path.basename(pdf_path)}. See console/log.")
            # Optionally log the full error traceback here


    end_time = time.time()
    processing_time = end_time - start_time
    final_summary = f"{files_processed_count}/{total_files} files processed"
    if files_with_errors > 0:
        final_summary += f" ({files_with_errors} with errors)"


    if not all_data:
        final_message = f"Processing complete ({final_summary} in {processing_time:.2f}s).\nNo relevant data extracted or saved."
        app_instance.master.after(0, app_instance.update_status, final_message)
        app_instance.master.after(0, app_instance.enable_buttons)
        app_instance.master.after(0, messagebox.showinfo, "Information", final_message)
        return

    # Create DataFrame
    try:
        df = pd.DataFrame(all_data)
        # Reorder columns
        df = df[["Source PDF", "Invoice Folio", "Document Type", "Reference Number"]]
        # Optional: Sort DataFrame for better readability
        df.sort_values(by=["Source PDF", "Invoice Folio", "Document Type", "Reference Number"], inplace=True)

    except Exception as e:
        error_msg = f"Error creating DataFrame: {e}"
        print(error_msg)
        app_instance.master.after(0, app_instance.update_status, f"Error creating DataFrame. See console.")
        app_instance.master.after(0, messagebox.showerror, "DataFrame Error", error_msg)
        app_instance.master.after(0, app_instance.enable_buttons)
        return


    # Generate unique output filename
    output_filename = find_unique_output_filename("Extracted_Invoices.xlsx")

    # Save to Excel
    try:
        df.to_excel(output_filename, index=False, engine='openpyxl')
        final_message = f"Processing complete ({final_summary} in {processing_time:.2f}s).\nData saved to:\n{output_filename}"
        app_instance.master.after(0, app_instance.update_status, final_message)
        app_instance.master.after(0, messagebox.showinfo, "Success", final_message)
    except PermissionError:
        error_message = f"Error saving Excel file '{output_filename}':\nPermission denied. The file might be open in another application.\nPlease close it and try again."
        print(error_message)
        app_instance.master.after(0, app_instance.update_status, f"Error saving file: Permission Denied. File likely open.")
        app_instance.master.after(0, messagebox.showerror, "Error", error_message)
    except Exception as e:
        error_message = f"Error saving Excel file '{output_filename}':\n{e}"
        print(error_message)
        app_instance.master.after(0, app_instance.update_status, f"Error saving file. See console.")
        app_instance.master.after(0, messagebox.showerror, "Error", error_message)

    # Re-enable buttons on the main thread
    app_instance.master.after(0, app_instance.enable_buttons)


# --- GUI Class ---

class InvoiceProcessorApp:
    def __init__(self, master):
        self.master = master
        master.title("Invoice PDF Processor v1.1") # Added version
        master.geometry("600x400") # Slightly larger for better status display

        self.selected_folder = tk.StringVar()
        self.selected_folder.set("No folder selected")
        self.processing_active = False

        # Configure styles
        style = ttk.Style(root)
        style.theme_use('clam') # Or 'alt', 'default', 'classic', 'vista', 'xpnative'
        style.configure('TButton', padding=6, relief="flat", background="#ccc")
        style.map('TButton', background=[('active', '#eee')])
        style.configure('TLabel', padding=5)
        style.configure('TEntry', padding=5)


        # Folder Selection Frame
        folder_frame = ttk.Frame(master, padding="10 10 10 5")
        folder_frame.pack(fill=tk.X)

        ttk.Label(folder_frame, text="Invoice Folder:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(folder_frame, textvariable=self.selected_folder, state="readonly", width=60).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5) # Increased width
        self.select_button = ttk.Button(folder_frame, text="Select Folder...", command=self.select_folder)
        self.select_button.pack(side=tk.LEFT)

        # Invoice Type and Processing Frame
        process_frame = ttk.Frame(master, padding="10 5 10 5")
        process_frame.pack(fill=tk.X)

        ttk.Label(process_frame, text="Select Invoice Type and Process:").pack(pady=(0, 10), anchor=tk.W)

        # Buttons for each type in a sub-frame for centering/layout
        button_frame = ttk.Frame(process_frame)
        button_frame.pack(pady=5) # Add some vertical padding

        self.dba_button = ttk.Button(button_frame, text="Process DBA", command=lambda: self.start_processing('DBA'), width=15) # Set width
        self.dba_button.pack(side=tk.LEFT, padx=10, pady=5, ipady=4) # Add internal padding

        self.totalnot_button = ttk.Button(button_frame, text="Process TOTALNOT", command=lambda: self.start_processing('TOTALNOT'), width=15)
        self.totalnot_button.pack(side=tk.LEFT, padx=10, pady=5, ipady=4)

        self.contpaq_button = ttk.Button(button_frame, text="Process CONTPAQ", command=lambda: self.start_processing('CONTPAQ'), width=15)
        self.contpaq_button.pack(side=tk.LEFT, padx=10, pady=5, ipady=4)

        # Status Area Frame
        status_frame = ttk.Frame(master, padding="10 5 10 10")
        status_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(status_frame, text="Status Log:").pack(anchor=tk.W)
        # Use a Text widget for multiline status, make it read-only
        # Encapsulate Text and Scrollbar in their own frame for better packing
        text_scroll_frame = ttk.Frame(status_frame)
        text_scroll_frame.pack(fill=tk.BOTH, expand=True, pady=(5,0))

        scrollbar = ttk.Scrollbar(text_scroll_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.status_text = tk.Text(text_scroll_frame, height=10, wrap=tk.WORD, state=tk.DISABLED,
                                   relief=tk.SUNKEN, borderwidth=1, yscrollcommand=scrollbar.set,
                                   font=("Consolas", 9)) # Use a fixed-width font?
        self.status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar.config(command=self.status_text.yview)

        self.update_status("Ready. Please select a folder and invoice type.")


    def select_folder(self):
        if self.processing_active: return # Don't allow changes during processing
        folder = filedialog.askdirectory()
        if folder:
            # Normalize path separators for consistency
            normalized_folder = os.path.normpath(folder)
            self.selected_folder.set(normalized_folder)
            self.clear_status() # Clear status on new folder selection
            self.update_status(f"Folder selected: {normalized_folder}\nReady to process.")
        else:
             # Only update if it wasn't already "No folder"
             if self.selected_folder.get() != "No folder selected":
                 self.selected_folder.set("No folder selected")
                 self.update_status("Folder selection cancelled.")

    def clear_status(self):
         self.status_text.config(state=tk.NORMAL)
         self.status_text.delete('1.0', tk.END)
         self.status_text.config(state=tk.DISABLED)

    def update_status(self, message):
        # Ensure this runs on the main thread if called from another thread
        # Use lambda to ensure the message is captured correctly at call time
        if threading.current_thread() != threading.main_thread():
            self.master.after(0, lambda msg=message: self._update_status_text(msg))
        else:
            self._update_status_text(message)

    def _update_status_text(self, message):
        """Internal method to update the text widget, always runs on main thread."""
        current_state = self.status_text.cget('state')
        self.status_text.config(state=tk.NORMAL) # Enable writing
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END) # Scroll to the bottom
        self.status_text.config(state=current_state) # Restore previous state (usually DISABLED)
        # self.master.update_idletasks() # Force update (usually not needed with 'after')

    def disable_buttons(self):
        self.processing_active = True
        self.select_button.config(state=tk.DISABLED)
        self.dba_button.config(state=tk.DISABLED)
        self.totalnot_button.config(state=tk.DISABLED)
        self.contpaq_button.config(state=tk.DISABLED)

    def enable_buttons(self):
        self.processing_active = False
        self.select_button.config(state=tk.NORMAL)
        self.dba_button.config(state=tk.NORMAL)
        self.totalnot_button.config(state=tk.NORMAL)
        self.contpaq_button.config(state=tk.NORMAL)

    def start_processing(self, invoice_type):
        folder = self.selected_folder.get()
        if not folder or folder == "No folder selected":
            messagebox.showerror("Error", "Please select a folder containing the invoices first.")
            return

        if not os.path.isdir(folder):
             messagebox.showerror("Error", f"The selected path is not a valid directory:\n{folder}")
             return

        if self.processing_active:
            messagebox.showwarning("Busy", "Processing is already in progress.")
            return

        self.disable_buttons()
        self.clear_status() # Clear status before starting new process
        self.update_status(f"Starting processing for {invoice_type} invoices in:\n{folder}")
        self.update_status("-" * 30) # Separator

        # Run the processing in a separate thread to avoid freezing the GUI
        process_thread = threading.Thread(target=run_processing,
                                          args=(folder, invoice_type, self),
                                          daemon=True) # Daemon thread exits when main app exits
        process_thread.start()


# --- Main Execution ---
if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceProcessorApp(root)
    root.mainloop()