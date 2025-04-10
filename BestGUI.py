import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import re
import fitz  # PyMuPDF
import pandas as pd
import threading
import time # For status updates

# --- Pillow Import and Check ---
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("WARNING: Pillow library not found (pip install Pillow). PDF preview will be disabled.")







# folio 878 casi no implementado





# --- Regex Patterns ---
# -- Reference Patterns --
REGEX_ESCRITURA_RANGE = r'(?:Escritura|Esc)\s*(?:P[úu]blica)?\s*(?:N[uú]meros?|Nos\.?|N°)\s+(\d+)\s+A\s+(\d+)\b'
REGEX_ACTA_RANGE = r'Acta(?:s)?\s+Fuera\s+de\s+Protocolo\s+(?:N[uú]meros?|Nos\.?|N°)\s+(\d+)\s+A\s+(\d+)\b'
REGEX_ESCRITURA_LIST_Y = r'(?:Escritura|Esc)\s*(?:P[úu]blica)?\s*(?:N[uú]meros?|Nos\.?|N°)\s+(\d+)\s+Y\s+(\d+)\b'
REGEX_ACTA_LIST_Y = r'Acta(?:s)?\s+Fuera\s+de\s+Protocolo\s+(?:N[uú]meros?|Nos\.?|N°)\s+(\d+)\s+Y\s+(\d+)\b'
REGEX_ACTA_SPECIAL = r'Acta\s+Fuera\s+de\s+Protocolo\s+N[uú]mero\s+\d+\/(\d+)\/\d+\b'
REGEX_ESCRITURA_SINGLE = r'(?:Escritura|Esc)\s*(?:P[úu]blica)?\s*(?:N[uú]mero|No\.?|N°)?\s*[-:\s]*(\d+)\b'
REGEX_ACTA_SINGLE = r'Acta\s+Fuera\s+de\s+Protocolo\s+(?:N[uú]mero|No\.?|N°)?\s*[-:\s]*(\d+)\b'
# -- Folio Patterns --
REGEX_FOLIO_DBA = r'(?i)\bSerie\s*(?:RP)?\s*Folio\s*(\d+)\b'
REGEX_FOLIO_DBA_ALT = r'(?i)DATOS\s+CFDI.*?Folio:\s*(\d+)'
REGEX_FOLIO_TOTALNOT = r'(?i)Folio\s+interno:\s*(\w+)\b'


# --- Helper Functions ---
def find_unique_output_filename(base_name="Extracted_Invoices.xlsx"):
    """Checks if a file exists and appends a number if it does."""
    directory = os.getcwd()
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
            # Using TEXT_PRESERVE_WHITESPACE helps maintain layout for regex slightly better sometimes
            full_text += page.get_text("text", flags=fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE) + "\n"
        doc.close()
        # More aggressive whitespace cleaning might be needed depending on PDF
        full_text = re.sub(r'[ \t]{2,}', ' ', full_text) # Replace multiple spaces/tabs with one space
        full_text = re.sub(r'\n\s*\n', '\n', full_text) # Replace multiple newlines (with optional space) with one
        return full_text
    except Exception as e:
        print(f"Error opening or reading PDF {pdf_path}: {e}")
        return None

# --- Extraction Logic ---
def find_folio(text, invoice_type):
    """Finds the folio number based on invoice type."""
    folio = None
    if not text:
        return None

    if invoice_type == 'DBA':
        match = re.search(REGEX_FOLIO_DBA, text, re.IGNORECASE)
        if match:
            folio = match.group(1)
        else:
            match_alt = re.search(REGEX_FOLIO_DBA_ALT, text, re.IGNORECASE | re.DOTALL)
            if match_alt:
                folio = match_alt.group(1)
    elif invoice_type == 'TOTALNOT':
        match = re.search(REGEX_FOLIO_TOTALNOT, text, re.IGNORECASE)
        if match:
            folio = match.group(1)
    elif invoice_type == 'CONTPAQ':
        # Simpler CONTPAQ pattern - adjust if needed
        contpaq_simple_pattern = r'\bFOLIO:\s*(\w+)\b'
        # Look for "Folio fiscal" nearby and exclude if found
        # We iterate to find the *last* match that isn't preceded by "Folio fiscal"
        # (assuming internal folio might appear after fiscal) - adjust logic if needed
        candidate_folio = None
        for match in re.finditer(contpaq_simple_pattern, text, re.IGNORECASE):
            start_index = match.start()
            # Look back a reasonable distance for "Folio fiscal"
            lookback_chars = 30
            preceding_text = text[max(0, start_index - lookback_chars) : start_index]
            if not re.search(r'\bFolio\s+fiscal\b', preceding_text, re.IGNORECASE):
                candidate_folio = match.group(1) # Found a potential internal folio

        folio = candidate_folio # Use the last valid one found

    # Filter out potentially very long strings that are likely UUIDs (Folio Fiscal)
    if folio and len(folio) > 20:
         # print(f"Warning: Found potentially long folio '{folio}' in {invoice_type} for file. Might be Folio Fiscal. Skipping.")
         return "FOLIO_FISCAL_SUSPECTED" # Return a specific marker instead of None
    elif not folio:
        return "NOT_FOUND" # Use explicit marker for not found
    return folio

def find_references(text):
    """Finds all Escritura and Acta references, handling single, range, list, and special formats."""
    references = []
    if not text: return []
    found_escritura_numbers = set()
    found_acta_numbers = set()
    flags = re.IGNORECASE | re.UNICODE

    # Process patterns in a specific order to avoid double matching (e.g., range before single)

    # --- Ranges ---
    for match in re.finditer(REGEX_ESCRITURA_RANGE, text, flags):
        try:
            start, end = int(match.group(1)), int(match.group(2))
            if start <= end:
                for num in range(start, end + 1):
                    num_str = str(num)
                    if num_str not in found_escritura_numbers:
                        references.append({"Type": "Escritura", "Number": num_str})
                        found_escritura_numbers.add(num_str)
        except ValueError: pass # Ignore if numbers are not valid integers
    for match in re.finditer(REGEX_ACTA_RANGE, text, flags):
        try:
            start, end = int(match.group(1)), int(match.group(2))
            if start <= end:
                for num in range(start, end + 1):
                    num_str = str(num)
                    if num_str not in found_acta_numbers:
                        references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
                        found_acta_numbers.add(num_str)
        except ValueError: pass # Ignore if numbers are not valid integers

    # --- Lists ('Y') ---
    for match in re.finditer(REGEX_ESCRITURA_LIST_Y, text, flags):
        for num_str_raw in [match.group(1), match.group(2)]:
            num_str = num_str_raw.strip()
            if num_str and num_str not in found_escritura_numbers:
                references.append({"Type": "Escritura", "Number": num_str})
                found_escritura_numbers.add(num_str)
    for match in re.finditer(REGEX_ACTA_LIST_Y, text, flags):
        for num_str_raw in [match.group(1), match.group(2)]:
            num_str = num_str_raw.strip()
            if num_str and num_str not in found_acta_numbers:
                references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
                found_acta_numbers.add(num_str)

    # --- Special Acta ---
    for match in re.finditer(REGEX_ACTA_SPECIAL, text, flags):
        num_str = match.group(1).strip()
        if num_str and num_str not in found_acta_numbers:
            references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
            found_acta_numbers.add(num_str)

    # --- Singles (check context to avoid parts of ranges/lists/specials already caught) ---
    # We iterate through all potential single matches and add them only if they haven't been added via other patterns
    # This requires checking the sets `found_escritura_numbers` and `found_acta_numbers`

    potential_escritura_singles = []
    for match in re.finditer(REGEX_ESCRITURA_SINGLE, text, flags):
         num_str = match.group(1).strip()
         if num_str:
             potential_escritura_singles.append(num_str)

    potential_acta_singles = []
    for match in re.finditer(REGEX_ACTA_SINGLE, text, flags):
        num_str = match.group(1).strip()
        if num_str:
             potential_acta_singles.append(num_str)

    # Add singles only if not already found by other patterns
    for num_str in potential_escritura_singles:
        if num_str not in found_escritura_numbers:
            references.append({"Type": "Escritura", "Number": num_str})
            found_escritura_numbers.add(num_str) # Add here to prevent duplicates if number appears twice as single

    for num_str in potential_acta_singles:
        if num_str not in found_acta_numbers:
            references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
            found_acta_numbers.add(num_str) # Add here

    # --- Sorting ---
    # Sort primarily by type, then numerically by number
    def sort_key(item):
        try:
            # Attempt to convert number to integer for proper numeric sort
            num_val = int(item["Number"])
        except ValueError:
            # If conversion fails (e.g., alphanumeric), treat as high value string sort
            num_val = float('inf')
        return (item["Type"], num_val, item["Number"]) # Use original string as tertiary sort key

    references.sort(key=sort_key)
    return references


def process_single_pdf(pdf_path, invoice_type):
    """Processes a single PDF to extract folio and references. Returns full path."""
    # print(f"Processing: {os.path.basename(pdf_path)}") # Keep console log minimal
    text = extract_text_from_pdf(pdf_path)
    if not text:
        # print(f"Warning: Could not extract text from {os.path.basename(pdf_path)}. Skipping.")
        return [{"Document Type": "ERROR", "Reference Number": "Text Extraction Failed",
                 "Invoice Folio": "ERROR", "Source PDF": os.path.basename(pdf_path),
                 "Full PDF Path": pdf_path}] # Include full path even on error

    folio = find_folio(text, invoice_type)
    # if not folio: # find_folio now returns "NOT_FOUND" or "FOLIO_FISCAL_SUSPECTED"
    #     folio = "NOT_FOUND" # Use explicit marker

    references = find_references(text)
    output_rows = []

    if not references:
         # Even if no references, add a row for the folio found (or not found)
         output_rows.append({
            "Document Type": "N/A",
            "Reference Number": "N/A",
            "Invoice Folio": folio, # Will be "NOT_FOUND", "FOLIO_FISCAL_SUSPECTED", or the actual folio
            "Source PDF": os.path.basename(pdf_path), # Keep basename for display/Excel
            "Full PDF Path": pdf_path # Add full path
        })
    else:
        # If references exist, create a row for each reference, associating the same folio
        for ref in references:
            output_rows.append({
                "Document Type": ref["Type"],
                "Reference Number": ref["Number"],
                "Invoice Folio": folio, # Associate same folio with all references from this PDF
                "Source PDF": os.path.basename(pdf_path),
                "Full PDF Path": pdf_path # Add full path
            })

    return output_rows

# --- Main Processing Function (runs in a separate thread) ---
def run_processing(folder_path, invoice_type, app_instance):
    """Iterates through folder (recursively), processes PDFs, saves to Excel, and updates GUI."""
    all_data = []
    pdf_files = []

    # Find all PDF files recursively
    try:
        app_instance.master.after(0, app_instance.update_status, f"Scanning folder for PDF files: {folder_path}")
        for root, _, files in os.walk(folder_path):
            for file in files:
                # Check for hidden files/folders (like .DS_Store on Mac) - optional but good practice
                if not file.startswith('.') and file.lower().endswith('.pdf'):
                     pdf_path = os.path.join(root, file)
                     # Basic check if it's actually a file (might be symlink etc)
                     if os.path.isfile(pdf_path):
                         pdf_files.append(pdf_path) # Collect full paths
    except Exception as e:
        error_msg = f"Error accessing folder structure in {folder_path}: {e}"
        print(error_msg)
        app_instance.master.after(0, app_instance.update_status, f"Error accessing folder. Check permissions or path.")
        app_instance.master.after(0, messagebox.showerror, "Folder Error", error_msg)
        app_instance.master.after(0, app_instance.enable_buttons)
        # Ensure preview placeholder is reset
        app_instance.master.after(10, app_instance.clear_pdf_preview, "Error during folder scan.")
        return # Stop processing

    total_files = len(pdf_files)
    if total_files == 0:
        app_instance.master.after(0, app_instance.update_status, "No PDF files found in the selected folder or subfolders.")
        app_instance.master.after(0, app_instance.enable_buttons)
        app_instance.master.after(0, messagebox.showinfo, "Information", "No PDF files found.")
        # Ensure preview placeholder is reset
        app_instance.master.after(10, app_instance.clear_pdf_preview, "No PDFs found to process.")
        return

    start_time = time.time()
    files_processed_count = 0
    files_with_errors = 0 # Count files that failed during processing (text extract, critical error)

    # Process each found PDF
    for i, pdf_path in enumerate(pdf_files):
        # Update status less frequently for large numbers of files? Maybe every 10?
        if i % 5 == 0 or i == total_files - 1: # Update every 5 files and on the last one
            status_message = f"Processing file {i+1}/{total_files}: {os.path.basename(pdf_path)}"
            app_instance.master.after(0, app_instance.update_status, status_message)

        try:
            results = process_single_pdf(pdf_path, invoice_type)
            # Check if the *first* result indicates a text extraction error specifically
            if results and results[0].get("Document Type") == "ERROR" and "Extraction Failed" in results[0].get("Reference Number", ""):
                 files_with_errors += 1
                 all_data.extend(results) # Add the error entry
            elif results: # Successfully processed (even if no refs/folio found)
                 files_processed_count += 1
                 all_data.extend(results)
            else: # Should not happen if process_single_pdf guarantees a list return
                 files_with_errors += 1
                 print(f"Warning: No data returned by process_single_pdf for {os.path.basename(pdf_path)}")
                 # Add a generic error entry if process_single_pdf failed unexpectedly
                 all_data.append({
                      "Document Type": "ERROR", "Reference Number": "Processing Function Failed",
                      "Invoice Folio": "ERROR", "Source PDF": os.path.basename(pdf_path),
                      "Full PDF Path": pdf_path})

        except Exception as e:
            # Catch critical errors during the call to process_single_pdf or list extension
            files_with_errors += 1
            error_msg = f"CRITICAL ERROR processing file {os.path.basename(pdf_path)}: {e}"
            print(error_msg) # Log detailed error to console
            import traceback
            traceback.print_exc()
            # Update status briefly about the file error
            app_instance.master.after(0, app_instance.update_status, f"ERROR processing {os.path.basename(pdf_path)}. See console.")
            # Add error entry to data to potentially show in excel/treeview
            all_data.append({
                 "Document Type": "ERROR", "Reference Number": f"Runtime Error: {e}",
                 "Invoice Folio": "ERROR", "Source PDF": os.path.basename(pdf_path),
                 "Full PDF Path": pdf_path})

    end_time = time.time()
    processing_time = end_time - start_time
    final_summary = f"{files_processed_count}/{total_files} files processed"
    if files_with_errors > 0:
        final_summary += f" ({files_with_errors} file(s) encountered errors during processing)"
    final_summary += f" in {processing_time:.2f}s."

    if not all_data:
        final_message = f"Processing complete. {final_summary}\nNo data extracted."
        app_instance.master.after(0, app_instance.update_status, final_message)
        app_instance.master.after(0, app_instance.enable_buttons)
        app_instance.master.after(0, messagebox.showinfo, "Information", final_message)
        # Reset preview placeholder
        app_instance.master.after(10, app_instance.clear_pdf_preview, "Processing complete. No data.")
        return

    # Create DataFrame - including the Full PDF Path
    try:
        df = pd.DataFrame(all_data)
        # Define desired columns for Excel output (exclude Full PDF Path)
        excel_columns = ["Source PDF", "Invoice Folio", "Document Type", "Reference Number"]
        # Define order for DataFrame (keep Full PDF Path for Treeview use)
        all_columns_ordered = ["Source PDF", "Invoice Folio", "Document Type", "Reference Number", "Full PDF Path"]
        # Ensure all expected columns exist, fill missing ones if necessary
        for col in all_columns_ordered:
            if col not in df.columns:
                df[col] = None # Or appropriate default like "N/A" or ""
        # Reorder/select columns for the DataFrame we'll use internally and pass to treeview
        df = df[all_columns_ordered]

        # Sort DataFrame before saving/displaying
        try:
            # Convert Ref Number to numeric for sorting, errors='coerce' handles non-numbers
            df['Reference Number Num'] = pd.to_numeric(df['Reference Number'], errors='coerce')
            # Sort by PDF name, then Type, then numeric Ref Number, then original Ref Number string (for ties/non-numeric)
            df.sort_values(by=["Source PDF", "Document Type", "Reference Number Num", "Reference Number"],
                           inplace=True,
                           na_position='last') # Put NaNs/Nones last in sorting
            df.drop(columns=['Reference Number Num'], inplace=True) # Remove temporary sort column
        except Exception as sort_e:
            print(f"Warning: Could not perform detailed sort on DataFrame: {sort_e}")
            # Fallback sort if complex sort fails
            df.sort_values(by=["Source PDF"], inplace=True, na_position='last')

    except Exception as e:
        error_msg = f"Error creating or sorting DataFrame: {e}"
        print(error_msg)
        import traceback
        traceback.print_exc()
        app_instance.master.after(0, app_instance.update_status, f"Error preparing data. See console.")
        app_instance.master.after(0, messagebox.showerror, "DataFrame Error", error_msg)
        app_instance.master.after(0, app_instance.enable_buttons)
        # Reset preview placeholder
        app_instance.master.after(10, app_instance.clear_pdf_preview, "Error creating data.")
        return

    # Generate unique output filename
    output_filename = find_unique_output_filename("Extracted_Invoices.xlsx")

    # Save to Excel (only the selected columns)
    try:
        # Create a DataFrame slice with only the columns for Excel
        df_to_save = df[excel_columns].copy()
        # Handle potential errors like file being open
        df_to_save.to_excel(output_filename, index=False, engine='openpyxl')
        final_message = f"Processing complete. {final_summary}\nData saved to:\n{output_filename}"

        # Schedule GUI updates on the main thread
        app_instance.master.after(0, app_instance.update_status, final_message)
        app_instance.master.after(0, app_instance.load_data_to_treeview, df) # Pass the full DataFrame
        app_instance.master.after(0, messagebox.showinfo, "Success", final_message)
        # Reset preview placeholder AFTER data loaded to treeview
        app_instance.master.after(10, app_instance.clear_pdf_preview, "Select a row above to preview PDF")

    except PermissionError:
        error_message = f"Error saving Excel file:\n{output_filename}\n\nPermission denied. The file might be open in another application.\nPlease close it and try processing again."
        print(error_message)
        app_instance.master.after(0, app_instance.update_status, f"Error saving file: Permission Denied. Data NOT saved.")
        app_instance.master.after(0, messagebox.showerror, "Save Error", error_message)
        # Load data to treeview even if saving failed, so user can see results
        app_instance.master.after(10, app_instance.load_data_to_treeview, df)
        app_instance.master.after(20, app_instance.clear_pdf_preview, "Save failed. Select row to preview.")
    except Exception as e:
        error_message = f"Error saving Excel file '{output_filename}':\n{e}"
        print(error_message)
        import traceback
        traceback.print_exc()
        app_instance.master.after(0, app_instance.update_status, f"Error saving file. See console. Data NOT saved.")
        app_instance.master.after(0, messagebox.showerror, "Save Error", error_message)
        # Load data to treeview even if saving failed
        app_instance.master.after(10, app_instance.load_data_to_treeview, df)
        app_instance.master.after(20, app_instance.clear_pdf_preview, "Save failed. Select row to preview.")


    # Always re-enable buttons, schedule after other updates
    app_instance.master.after(30, app_instance.enable_buttons) # Small delay after potential messages/updates




# --- GUI Class ---
class InvoiceProcessorApp:
    def __init__(self, master):
        self.master = master
        master.title("Invoice PDF Processor v2.1 - Zoom Added")
        master.geometry("1400x850")

        self.selected_folder = tk.StringVar(value="No folder selected")
        self.processing_active = False
        self.dataframe = None
        self.pdf_preview_image = None
        self._canvas_image_id = None
        self._pdf_path_map = {}
        self._placeholder_window_id = None # To manage the placeholder window item

        # --- Zoom State ---
        self.current_zoom_factor = 1.0 # Start at 100%
        self.zoom_step = 1.2 # How much to zoom in/out each click
        self.min_zoom = 0.1 # Minimum zoom factor
        self.max_zoom = 5.0 # Maximum zoom factor
        self.current_preview_pdf_path = None # Path of the PDF currently shown

        # --- Configure Styles ---
        style = ttk.Style(root)
        style.theme_use('clam')
        style.configure('TButton', padding=(10, 5), font=('Segoe UI', 10))
        style.map('TButton', background=[('active', '#e0e0e0')])
        style.configure('Zoom.TButton', padding=(5, 2), font=('Segoe UI', 9)) # Smaller zoom buttons
        style.configure('TLabel', padding=5, font=('Segoe UI', 10))
        style.configure('Header.TLabel', font=('Segoe UI', 11, 'bold'))
        style.configure('TEntry', padding=5, font=('Segoe UI', 10))
        style.configure("Treeview.Heading", font=('Segoe UI', 10, 'bold'), padding=5)
        style.configure("Treeview", rowheight=25, font=('Segoe UI', 9))
        style.configure("Placeholder.TLabel", foreground="grey", background="lightgrey", padding=10, anchor=tk.CENTER, font=('Segoe UI', 11, 'italic'))

        # --- Top Bar: Folder Selection & Processing Controls ---
        top_controls_frame = ttk.Frame(master, padding="10 10 10 10")
        top_controls_frame.pack(side=tk.TOP, fill=tk.X)
        # (Folder selection and Process buttons layout remains the same)
        folder_frame = ttk.Frame(top_controls_frame)
        folder_frame.pack(side=tk.LEFT, padx=(0, 20))
        ttk.Label(folder_frame, text="Invoice Folder:", style='Header.TLabel').pack(anchor=tk.W)
        folder_entry_frame = ttk.Frame(folder_frame)
        folder_entry_frame.pack(fill=tk.X)
        self.folder_entry = ttk.Entry(folder_entry_frame, textvariable=self.selected_folder, state="readonly", width=60)
        self.folder_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))
        self.select_button = ttk.Button(folder_entry_frame, text="Select...", command=self.select_folder)
        self.select_button.pack(side=tk.LEFT)
        ttk.Separator(top_controls_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=20)
        process_frame = ttk.Frame(top_controls_frame)
        process_frame.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(process_frame, text="Process Invoices:", style='Header.TLabel').pack(anchor=tk.W, pady=(0,5))
        button_frame = ttk.Frame(process_frame)
        button_frame.pack()
        self.dba_button = ttk.Button(button_frame, text="DBA", command=lambda: self.start_processing('DBA'), width=12)
        self.dba_button.pack(side=tk.LEFT, padx=5, ipady=2)
        self.totalnot_button = ttk.Button(button_frame, text="TOTALNOT", command=lambda: self.start_processing('TOTALNOT'), width=12)
        self.totalnot_button.pack(side=tk.LEFT, padx=5, ipady=2)
        self.contpaq_button = ttk.Button(button_frame, text="CONTPAQ", command=lambda: self.start_processing('CONTPAQ'), width=12)
        self.contpaq_button.pack(side=tk.LEFT, padx=5, ipady=2)

        # --- Main Content Area: Paned Window ---
        self.content_pane = ttk.PanedWindow(master, orient=tk.HORIZONTAL)
        self.content_pane.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(5, 5))

        # --- Left Panel: Excel Data Preview ---
        # (Treeview layout remains the same)
        tree_frame = ttk.Frame(self.content_pane, padding=5)
        self.content_pane.add(tree_frame, weight=2)
        ttk.Label(tree_frame, text="Extracted Data:", style='Header.TLabel').pack(anchor=tk.NW, pady=(0, 5))
        tree_scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        self.tree = ttk.Treeview(tree_frame,
                                 columns=("Source PDF", "Invoice Folio", "Document Type", "Reference Number"),
                                 show='headings', yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set,
                                 selectmode='browse')
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)
        self.tree.heading("Source PDF", text="Source PDF", command=lambda: self.sort_treeview_column("Source PDF", False))
        self.tree.heading("Invoice Folio", text="Invoice Folio", command=lambda: self.sort_treeview_column("Invoice Folio", False))
        self.tree.heading("Document Type", text="Document Type", command=lambda: self.sort_treeview_column("Document Type", False))
        self.tree.heading("Reference Number", text="Reference Number", command=lambda: self.sort_treeview_column("Reference Number", False))
        self.tree.column("Source PDF", anchor=tk.W, width=220, stretch=tk.NO)
        self.tree.column("Invoice Folio", anchor=tk.W, width=100, stretch=tk.NO)
        self.tree.column("Document Type", anchor=tk.W, width=150, stretch=tk.NO)
        self.tree.column("Reference Number", anchor=tk.W, width=120, stretch=tk.NO)
        self.tree.bind('<<TreeviewSelect>>', self.on_treeview_select)
        self._tree_sort_order = {col: False for col in self.tree["columns"]}

        # --- Right Panel: PDF Preview ---
        pdf_preview_frame = ttk.Frame(self.content_pane, padding=5)
        self.content_pane.add(pdf_preview_frame, weight=3)

        # PDF Header and Zoom Controls Frame
        pdf_header_frame = ttk.Frame(pdf_preview_frame)
        pdf_header_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(pdf_header_frame, text="PDF Preview (First Page):", style='Header.TLabel').pack(side=tk.LEFT, anchor=tk.W)

        zoom_controls_frame = ttk.Frame(pdf_header_frame)
        zoom_controls_frame.pack(side=tk.RIGHT)

        self.zoom_out_button = ttk.Button(zoom_controls_frame, text="Zoom Out (-)", command=self.zoom_out, style='Zoom.TButton', width=12)
        self.zoom_out_button.pack(side=tk.LEFT, padx=2)
        self.reset_zoom_button = ttk.Button(zoom_controls_frame, text="Reset Zoom", command=self.reset_zoom, style='Zoom.TButton', width=10)
        self.reset_zoom_button.pack(side=tk.LEFT, padx=2)
        self.zoom_in_button = ttk.Button(zoom_controls_frame, text="Zoom In (+)", command=self.zoom_in, style='Zoom.TButton', width=12)
        self.zoom_in_button.pack(side=tk.LEFT, padx=2)

        # Frame to hold canvas and its scrollbars
        canvas_frame = ttk.Frame(pdf_preview_frame, relief=tk.SUNKEN, borderwidth=1)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        pdf_scroll_y = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL)
        pdf_scroll_x = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL)
        self.pdf_canvas = tk.Canvas(canvas_frame, bg="lightgrey",
                                    yscrollcommand=pdf_scroll_y.set,
                                    xscrollcommand=pdf_scroll_x.set)
        pdf_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        pdf_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.pdf_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        pdf_scroll_y.config(command=self.pdf_canvas.yview)
        pdf_scroll_x.config(command=self.pdf_canvas.xview)

        # Placeholder label (managed by clear/update)
        self.pdf_placeholder_label = ttk.Label(self.pdf_canvas, text="Select a row above to preview PDF", style="Placeholder.TLabel")
        # We bind Configure to recenter the placeholder *if* it exists
        self.pdf_canvas.bind('<Configure>', self._center_placeholder)

        # --- Bottom Bar: Status Log ---
        # (Status log layout remains the same)
        status_frame = ttk.Frame(master, padding="10 5 10 10")
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Label(status_frame, text="Status Log:", style='Header.TLabel').pack(anchor=tk.W)
        text_scroll_frame = ttk.Frame(status_frame)
        text_scroll_frame.pack(fill=tk.X, expand=False, pady=(5,0))
        scrollbar_status = ttk.Scrollbar(text_scroll_frame)
        scrollbar_status.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text = tk.Text(text_scroll_frame, height=6, wrap=tk.WORD, state=tk.DISABLED,
                                   relief=tk.FLAT, borderwidth=0,
                                   yscrollcommand=scrollbar_status.set,
                                   font=("Consolas", 9), background="#f0f0f0")
        self.status_text.pack(side=tk.LEFT, fill=tk.X, expand=True)
        scrollbar_status.config(command=self.status_text.yview)

        # --- Initial Setup ---
        self.clear_pdf_preview("Select a row above to preview PDF") # Show initial placeholder
        self.update_status("Ready. Please select a folder and invoice type.")
        if not PIL_AVAILABLE:
             self.update_status("WARNING: Pillow library not found. PDF Preview is disabled.")
             self.clear_pdf_preview("PDF Preview disabled:\nPillow library not installed.")

    # --- GUI Methods ---

    def _center_placeholder(self, event=None):
        """Centers the placeholder label *if* its window item exists."""
        # Check if the placeholder's canvas window item still exists
        if self._placeholder_window_id and self.pdf_canvas.winfo_exists() and \
           self._placeholder_window_id in self.pdf_canvas.find_withtag("placeholder"):
            canvas_w = self.pdf_canvas.winfo_width()
            canvas_h = self.pdf_canvas.winfo_height()
            # Re-center the existing window item
            self.pdf_canvas.coords(self._placeholder_window_id, canvas_w//2, canvas_h//2)
            # No need to use itemconfigure or recreate unless size changes drastically

    def select_folder(self):
        if self.processing_active: return
        folder = filedialog.askdirectory()
        if folder:
            normalized_folder = os.path.normpath(folder)
            self.selected_folder.set(normalized_folder)
            self.clear_status()
            self.clear_treeview()
            self.clear_pdf_preview("Select a row to preview PDF")
            # Reset zoom when folder changes
            self.current_zoom_factor = 1.0
            self.current_preview_pdf_path = None
            self.update_status(f"Folder selected: {normalized_folder}")
            self.update_status("Ready to process.")
        else:
            if self.selected_folder.get() != "No folder selected":
                self.update_status("Folder selection cancelled.")

    # (clear_status, update_status, _update_status_text, disable_buttons, enable_buttons methods remain the same)
    def clear_status(self):
         self.status_text.config(state=tk.NORMAL)
         self.status_text.delete('1.0', tk.END)
         self.status_text.config(state=tk.DISABLED)

    def update_status(self, message):
        if threading.current_thread() != threading.main_thread():
            self.master.after(0, self._update_status_text, message)
        else:
            self._update_status_text(message)

    def _update_status_text(self, message):
        """Internal method to update the status text widget safely."""
        current_state = self.status_text.cget('state')
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state=current_state)

    def disable_buttons(self):
        self.processing_active = True
        self.select_button.config(state=tk.DISABLED)
        self.dba_button.config(state=tk.DISABLED)
        self.totalnot_button.config(state=tk.DISABLED)
        self.contpaq_button.config(state=tk.DISABLED)
        # Optionally disable zoom buttons during processing too
        self.zoom_in_button.config(state=tk.DISABLED)
        self.zoom_out_button.config(state=tk.DISABLED)
        self.reset_zoom_button.config(state=tk.DISABLED)


    def enable_buttons(self):
        self.processing_active = False
        self.select_button.config(state=tk.NORMAL)
        self.dba_button.config(state=tk.NORMAL)
        self.totalnot_button.config(state=tk.NORMAL)
        self.contpaq_button.config(state=tk.NORMAL)
        # Enable zoom buttons
        self.zoom_in_button.config(state=tk.NORMAL)
        self.zoom_out_button.config(state=tk.NORMAL)
        self.reset_zoom_button.config(state=tk.NORMAL)

    def start_processing(self, invoice_type):
        folder = self.selected_folder.get()
        if not folder or folder == "No folder selected":
            messagebox.showerror("Error", "Please select a folder first.", parent=self.master)
            return
        if not os.path.isdir(folder):
             messagebox.showerror("Error", f"Invalid directory selected:\n{folder}", parent=self.master)
             return
        if self.processing_active:
            messagebox.showwarning("Busy", "Processing is already in progress.", parent=self.master)
            return

        self.disable_buttons()
        self.clear_status()
        self.clear_treeview()
        self.clear_pdf_preview(f"Processing {invoice_type} invoices...\nPlease wait.")
        self.update_status(f"Starting recursive processing for {invoice_type} in: {folder}")
        self.update_status("-" * 40)

        process_thread = threading.Thread(target=run_processing,
                                          args=(folder, invoice_type, self),
                                          daemon=True)
        process_thread.start()

    # (clear_treeview, load_data_to_treeview, sort_treeview_column methods remain the same)
    def clear_treeview(self):
        """Clears all items from the Treeview and the path map."""
        if hasattr(self, 'tree'):
            for item in self.tree.get_children():
                self.tree.delete(item)
        self.dataframe = None
        self._pdf_path_map.clear()

    def load_data_to_treeview(self, df):
        """Loads data from the DataFrame into the Treeview."""
        self.clear_treeview()
        if df is None or df.empty:
             self.update_status("Processing finished. No data extracted or displayable.")
             self.clear_pdf_preview("No data loaded. Select folder and process.")
             return

        self.dataframe = df
        display_columns = list(self.tree["columns"])
        required_cols = display_columns + ["Full PDF Path"]
        if not all(col in df.columns for col in required_cols):
            missing = [col for col in required_cols if col not in df.columns]
            errmsg = f"Error: DataFrame is missing required columns: {missing}"
            self.update_status(errmsg)
            messagebox.showerror("Data Loading Error", errmsg, parent=self.master)
            print(f"DataFrame columns available: {df.columns.tolist()}")
            self.clear_pdf_preview("Error loading data. See status log.")
            return

        self.tree.configure(displaycolumns=display_columns)
        for index, row in df.iterrows():
            try:
                display_values = tuple(str(row[col]) if pd.notna(row[col]) else "" for col in display_columns)
                full_path = row["Full PDF Path"]
                item_id = self.tree.insert("", tk.END, values=display_values)
                if full_path and isinstance(full_path, str):
                    self._pdf_path_map[item_id] = full_path
            except Exception as e:
                print(f"Error adding row index {index} to treeview: {e}")
                self.update_status(f"Warning: Could not display row for PDF '{row.get('Source PDF', 'Unknown')}' in table.")

        row_count = len(self.tree.get_children())
        self.update_status(f"Loaded {row_count} rows into the table.")
        if row_count == 0:
             self.update_status("Loaded data frame, but no rows were added to the table (check data validity).")
             self.clear_pdf_preview("Data loaded, but no rows to display.")

    def sort_treeview_column(self, col, reverse):
        """Sorts the treeview column based on the clicked header."""
        if self.dataframe is None or self.dataframe.empty: return
        try:
            sorted_df = self.dataframe.copy()
            try:
                numeric_col = pd.to_numeric(sorted_df[col], errors='coerce')
                is_numeric = not numeric_col.isna().all()
            except Exception: is_numeric = False

            if is_numeric:
                 sorted_df = sorted_df.iloc[numeric_col.argsort(na_position='last')]
                 if reverse: sorted_df = sorted_df.iloc[::-1]
            else:
                 sorted_df = sorted_df.sort_values(by=col, ascending=not reverse, key=lambda x: x.astype(str).str.lower(), na_position='last')

            for c in self.tree["columns"]:
                current_text = self.tree.heading(c, "text").replace(' ▲', '').replace(' ▼', '')
                if c == col:
                    indicator = ' ▲' if not reverse else ' ▼'
                    self.tree.heading(c, text=current_text + indicator)
                else:
                    self.tree.heading(c, text=current_text)

            self.load_data_to_treeview(sorted_df)
            self._tree_sort_order[col] = not reverse

        except KeyError:
             print(f"Error: Column '{col}' not found in DataFrame for sorting.")
             messagebox.showerror("Sort Error", f"Column '{col}' not found.", parent=self.master)
        except Exception as e:
             print(f"Error sorting treeview column '{col}': {e}")
             import traceback; traceback.print_exc()
             messagebox.showerror("Sort Error", f"Could not sort column '{col}':\n{e}", parent=self.master)


    def on_treeview_select(self, event):
        """Handles selection changes in the Treeview to update the PDF preview."""
        selected_items = self.tree.selection()
        if not selected_items: return

        selected_item_id = selected_items[0]
        pdf_full_path = self._pdf_path_map.get(selected_item_id)

        if pdf_full_path and isinstance(pdf_full_path, str):
            if os.path.exists(pdf_full_path):
                # Store the path for zoom controls before calling update
                # self.current_preview_pdf_path = pdf_full_path # update_pdf_preview does this now
                self.update_pdf_preview(pdf_full_path)
            else:
                 base_name = os.path.basename(pdf_full_path)
                 err_msg = f"File Not Found:\n{base_name}\n(Path: {pdf_full_path})"
                 self.clear_pdf_preview(err_msg)
                 print(f"Error: File path from selection '{pdf_full_path}' does not exist.")
                 self.update_status(f"Preview Error: File not found - {base_name}")
                 self.current_preview_pdf_path = None # Clear path if file not found
        elif pdf_full_path is None:
            print(f"Error: No path found in map for selected item ID: {selected_item_id}")
            self.clear_pdf_preview("Error: Could not retrieve\nfile path for this row.")
            self.current_preview_pdf_path = None
        else:
             print(f"Error: Invalid path data associated with selected row: {pdf_full_path}")
             self.clear_pdf_preview("Error: Invalid file path\nin selected row data.")
             self.current_preview_pdf_path = None


    # --- PDF Preview and Zoom Methods ---

    def clear_pdf_preview(self, message="Select a row to preview PDF"):
        """Clears the PDF preview area, shows placeholder, and resets path."""
        # Delete previous image item
        if self._canvas_image_id:
             try: self.pdf_canvas.delete(self._canvas_image_id)
             except tk.TclError: pass
             self._canvas_image_id = None
        self.pdf_preview_image = None # Clear PhotoImage reference

        # Delete existing placeholder window item *if* it exists
        if self._placeholder_window_id:
            try: self.pdf_canvas.delete("placeholder") # Delete by tag
            except tk.TclError: pass
            self._placeholder_window_id = None

        # Recreate the placeholder label and its window item
        if hasattr(self, 'pdf_placeholder_label') and self.pdf_placeholder_label.winfo_exists():
             self.pdf_placeholder_label.config(text=message) # Update text
        else: # Label got destroyed somehow, recreate it
             self.pdf_placeholder_label = ttk.Label(self.pdf_canvas, text=message, style="Placeholder.TLabel")

        # Create the window item for the label, centered
        canvas_w = self.pdf_canvas.winfo_width() if self.pdf_canvas.winfo_width() > 1 else 200
        canvas_h = self.pdf_canvas.winfo_height() if self.pdf_canvas.winfo_height() > 1 else 200
        self._placeholder_window_id = self.pdf_canvas.create_window(
            canvas_w//2, canvas_h//2,
            window=self.pdf_placeholder_label,
            anchor=tk.CENTER,
            tags="placeholder" # Add tag for easy deletion later
        )

        # Reset scroll region and current path
        self.master.after_idle(self._update_canvas_scrollregion)
        self.current_preview_pdf_path = None


    def update_pdf_preview(self, pdf_path):
        """Renders PDF page using current zoom factor and displays it."""
        if not PIL_AVAILABLE:
             self.clear_pdf_preview("PDF Preview disabled:\nPillow library not found.")
             return

        # --- Start: Clear existing content ---
        # Delete existing placeholder window item FIRST
        if self._placeholder_window_id:
            try: self.pdf_canvas.delete("placeholder")
            except tk.TclError: pass
            self._placeholder_window_id = None

        # Delete previous canvas image item
        if self._canvas_image_id:
            try: self.pdf_canvas.delete(self._canvas_image_id)
            except tk.TclError: pass
            self._canvas_image_id = None
        self.pdf_preview_image = None # Clear PhotoImage reference
        # --- End: Clear existing content ---

        try:
            doc = fitz.open(pdf_path)
            if len(doc) == 0:
                raise ValueError("PDF has no pages.") # Treat as error for preview

            page = doc.load_page(0)
            page_rect = page.rect
            if page_rect.width == 0 or page_rect.height == 0:
                 raise ValueError("PDF page has zero dimensions.")

            # --- Use the current zoom factor ---
            zoom_factor = self.current_zoom_factor
            mat = fitz.Matrix(zoom_factor, zoom_factor)
            pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB, alpha=False)
            doc.close()

            img_bytes = pix.samples
            if not isinstance(img_bytes, bytes): img_bytes = bytes(img_bytes)
            if not img_bytes: raise ValueError("Pixmap samples are empty.")

            pil_image = Image.frombytes("RGB", [pix.width, pix.height], img_bytes)
            self.pdf_preview_image = ImageTk.PhotoImage(image=pil_image)

            if self.pdf_preview_image:
                # Display image
                self._canvas_image_id = self.pdf_canvas.create_image(0, 0, anchor=tk.NW, image=self.pdf_preview_image, tags="pdf_image")
                # Store the path of the successfully loaded PDF
                self.current_preview_pdf_path = pdf_path
                # Update scroll region and reset view
                self.master.after_idle(self._update_canvas_scrollregion)
                self.pdf_canvas.xview_moveto(0)
                self.pdf_canvas.yview_moveto(0)
            else:
                 raise ValueError("Failed to create PhotoImage object.")

        except (fitz.fitz.FileNotFoundError, fitz.fitz.PasswordError, ValueError, Exception) as e:
            # Catch known and general errors during PDF processing/rendering
            base_name = os.path.basename(pdf_path) if pdf_path else "Unknown File"
            error_type = type(e).__name__
            error_msg = f"Preview Error ({error_type}):\n{base_name}"
            if isinstance(e, fitz.fitz.PasswordError): error_msg += "\n(Password Protected?)"
            elif isinstance(e, ValueError): error_msg += f"\n({e})" # Show specific value error msg
            # Log detailed error to console
            import traceback
            print(f"--- PDF Preview Exception Traceback for {base_name} ---")
            traceback.print_exc()
            print(f"--- End Traceback ---")
            # Display simplified error in preview area and clear state
            self.clear_pdf_preview(error_msg) # Also resets current_preview_pdf_path


    def zoom_in(self):
        """Increases zoom factor and re-renders the current PDF."""
        if self.current_preview_pdf_path:
            new_zoom = min(self.current_zoom_factor * self.zoom_step, self.max_zoom)
            if new_zoom != self.current_zoom_factor:
                self.current_zoom_factor = new_zoom
                # print(f"Zoom In: Factor = {self.current_zoom_factor:.2f}") # Debug
                self.update_pdf_preview(self.current_preview_pdf_path)

    def zoom_out(self):
        """Decreases zoom factor and re-renders the current PDF."""
        if self.current_preview_pdf_path:
            new_zoom = max(self.current_zoom_factor / self.zoom_step, self.min_zoom)
            if new_zoom != self.current_zoom_factor:
                self.current_zoom_factor = new_zoom
                # print(f"Zoom Out: Factor = {self.current_zoom_factor:.2f}") # Debug
                self.update_pdf_preview(self.current_preview_pdf_path)

    def reset_zoom(self):
        """Resets zoom factor to 1.0 and re-renders the current PDF."""
        if self.current_preview_pdf_path and self.current_zoom_factor != 1.0:
            self.current_zoom_factor = 1.0
            # print(f"Reset Zoom: Factor = {self.current_zoom_factor:.2f}") # Debug
            self.update_pdf_preview(self.current_preview_pdf_path)

    def _update_canvas_scrollregion(self):
        """ Safely update canvas scroll region based on the pdf image bounds. """
        try:
            # Check if the image item exists
            if self._canvas_image_id and self.pdf_canvas.find_withtag(self._canvas_image_id):
                bbox = self.pdf_canvas.bbox(self._canvas_image_id)
                if bbox:
                    scroll_bbox = (bbox[0], bbox[1], bbox[2] + 5, bbox[3] + 5)
                    self.pdf_canvas.config(scrollregion=scroll_bbox)
                    return
            # Fallback: No image or bbox failed, fit to canvas widget size
            current_width = self.pdf_canvas.winfo_width()
            current_height = self.pdf_canvas.winfo_height()
            self.pdf_canvas.config(scrollregion=(0, 0, max(1, current_width), max(1, current_height)))
        except tk.TclError as e:
             print(f"Warning: TclError updating canvas scrollregion: {e}")
        except Exception as e:
             print(f"Error updating canvas scrollregion: {e}")


# --- Main Execution ---
if __name__ == "__main__":
    if not PIL_AVAILABLE:
         root_check = tk.Tk(); root_check.withdraw()
         messagebox.showwarning("Dependency Warning",
                                "Python Imaging Library (Pillow) not found.\n"
                                "PDF Preview functionality will be disabled.\n\n"
                                "Please install it using pip:\n"
                                "pip install Pillow", parent=None)
         root_check.destroy()

    root = tk.Tk()
    app = InvoiceProcessorApp(root)
    root.mainloop()



