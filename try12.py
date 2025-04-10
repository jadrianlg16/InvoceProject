import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import re
import fitz  # PyMuPDF
import pandas as pd
import threading
import time # For status updates
import copy # For deep copying clipboard data

# --- Pillow Import and Check ---
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("WARNING: Pillow library not found (pip install Pillow). PDF preview will be disabled.")


'''
This version
has basic editing and saving to the excel, sorting, has stable functionality, etc

Folio 1272 Honorarios Esctrctura no implementada
Folio 1264 Honorarios estructura no implementada
Folio 1262 Honorarios estructura no implementada
Folio 1238 Honorarios estructura no implementada
Folio 1264 Honorarios estructura no implementada
Folio 1180 Honorarios estructura no implementada
Folio 1189 Honorarios estructura no implementada
Folio 1188 Honorarios estructura no implementada
Folio 1069 Honorarios estructura no implementada
Folio 1067 Honorarios estructura no implementada
Folio 1063 Honorarios estructura no implementada
Folio 1238 Honorarios estructura no implementada
Folio 1019 Honorarios estructura no implementada
Folio 1011 Honorarios estructura no implementada
Folio 936 Honorarios estructura no implementada
Folio 921 Honorarios estructura no implementada
Folio 904 Honorarios estructura no implementada
Folio 979 Honorarios estructura no implementada
Folio 952 Honorarios estructura no implementada
Folio 950 Honorarios estructura no implementada
Folio 966 Honorarios estructura no implementada
Folio 962 Honorarios estructura no implementada
Folio 839 Honorarios estructura no implementada

**Added Feature**: Remembers PDF preview scroll position between selections.
'''


# --- Regex Patterns ---
REGEX_ESCRITURA_RANGE = r'(?:Escritura|Esc)\s*(?:P[úu]blica)?\s*(?:N[uú]meros?|Nos.?|N°)\s+(\d+)\s+A\s+(\d+)\b'
REGEX_ACTA_RANGE = r'Acta(?:s)?\s+Fuera\s+de\s+Protocolo\s+(?:N[uú]meros?|Nos.?|N°)\s+(\d+)\s+A\s+(\d+)\b'
REGEX_ACTA_RANGE_AL = r'Acta(?:s)?\s+Fuera\s+de\s+Protocolo\s+N[uú]meros?\s+(\d+)\s+AL\s+(\d+)\b' # Handles "NUMEROS start AL end"
REGEX_ESCRITURA_LIST_Y = r'(?:Escritura|Esc)\s*(?:P[úu]blica)?\s*(?:N[uú]meros?|Nos.?|N°)\s+(\d+)\s+Y\s+(\d+)\b'
REGEX_ACTA_LIST_Y = r'Acta(?:s)?\s+Fuera\s+de\s+Protocolo\s+(?:N[uú]meros?|Nos.?|N°)\s+(\d+)\s+Y\s+(\d+)\b'
REGEX_ACTA_SPECIAL = r'Acta\s+Fuera\s+de\s+Protocolo\s+N[uú]mero\s+\d+/(\d+)/\d+\b'
REGEX_ESCRITURA_SINGLE = r'(?:Escritura|Esc)\s*(?:P[úu]blica)?\s*(?:N[uú]mero|No.?|N°)?\s*[-:\s]?\s*(\d+)\b(?!\s+A\s+\d+)'
REGEX_ACTA_SINGLE = r'Acta\s+Fuera\s+de\s+Protocolo\s+(?:N[uú]mero|No.?|N°)?\s*[-:\s]?\s*(\d+)\b(?!\s+(?:A|AL)\s+\d+)'
REGEX_FOLIO_DBA = r'(?i)\bSerie\s(?:RP)?\sFolio\s(\d+)\b'
REGEX_FOLIO_DBA_ALT = r'(?i)DATOS\s+CFDI.?Folio:\s(\d+)'
REGEX_FOLIO_TOTALNOT = r'(?i)Folio\s+interno:\s*(\w+)\b'

# --- Helper Functions ---
def find_unique_output_filename(base_name="Extracted_Invoices.xlsx"):
    directory = os.getcwd()
    output_path = os.path.join(directory, base_name)
    counter = 1
    name, ext = os.path.splitext(base_name)
    while os.path.exists(output_path):
        output_path = os.path.join(directory, f"{name}_{counter}{ext}")
        counter += 1
    return output_path

def extract_text_from_pdf(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        full_text = ""
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            full_text += page.get_text("text", flags=fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE) + "\n"
        doc.close()
        full_text = re.sub(r'[ \t]{2,}', ' ', full_text)
        full_text = re.sub(r'\n\s*\n', '\n', full_text)
        return full_text
    except Exception as e:
        print(f"Error opening or reading PDF {pdf_path}: {e}")
        return None

# --- Extraction Logic ---
def find_folio(text, invoice_type):
    folio = None
    if not text: return None
    if invoice_type == 'DBA':
        match = re.search(REGEX_FOLIO_DBA, text, re.IGNORECASE)
        if match: folio = match.group(1)
        else:
            match_alt = re.search(REGEX_FOLIO_DBA_ALT, text, re.IGNORECASE | re.DOTALL)
            if match_alt: folio = match_alt.group(1)
    elif invoice_type == 'TOTALNOT':
        match = re.search(REGEX_FOLIO_TOTALNOT, text, re.IGNORECASE)
        if match: folio = match.group(1)
    elif invoice_type == 'CONTPAQ':
        contpaq_simple_pattern = r'\bFOLIO:\s*(\w+)\b'
        candidate_folio = None
        for match in re.finditer(contpaq_simple_pattern, text, re.IGNORECASE):
            start_index = match.start()
            lookback_chars = 30
            preceding_text = text[max(0, start_index - lookback_chars) : start_index]
            if not re.search(r'\bFolio\s+fiscal\b', preceding_text, re.IGNORECASE):
                candidate_folio = match.group(1)
        folio = candidate_folio
    if folio and len(folio) > 20: return "FOLIO_FISCAL_SUSPECTED"
    elif not folio: return "NOT_FOUND"
    return folio

def find_references(text):
    """
    Extracts Escritura and Acta Fuera de Protocolo references from text.
    Handles single numbers, ranges ("A", "AL"), lists ("Y"), and special formats.
    Uses refined single regex with lookaheads to avoid conflicts with ranges.
    """
    references = []
    if not text:
        return []

    found_escritura_numbers = set()
    found_acta_numbers = set()
    flags = re.IGNORECASE | re.UNICODE

    # --- STEP 1: Process Ranges, Lists, and Special Formats FIRST ---

    # Escritura Range ("A")
    for match in re.finditer(REGEX_ESCRITURA_RANGE, text, flags):
        try:
            start, end = int(match.group(1)), int(match.group(2))
            if start <= end:
                for num in range(start, end + 1):
                    num_str = str(num)
                    if num_str not in found_escritura_numbers:
                        references.append({"Type": "Escritura", "Number": num_str})
                        found_escritura_numbers.add(num_str)
        except ValueError:
            print(f"Warning: Could not parse Escritura 'A' range numbers: {match.groups()}")
            pass

    # Acta Range ("A")
    for match in re.finditer(REGEX_ACTA_RANGE, text, flags):
        try:
            start, end = int(match.group(1)), int(match.group(2))
            if start <= end:
                for num in range(start, end + 1):
                    num_str = str(num)
                    if num_str not in found_acta_numbers:
                        references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
                        found_acta_numbers.add(num_str)
        except ValueError:
            print(f"Warning: Could not parse Acta 'A' range numbers: {match.groups()}")
            pass

    # Acta Range ("AL") - The new pattern
    for match in re.finditer(REGEX_ACTA_RANGE_AL, text, flags):
        try:
            start, end = int(match.group(1)), int(match.group(2))
            if start <= end:
                for num in range(start, end + 1):
                    num_str = str(num)
                    if num_str not in found_acta_numbers:
                        references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
                        found_acta_numbers.add(num_str)
        except ValueError:
            print(f"Warning: Could not parse Acta 'AL' range numbers: {match.groups()}")
            pass

    # Escritura List ("Y")
    for match in re.finditer(REGEX_ESCRITURA_LIST_Y, text, flags):
        for num_str_raw in [match.group(1), match.group(2)]:
            num_str = num_str_raw.strip()
            if num_str and num_str not in found_escritura_numbers:
                references.append({"Type": "Escritura", "Number": num_str})
                found_escritura_numbers.add(num_str)

    # Acta List ("Y")
    for match in re.finditer(REGEX_ACTA_LIST_Y, text, flags):
        for num_str_raw in [match.group(1), match.group(2)]:
            num_str = num_str_raw.strip()
            if num_str and num_str not in found_acta_numbers:
                references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
                found_acta_numbers.add(num_str)

    # Special Acta Format
    for match in re.finditer(REGEX_ACTA_SPECIAL, text, flags):
        num_str = match.group(1).strip()
        if num_str and num_str not in found_acta_numbers:
            references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
            found_acta_numbers.add(num_str)

    # --- STEP 2: Process Singles LAST (using the refined regex) ---

    # Potential Escritura Singles
    potential_escritura_singles = [m.group(1).strip() for m in re.finditer(REGEX_ESCRITURA_SINGLE, text, flags) if m.group(1)]
    for num_str in potential_escritura_singles:
        # Check if it wasn't already added by a range/list
        if num_str not in found_escritura_numbers:
            references.append({"Type": "Escritura", "Number": num_str})
            found_escritura_numbers.add(num_str)

    # Potential Acta Singles
    potential_acta_singles = [m.group(1).strip() for m in re.finditer(REGEX_ACTA_SINGLE, text, flags) if m.group(1)]
    for num_str in potential_acta_singles:
        # Check if it wasn't already added by a range/list/special
        if num_str not in found_acta_numbers:
            references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
            found_acta_numbers.add(num_str)

    # --- STEP 3: Sorting ---
    def sort_key(item):
        try:
            num_val = int(item["Number"])
        except ValueError:
            num_val = float('inf') # Place non-numeric numbers last within their type
        return (item["Type"], num_val, item["Number"]) # Sort by Type, then Number (numeric), then Number (string)

    references.sort(key=sort_key)

    return references

def process_single_pdf(pdf_path, invoice_type):
    text = extract_text_from_pdf(pdf_path)
    if not text:
        return [{"Document Type": "ERROR", "Reference Number": "Text Extraction Failed",
                 "Invoice Folio": "ERROR", "Source PDF": os.path.basename(pdf_path),
                 "Full PDF Path": pdf_path}]
    folio = find_folio(text, invoice_type)
    references = find_references(text)
    output_rows = []
    if not references:
         # If no references, still add a row with the folio and PDF info
         output_rows.append({
            "Document Type": "N/A", "Reference Number": "N/A",
            "Invoice Folio": folio, "Source PDF": os.path.basename(pdf_path),
            "Full PDF Path": pdf_path })
    else:
        for ref in references:
            output_rows.append({
                "Document Type": ref["Type"], "Reference Number": ref["Number"],
                "Invoice Folio": folio, "Source PDF": os.path.basename(pdf_path),
                "Full PDF Path": pdf_path })
    return output_rows

# --- Main Processing Function (runs in a separate thread) ---
def run_processing(folder_path, invoice_type, app_instance):
    all_data = []
    pdf_files = []
    output_filename = None # Initialize here
    try:
        # Inform user about scanning
        app_instance.master.after(0, app_instance.update_status, f"Scanning folder for PDF files: {folder_path}")
        # Recursively find all PDF files
        for root, _, files in os.walk(folder_path):
            for file in files:
                # Ignore hidden files/folders and non-PDF files
                if not file.startswith('.') and file.lower().endswith('.pdf'):
                     pdf_path = os.path.join(root, file)
                     # Double-check it's actually a file (not a broken link, etc.)
                     if os.path.isfile(pdf_path):
                         pdf_files.append(pdf_path)
    except Exception as e:
        # Handle errors during folder traversal (permissions, etc.)
        error_msg = f"Error accessing folder structure in {folder_path}: {e}"
        print(error_msg)
        app_instance.master.after(0, app_instance.update_status, f"Error accessing folder. Check permissions or path.")
        app_instance.master.after(0, messagebox.showerror, "Folder Error", error_msg)
        app_instance.master.after(0, app_instance.enable_buttons) # Re-enable GUI
        app_instance.master.after(10, app_instance.clear_pdf_preview, "Error during folder scan.") # Update preview area
        return # Stop processing

    total_files = len(pdf_files)
    if total_files == 0:
        # Inform user if no PDFs were found
        app_instance.master.after(0, app_instance.update_status, "No PDF files found in the selected folder or subfolders.")
        app_instance.master.after(0, app_instance.enable_buttons)
        app_instance.master.after(0, messagebox.showinfo, "Information", "No PDF files found.")
        app_instance.master.after(10, app_instance.clear_pdf_preview, "No PDFs found to process.")
        return

    # Process each found PDF file
    start_time = time.time()
    files_processed_count = 0
    files_with_errors = 0
    for i, pdf_path in enumerate(pdf_files):
        # Update status periodically
        if i % 5 == 0 or i == total_files - 1: # Update every 5 files and for the last one
            status_message = f"Processing file {i+1}/{total_files}: {os.path.basename(pdf_path)}"
            app_instance.master.after(0, app_instance.update_status, status_message)

        try:
            # Process the individual PDF
            results = process_single_pdf(pdf_path, invoice_type)
            # Check for specific errors returned by process_single_pdf
            if results and results[0].get("Document Type") == "ERROR" and "Extraction Failed" in results[0].get("Reference Number", ""):
                 files_with_errors += 1
                 all_data.extend(results) # Add the error row
            elif results: # Successful processing (even if no references found)
                 files_processed_count += 1
                 all_data.extend(results)
            else:
                 # Handle unexpected case where process_single_pdf returns nothing
                 files_with_errors += 1
                 print(f"Warning: No data returned by process_single_pdf for {os.path.basename(pdf_path)}")
                 # Add a generic error row for this file
                 all_data.append({
                     "Document Type": "ERROR", "Reference Number": "Processing Function Failed",
                     "Invoice Folio": "ERROR", "Source PDF": os.path.basename(pdf_path),
                     "Full PDF Path": pdf_path
                 })
        except Exception as e:
            # Catch unexpected errors during process_single_pdf call itself
            files_with_errors += 1
            error_msg = f"CRITICAL ERROR processing file {os.path.basename(pdf_path)}: {e}"
            print(error_msg)
            import traceback; traceback.print_exc() # Print full traceback to console
            app_instance.master.after(0, app_instance.update_status, f"ERROR processing {os.path.basename(pdf_path)}. See console.")
            # Add an error row indicating the runtime error
            all_data.append({
                "Document Type": "ERROR", "Reference Number": f"Runtime Error: {e}",
                "Invoice Folio": "ERROR", "Source PDF": os.path.basename(pdf_path),
                "Full PDF Path": pdf_path
            })

    # Processing finished, calculate time and summarize
    end_time = time.time()
    processing_time = end_time - start_time
    final_summary = f"{files_processed_count}/{total_files} files processed"
    if files_with_errors > 0:
        final_summary += f" ({files_with_errors} file(s) encountered errors)"
    final_summary += f" in {processing_time:.2f}s."

    # Handle case where no data was extracted at all
    if not all_data:
        final_message = f"Processing complete. {final_summary}\nNo data extracted."
        app_instance.master.after(0, app_instance.update_status, final_message)
        app_instance.master.after(0, app_instance.enable_buttons)
        app_instance.master.after(0, messagebox.showinfo, "Information", final_message)
        app_instance.master.after(10, app_instance.clear_pdf_preview, "Processing complete. No data.")
        return

    # Convert extracted data to DataFrame and prepare for display/saving
    try:
        df = pd.DataFrame(all_data)
        # Define column order for Excel and internal DataFrame
        excel_columns = ["Source PDF", "Invoice Folio", "Document Type", "Reference Number"]
        all_columns_ordered = ["Source PDF", "Invoice Folio", "Document Type", "Reference Number", "Full PDF Path"]

        # Ensure all required columns exist, add if missing (shouldn't happen with current logic, but safe)
        for col in all_columns_ordered:
            if col not in df.columns:
                df[col] = None # Add missing column filled with None

        # Reorder columns for internal use
        df = df[all_columns_ordered]

        # Initial sort (optional, but helpful for initial view)
        try:
            # Attempt a more detailed sort: by PDF name, then Doc Type, then Ref Num (numeric first)
            # Create a temporary numeric column for sorting references
            df['Reference Number Num'] = pd.to_numeric(df['Reference Number'], errors='coerce')
            df.sort_values(by=["Source PDF", "Document Type", "Reference Number Num", "Reference Number"],
                           inplace=True, na_position='last')
            # Remove the temporary numeric column
            df.drop(columns=['Reference Number Num'], inplace=True)
        except Exception as sort_e:
            # Fallback sort if the detailed sort fails
            print(f"Warning: Could not perform detailed initial sort on DataFrame: {sort_e}")
            df.sort_values(by=["Source PDF"], inplace=True, na_position='last')

    except Exception as e:
        # Handle errors during DataFrame creation or sorting
        error_msg = f"Error creating or sorting DataFrame: {e}"
        print(error_msg)
        import traceback; traceback.print_exc()
        app_instance.master.after(0, app_instance.update_status, f"Error preparing data. See console.")
        app_instance.master.after(0, messagebox.showerror, "DataFrame Error", error_msg)
        app_instance.master.after(0, app_instance.enable_buttons)
        app_instance.master.after(10, app_instance.clear_pdf_preview, "Error creating data.")
        return

    # Find a unique filename for the Excel output
    output_filename = find_unique_output_filename("Extracted_Invoices.xlsx")

    # Attempt to save the results to Excel
    try:
        # Create a copy with only the columns intended for Excel
        df_to_save = df[excel_columns].copy()
        df_to_save.to_excel(output_filename, index=False, engine='openpyxl')

        # Success: Update status, pass data to GUI, load treeview, show success message
        final_message = f"Processing complete. {final_summary}\nData saved to:\n{output_filename}"
        app_instance.master.after(0, app_instance.update_status, final_message)
        # Pass both the full DataFrame and the saved Excel path to the main app instance
        app_instance.master.after(0, app_instance.set_data_and_file, df, output_filename)
        # Trigger loading the data into the Treeview
        app_instance.master.after(10, app_instance.load_data_to_treeview)
        # Show a success popup
        app_instance.master.after(20, messagebox.showinfo, "Success", final_message)
        # Update preview area prompt
        app_instance.master.after(30, app_instance.clear_pdf_preview, "Select a row above to preview PDF")

    except PermissionError:
        # Handle file permission errors (e.g., file open in Excel)
        error_message = (f"Error saving Excel file:\n{output_filename}\n\n"
                         f"Permission denied. Is the file open in another program (like Excel)?\n\n"
                         f"Data has been processed and is shown below, but WAS NOT saved to Excel initially. "
                         f"You can try 'Save Changes to Excel' later after closing the file.")
        print(error_message)
        app_instance.master.after(0, app_instance.update_status, f"Error saving file: Permission Denied. Data NOT saved initially.")
        app_instance.master.after(0, messagebox.showerror, "Save Error", error_message)
        # Pass the DataFrame but None for the filename, disable save button initially
        app_instance.master.after(10, app_instance.set_data_and_file, df, None)
        app_instance.master.after(20, app_instance.load_data_to_treeview)
        app_instance.master.after(30, app_instance.clear_pdf_preview, "Initial save failed (Permission Error). Edits cannot be saved until resolved.")

    except Exception as e:
        # Handle other potential saving errors
        error_message = f"Error saving Excel file '{output_filename}':\n{e}"
        print(error_message)
        import traceback; traceback.print_exc()
        app_instance.master.after(0, app_instance.update_status, f"Error saving file. See console. Data NOT saved initially.")
        app_instance.master.after(0, messagebox.showerror, "Save Error", error_message)
        # Pass the DataFrame but None for the filename, disable save button initially
        app_instance.master.after(10, app_instance.set_data_and_file, df, None)
        app_instance.master.after(20, app_instance.load_data_to_treeview)
        app_instance.master.after(30, app_instance.clear_pdf_preview, "Initial save failed. Edits cannot be saved.")

    # Re-enable GUI buttons regardless of save success/failure
    app_instance.master.after(40, app_instance.enable_buttons)


# --- GUI Class ---
class InvoiceProcessorApp:
    def __init__(self, master):
        self.master = master
        master.title("Invoice PDF Processor v2.5 - Interactive Sort & Scroll Memory") # Version bump
        master.geometry("1400x850")

        self.selected_folder = tk.StringVar(value="No folder selected")
        self.processing_active = False
        self.dataframe = None
        self.current_excel_file = None # Path to the *saved* Excel file
        self.pdf_preview_image = None # Stores the PhotoImage for canvas
        self._canvas_image_id = None # ID of the image on the canvas
        self._pdf_path_map = {} # Maps tree item ID -> full PDF path
        self._df_index_map = {} # Maps tree item ID -> df index (IMPORTANT for edits/deletes)
        self._placeholder_window_id = None # ID of the placeholder text on canvas
        self._edit_entry = None # Widget for inline cell editing
        self._edit_item_id = None # Treeview item ID being edited
        self._edit_column_id = None # Treeview column ID (#n) being edited
        self._clipboard = [] # For copy/paste rows (list of dictionaries)

        # --- Zoom State ---
        self.current_zoom_factor = 1.0
        self.zoom_step = 1.2
        self.min_zoom = 0.1
        self.max_zoom = 5.0
        self.current_preview_pdf_path = None # Path of the PDF currently in preview

        # --- Scroll Position Memory --- [NEW]
        self.pdf_scroll_positions = {} # Stores {pdf_path: (x_fraction, y_fraction)}

        # --- Sorting State ---
        self._tree_current_sort_col = None  # Track the ID of the currently sorted column
        self._tree_current_sort_ascending = True # Track the direction of the current sort

        # --- Configure Styles ---
        style = ttk.Style(master)
        style.theme_use('clam') # Or 'vista', 'xpnative', etc.
        style.configure('TButton', padding=(10, 5), font=('Segoe UI', 10))
        style.map('TButton', background=[('active', '#e0e0e0')])
        style.configure('Zoom.TButton', padding=(5, 2), font=('Segoe UI', 9)) # Smaller zoom buttons
        style.configure('TLabel', padding=5, font=('Segoe UI', 10))
        style.configure('Header.TLabel', font=('Segoe UI', 11, 'bold'))
        style.configure('TEntry', padding=5, font=('Segoe UI', 10))
        style.configure("Treeview.Heading", font=('Segoe UI', 10, 'bold'), padding=5)
        style.configure("Treeview", rowheight=25, font=('Segoe UI', 9))
        # Style for the placeholder label in the PDF preview area
        style.configure("Placeholder.TLabel", foreground="grey", background="lightgrey", padding=10, anchor=tk.CENTER, font=('Segoe UI', 11, 'italic'))


        # --- Top Bar (Folder Selection & Process Buttons) ---
        top_controls_frame = ttk.Frame(master, padding="10 10 10 10")
        top_controls_frame.pack(side=tk.TOP, fill=tk.X)

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
        self.dba_button.pack(side=tk.LEFT, padx=5, ipady=2) # ipady adds vertical padding inside button
        self.totalnot_button = ttk.Button(button_frame, text="TOTALNOT", command=lambda: self.start_processing('TOTALNOT'), width=12)
        self.totalnot_button.pack(side=tk.LEFT, padx=5, ipady=2)
        self.contpaq_button = ttk.Button(button_frame, text="CONTPAQ", command=lambda: self.start_processing('CONTPAQ'), width=12)
        self.contpaq_button.pack(side=tk.LEFT, padx=5, ipady=2)

        # --- Main Content Area (Paned Window) ---
        self.content_pane = ttk.PanedWindow(master, orient=tk.HORIZONTAL)
        self.content_pane.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(5, 5))

        # --- Left Panel (Treeview and Controls) ---
        tree_frame = ttk.Frame(self.content_pane, padding=5)
        self.content_pane.add(tree_frame, weight=2) # Adjust weight as needed

        tree_header_frame = ttk.Frame(tree_frame)
        tree_header_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(tree_header_frame, text="Extracted Data (Double-click cell to edit, Right-click row for options):", style='Header.TLabel').pack(side=tk.LEFT, anchor=tk.W)
        self.save_changes_button = ttk.Button(tree_header_frame, text="Save Changes to Excel", command=self.save_changes_to_excel, state=tk.DISABLED)
        self.save_changes_button.pack(side=tk.RIGHT, padx=5)

        tree_scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)

        self.tree = ttk.Treeview(tree_frame,
                                 columns=("Source PDF", "Invoice Folio", "Document Type", "Reference Number"),
                                 show='headings', yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set,
                                 selectmode='extended') # extended allows multi-select

        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)

        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)

        # Define Treeview headings and column properties
        self.tree.heading("Source PDF", text="Source PDF", command=lambda c="Source PDF": self.sort_treeview_column(c))
        self.tree.heading("Invoice Folio", text="Invoice Folio", command=lambda c="Invoice Folio": self.sort_treeview_column(c))
        self.tree.heading("Document Type", text="Document Type", command=lambda c="Document Type": self.sort_treeview_column(c))
        self.tree.heading("Reference Number", text="Reference Number", command=lambda c="Reference Number": self.sort_treeview_column(c))

        self.tree.column("Source PDF", anchor=tk.W, width=220, stretch=tk.NO) # Prevent Source PDF from stretching too much
        self.tree.column("Invoice Folio", anchor=tk.W, width=100)
        self.tree.column("Document Type", anchor=tk.W, width=150)
        self.tree.column("Reference Number", anchor=tk.W, width=120)

        # Define which columns can be edited by double-clicking
        self.editable_columns = ["Invoice Folio", "Document Type", "Reference Number"]

        # Bind events
        self.tree.bind('<<TreeviewSelect>>', self.on_treeview_select)
        self.tree.bind('<Double-1>', self.on_tree_double_click) # Double-left-click
        self.tree.bind('<Button-3>', self.show_context_menu) # Right-click

        # --- Context Menu ---
        self.context_menu = tk.Menu(master, tearoff=0)
        self.context_menu.add_command(label="Add Blank Row", command=self._add_row)
        self.context_menu.add_command(label="Copy Selected Row(s)", command=self._copy_selected_rows)
        self.context_menu.add_command(label="Paste Row(s)", command=self._paste_rows)
        self.context_menu.add_command(label="Delete Selected Row(s)", command=self._delete_selected_rows)

        # --- Right Panel (PDF Preview) ---
        pdf_preview_frame = ttk.Frame(self.content_pane, padding=5)
        self.content_pane.add(pdf_preview_frame, weight=3) # Adjust weight

        pdf_header_frame = ttk.Frame(pdf_preview_frame)
        pdf_header_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(pdf_header_frame, text="PDF Preview (First Page):", style='Header.TLabel').pack(side=tk.LEFT, anchor=tk.W)

        # Zoom Controls
        zoom_controls_frame = ttk.Frame(pdf_header_frame)
        zoom_controls_frame.pack(side=tk.RIGHT)
        self.zoom_out_button = ttk.Button(zoom_controls_frame, text="Zoom Out (-)", command=self.zoom_out, style='Zoom.TButton', width=12)
        self.zoom_out_button.pack(side=tk.LEFT, padx=2)
        self.reset_zoom_button = ttk.Button(zoom_controls_frame, text="Reset Zoom", command=self.reset_zoom, style='Zoom.TButton', width=10)
        self.reset_zoom_button.pack(side=tk.LEFT, padx=2)
        self.zoom_in_button = ttk.Button(zoom_controls_frame, text="Zoom In (+)", command=self.zoom_in, style='Zoom.TButton', width=12)
        self.zoom_in_button.pack(side=tk.LEFT, padx=2)

        # Canvas Frame with Border
        canvas_frame = ttk.Frame(pdf_preview_frame, relief=tk.SUNKEN, borderwidth=1)
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        pdf_scroll_y = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL)
        pdf_scroll_x = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL)

        self.pdf_canvas = tk.Canvas(canvas_frame, bg="lightgrey", yscrollcommand=pdf_scroll_y.set, xscrollcommand=pdf_scroll_x.set)

        pdf_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        pdf_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.pdf_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        pdf_scroll_y.config(command=self.pdf_canvas.yview)
        pdf_scroll_x.config(command=self.pdf_canvas.xview)

        # Placeholder Label (initially hidden, shown in clear_pdf_preview)
        self.pdf_placeholder_label = ttk.Label(self.pdf_canvas, text="", style="Placeholder.TLabel")
        # Center placeholder when canvas resizes
        self.pdf_canvas.bind('<Configure>', self._center_placeholder)

        # --- Bottom Bar (Status Log) ---
        status_frame = ttk.Frame(master, padding="10 5 10 10") # Less top padding
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Label(status_frame, text="Status Log:", style='Header.TLabel').pack(anchor=tk.W)

        # Frame to contain Text and Scrollbar properly
        text_scroll_frame = ttk.Frame(status_frame)
        text_scroll_frame.pack(fill=tk.X, expand=False, pady=(5,0)) # Don't expand vertically

        scrollbar_status = ttk.Scrollbar(text_scroll_frame)
        scrollbar_status.pack(side=tk.RIGHT, fill=tk.Y)

        self.status_text = tk.Text(text_scroll_frame, height=6, wrap=tk.WORD, state=tk.DISABLED,
                                   relief=tk.FLAT, borderwidth=0, # Make it look less like an entry
                                   yscrollcommand=scrollbar_status.set,
                                   font=("Consolas", 9), background="#f0f0f0") # Monospaced font, slightly off-white bg
        self.status_text.pack(side=tk.LEFT, fill=tk.X, expand=True)
        scrollbar_status.config(command=self.status_text.yview)

        # --- Initial Setup ---
        self.clear_pdf_preview("Select a row above to preview PDF") # Set initial placeholder text
        self.update_status("Ready. Please select a folder and invoice type.")
        if not PIL_AVAILABLE:
             self.update_status("WARNING: Pillow library not found. PDF Preview is disabled.")
             self.clear_pdf_preview("PDF Preview disabled:\nPillow library not installed.")

    def _center_placeholder(self, event=None):
        """Centers the placeholder label within the PDF canvas."""
        # Check if the placeholder window exists on the canvas
        if self._placeholder_window_id and self.pdf_canvas.winfo_exists() and \
           self._placeholder_window_id in self.pdf_canvas.find_withtag("placeholder"):
            canvas_w = self.pdf_canvas.winfo_width()
            canvas_h = self.pdf_canvas.winfo_height()
            # Move the window to the center
            self.pdf_canvas.coords(self._placeholder_window_id, canvas_w//2, canvas_h//2)

    def select_folder(self):
        """Opens a dialog to select a folder and updates the entry field."""
        if self.processing_active: return # Don't allow while processing

        # Warn if changing folder with unsaved data
        if self.dataframe is not None and self.current_excel_file: # Check if data exists and was initially saved
             if not messagebox.askokcancel("Confirm Folder Change",
                                        "Changing the folder will clear the current data and edits.\n"
                                        "Any unsaved changes will be lost.\n\nProceed?", parent=self.master):
                  return # User cancelled

        folder = filedialog.askdirectory()
        if folder:
            # Normalize the path (e.g., converts '/' to '\' on Windows)
            normalized_folder = os.path.normpath(folder)
            self.selected_folder.set(normalized_folder)
            # Clear previous results and state
            self.clear_status()
            self.clear_treeview() # Also clears dataframe, excel path, clipboard
            self.pdf_scroll_positions.clear() # [NEW] Clear scroll positions
            self.clear_pdf_preview("Select a row to preview PDF") # Reset preview
            self.current_zoom_factor = 1.0 # Reset zoom
            self.current_preview_pdf_path = None
            self.update_status(f"Folder selected: {normalized_folder}")
            self.update_status("Ready to process.")
            self.save_changes_button.config(state=tk.DISABLED) # Disable save until new data is loaded/saved
            # Reset sort state as well
            self._tree_current_sort_col = None
            self._tree_current_sort_ascending = True
        else:
            # User cancelled the dialog, only log if a folder was previously selected
            if self.selected_folder.get() != "No folder selected":
                self.update_status("Folder selection cancelled.")

    def clear_status(self):
         """Clears the status log text area."""
         self.status_text.config(state=tk.NORMAL)
         self.status_text.delete('1.0', tk.END)
         self.status_text.config(state=tk.DISABLED)

    def update_status(self, message):
        """Appends a message to the status log, handling thread safety."""
        # Ensure GUI updates happen on the main thread
        if threading.current_thread() != threading.main_thread():
            # If called from a background thread, schedule the update
            self.master.after(0, self._update_status_text, message)
        else:
            # If already on the main thread, update directly
            self._update_status_text(message)

    def _update_status_text(self, message):
        """Internal helper to update the status text widget."""
        current_state = self.status_text.cget('state') # Store current state
        self.status_text.config(state=tk.NORMAL)      # Enable writing
        self.status_text.insert(tk.END, message + "\n") # Add message
        self.status_text.see(tk.END)                   # Scroll to the end
        self.status_text.config(state=current_state)   # Restore original state (usually DISABLED)

    def disable_buttons(self):
        """Disables buttons during processing."""
        self.processing_active = True
        self.select_button.config(state=tk.DISABLED)
        self.dba_button.config(state=tk.DISABLED)
        self.totalnot_button.config(state=tk.DISABLED)
        self.contpaq_button.config(state=tk.DISABLED)
        self.zoom_in_button.config(state=tk.DISABLED)
        self.zoom_out_button.config(state=tk.DISABLED)
        self.reset_zoom_button.config(state=tk.DISABLED)
        self.save_changes_button.config(state=tk.DISABLED) # Also disable save during processing
        # Cancel any active cell edit
        self._cancel_cell_edit()

    def enable_buttons(self):
        """Enables buttons after processing or on startup."""
        self.processing_active = False
        self.select_button.config(state=tk.NORMAL)
        self.dba_button.config(state=tk.NORMAL)
        self.totalnot_button.config(state=tk.NORMAL)
        self.contpaq_button.config(state=tk.NORMAL)
        self.zoom_in_button.config(state=tk.NORMAL)
        self.zoom_out_button.config(state=tk.NORMAL)
        self.reset_zoom_button.config(state=tk.NORMAL)
        # Enable save button only if data exists AND an Excel file path is known
        if self.current_excel_file and self.dataframe is not None and not self.dataframe.empty:
            self.save_changes_button.config(state=tk.NORMAL)
        else:
            self.save_changes_button.config(state=tk.DISABLED)

    def start_processing(self, invoice_type):
        """Initiates the PDF processing in a separate thread."""
        folder = self.selected_folder.get()
        # Validate folder selection
        if not folder or folder == "No folder selected":
            messagebox.showerror("Error", "Please select a folder first.", parent=self.master)
            return
        if not os.path.isdir(folder):
             messagebox.showerror("Error", f"Invalid directory selected:\n{folder}", parent=self.master)
             return
        # Prevent starting multiple processes
        if self.processing_active:
            messagebox.showwarning("Busy", "Processing is already in progress.", parent=self.master)
            return

        # Prepare GUI for processing
        self.disable_buttons()
        self.clear_status()
        self.clear_treeview() # Clear previous data and state
        self.pdf_scroll_positions.clear() # [NEW] Clear scroll positions
        self.clear_pdf_preview(f"Processing {invoice_type} invoices...\nPlease wait.")
        self.update_status(f"Starting recursive processing for {invoice_type} in: {folder}")
        self.update_status("-" * 40) # Separator in log

        # Reset sort state before processing new data
        self._tree_current_sort_col = None
        self._tree_current_sort_ascending = True

        # Run the processing logic in a background thread
        process_thread = threading.Thread(target=run_processing, args=(folder, invoice_type, self), daemon=True)
        process_thread.start()

    def set_data_and_file(self, df, excel_file_path):
        """Stores the processed DataFrame and the path it was saved to."""
        self.dataframe = df
        self.current_excel_file = excel_file_path
        self._clipboard = [] # Clear clipboard when new data is loaded
        # Update save button state based on whether a file path is known
        if self.current_excel_file and self.dataframe is not None and not self.dataframe.empty:
             self.save_changes_button.config(state=tk.NORMAL)
        else:
             self.save_changes_button.config(state=tk.DISABLED)

    def clear_treeview(self):
        """Clears the Treeview, internal maps, DataFrame, and related state."""
        if self._edit_entry: # Ensure any active edit is cancelled
            self._cancel_cell_edit()

        # Safely clear tree items (unbind/rebind to avoid Tcl errors during deletion)
        if hasattr(self, 'tree'):
            try:
                 # Unbind right-click during clear to prevent issues if menu pops up
                 self.tree.unbind('<Button-3>')
            except tk.TclError: pass # Ignore if already unbound or widget destroyed
            # Delete all items
            for item in self.tree.get_children():
                try:
                    self.tree.delete(item)
                except tk.TclError:
                    pass # Item might already be gone if clearing rapidly
            # Rebind right-click
            try:
                 self.tree.bind('<Button-3>', self.show_context_menu)
            except tk.TclError: pass

        # Clear data and state
        self.dataframe = None
        self.current_excel_file = None
        self._pdf_path_map.clear()
        self._df_index_map.clear()
        self._clipboard = []
        # Note: We don't clear pdf_scroll_positions here, it's cleared in select_folder/start_processing

        # Reset sort indicators in headers
        if hasattr(self, 'tree'):
             for col in self.tree["columns"]:
                 try:
                      current_text = self.tree.heading(col, "text").replace(' ▲', '').replace(' ▼', '')
                      self.tree.heading(col, text=current_text)
                 except tk.TclError: pass # Ignore if widget destroyed

        # Reset sort state variables
        self._tree_current_sort_col = None
        self._tree_current_sort_ascending = True

        # Ensure save button is disabled
        if hasattr(self, 'save_changes_button'):
            self.save_changes_button.config(state=tk.DISABLED)

    def load_data_to_treeview(self):
        """Populates the Treeview from the current self.dataframe."""
        if self._edit_entry: self._cancel_cell_edit() # Cancel edits before reload

        # --- Start Fresh: Clear existing items and maps ---
        # Safely clear tree items
        self.tree.unbind('<Button-3>') # Unbind temporarily
        for item in self.tree.get_children():
            try: self.tree.delete(item)
            except tk.TclError: pass
        self.tree.bind('<Button-3>', self.show_context_menu) # Rebind

        # Clear internal maps that link Treeview IDs to data
        self._pdf_path_map.clear()
        self._df_index_map.clear()
        # --- End Fresh Start ---

        # Check if there's data to load
        if self.dataframe is None or self.dataframe.empty:
             self.update_status("No data available to display.")
             self.clear_pdf_preview("No data loaded.")
             self.save_changes_button.config(state=tk.DISABLED)
             return

        # Verify required columns exist in the DataFrame
        display_columns = list(self.tree["columns"])
        required_cols = display_columns + ["Full PDF Path"] # Need Full PDF Path for previews
        if not all(col in self.dataframe.columns for col in required_cols):
            missing = [col for col in required_cols if col not in self.dataframe.columns]
            errmsg = f"Error: DataFrame is missing required columns: {missing}"
            self.update_status(errmsg)
            messagebox.showerror("Data Loading Error", errmsg, parent=self.master)
            print(f"DataFrame columns available: {self.dataframe.columns.tolist()}")
            self.clear_pdf_preview("Error loading data. See status log.")
            # Clear the invalid data
            self.dataframe = None
            self.current_excel_file = None
            self._clipboard = []
            self.save_changes_button.config(state=tk.DISABLED)
            return

        # Ensure Treeview only shows the intended columns
        self.tree.configure(displaycolumns=display_columns)

        last_item_id = None # Keep track of the last added item ID

        # Check and potentially fix DataFrame index if it's not unique
        if not self.dataframe.index.is_unique:
             self.update_status("Warning: DataFrame index is not unique. Resetting index.")
             print("DataFrame index was not unique, resetting...")
             self.dataframe.reset_index(drop=True, inplace=True)

        # --- Populate Treeview ---
        # Iterate through DataFrame rows using iterrows() which gives (index, Series)
        for df_index, row in self.dataframe.iterrows():
            try:
                # Prepare values for display (convert NaN/None to empty string)
                display_values = tuple(str(row[col]) if pd.notna(row[col]) else "" for col in display_columns)
                full_path = row["Full PDF Path"] # Get the full path for mapping

                # Use the DataFrame index as the unique Treeview item ID (iid)
                # Convert to string as iid should be a string
                item_id = str(df_index)

                # Insert the row into the Treeview
                self.tree.insert("", tk.END, values=display_values, iid=item_id)

                # Store mappings: Treeview item ID -> Full PDF Path (for preview)
                # Store only if path is a valid string
                if full_path and isinstance(full_path, str):
                    self._pdf_path_map[item_id] = full_path

                # Store mappings: Treeview item ID -> DataFrame index (for edits/deletes)
                self._df_index_map[item_id] = df_index # df_index is the actual index from the DataFrame

                last_item_id = item_id # Track the last added ID

            except Exception as e:
                # Catch errors during row insertion (less likely with checks, but good practice)
                print(f"Error adding row with DataFrame index {df_index} (iid: {item_id}) to treeview: {e}")
                self.update_status(f"Warning: Could not display row for PDF '{row.get('Source PDF', 'Unknown')}' in table.")

        # --- Post-Load Updates ---
        row_count = len(self.tree.get_children())
        if row_count == 0 and not self.dataframe.empty:
            # This case might indicate an issue with the insertion loop or data format
            self.update_status("Loaded data frame, but no rows were added to the table. Check data integrity.")
            self.clear_pdf_preview("Data loaded, but no rows to display.")
        elif row_count == 0 and self.dataframe.empty:
             # This is expected if the dataframe is empty
             self.update_status("Data loaded, but it contains no rows.")
             self.clear_pdf_preview("No data loaded.")

        # Update save button state (might be enabled if initial save failed but data is now loaded)
        if self.current_excel_file and not self.dataframe.empty:
            self.save_changes_button.config(state=tk.NORMAL)
        else:
            # Disable if no excel file path or if dataframe became empty
             self.save_changes_button.config(state=tk.DISABLED if self.dataframe.empty else tk.NORMAL if self.current_excel_file else tk.DISABLED)


    # --- Treeview Sorting ---
    def sort_treeview_column(self, col): # Takes only the column ID/Name
        """Sorts the Treeview based on the clicked column header."""
        if self.dataframe is None or self.dataframe.empty: return
        if self._edit_entry: self._cancel_cell_edit() # Cancel edits before sorting

        try:
            # --- Determine Sort Direction ---
            if self._tree_current_sort_col == col:
                # Same column clicked again, reverse direction
                sort_ascending = not self._tree_current_sort_ascending
            else:
                # New column clicked, sort ascending by default
                sort_ascending = True

            # --- Perform Sorting on DataFrame Copy ---
            sorted_df = self.dataframe.copy()

            # Determine data type for appropriate sorting strategy
            numeric_col_check = pd.to_numeric(sorted_df[col], errors='coerce')
            is_numeric_type = pd.api.types.is_numeric_dtype(numeric_col_check)
            sort_numerically = not numeric_col_check.isna().all() and is_numeric_type

            print(f"Sorting column '{col}'. Ascending: {sort_ascending}. Sort numerically: {sort_numerically}") # Debugging output

            if sort_numerically:
                # --- Numeric Sort ---
                temp_sort_col = f"__{col}_numeric_sort"
                sorted_df[temp_sort_col] = numeric_col_check
                sorted_df = sorted_df.sort_values(by=temp_sort_col, ascending=sort_ascending, na_position='last')
                sorted_df.drop(columns=[temp_sort_col], inplace=True)
            else:
                # --- String Sort ---
                sorted_df = sorted_df.sort_values(
                    by=col,
                    ascending=sort_ascending,
                    key=lambda x: x.map(lambda s: str(s).lower() if pd.notna(s) else ''),
                    na_position='last'
                )

            # --- Update Main DataFrame and State ---
            self.dataframe = sorted_df # Replace the old dataframe with the sorted one
            self._tree_current_sort_col = col # Store the column that was just sorted
            self._tree_current_sort_ascending = sort_ascending # Store the direction

            # --- Update Treeview Headers (Arrows) ---
            for c in self.tree["columns"]:
                current_text = self.tree.heading(c, "text").replace(' ▲', '').replace(' ▼', '')
                if c == col:
                    indicator = ' ▲' if sort_ascending else ' ▼'
                    self.tree.heading(c, text=current_text + indicator)
                else:
                    self.tree.heading(c, text=current_text)

            # --- Reload Treeview with Sorted Data ---
            self.load_data_to_treeview()

        except KeyError:
            errmsg = f"Error: Column '{col}' not found in DataFrame for sorting."
            print(errmsg)
            messagebox.showerror("Sort Error", f"Column '{col}' not found.", parent=self.master)
            self._tree_current_sort_col = None
            self._tree_current_sort_ascending = True
        except Exception as e:
            errmsg = f"Error sorting treeview column '{col}': {e}"
            print(errmsg)
            import traceback
            traceback.print_exc()
            messagebox.showerror("Sort Error", f"Could not sort column '{col}':\n{e}", parent=self.master)
            self._tree_current_sort_col = None
            self._tree_current_sort_ascending = True


    # --- Inline Editing ---
    def on_tree_double_click(self, event):
        """Handles double-clicking on a cell for inline editing."""
        if self._edit_entry: return # Already editing

        region = self.tree.identify_region(event.x, event.y)
        if region != "cell": return

        item_id = self.tree.identify_row(event.y)
        column_id_str = self.tree.identify_column(event.x)

        if not item_id or not column_id_str: return

        try:
            column_index = int(column_id_str.replace('#', '')) - 1
            if column_index < 0 or column_index >= len(self.tree["columns"]): return
            column_name = self.tree["columns"][column_index]
        except (ValueError, IndexError):
            print(f"Error identifying column from ID: {column_id_str}")
            return

        if column_name not in self.editable_columns:
            return

        try:
             bbox = self.tree.bbox(item_id, column=column_id_str)
        except Exception:
             return
        if not bbox: return

        current_values = self.tree.item(item_id, 'values')
        if column_index >= len(current_values): return
        original_value = current_values[column_index]

        self._edit_item_id = item_id
        self._edit_column_id = column_id_str

        self._edit_entry = ttk.Entry(self.tree, font=('Segoe UI', 9))
        self._edit_entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])

        self._edit_entry.insert(0, original_value)
        self._edit_entry.select_range(0, tk.END)
        self._edit_entry.focus_set()

        self._edit_entry.bind("<Return>", lambda e: self._save_cell_edit(item_id, column_index, column_name))
        self._edit_entry.bind("<KP_Enter>", lambda e: self._save_cell_edit(item_id, column_index, column_name))
        self._edit_entry.bind("<Escape>", lambda e: self._cancel_cell_edit())
        self._edit_entry.bind("<FocusOut>", lambda e: self._save_cell_edit(item_id, column_index, column_name))

    def _save_cell_edit(self, item_id, column_index, column_name):
        """Saves the value from the inline edit Entry to the Treeview and DataFrame."""
        if not self._edit_entry: return

        new_value = self._edit_entry.get()
        self._cancel_cell_edit() # Cleanup first

        try:
             current_values = list(self.tree.item(item_id, 'values'))
        except tk.TclError:
             print(f"Warning: Treeview item {item_id} no longer exists. Edit not saved.")
             return

        if column_index < len(current_values) and str(current_values[column_index]) != new_value:
            current_values[column_index] = new_value
            try:
                 self.tree.item(item_id, values=tuple(current_values))
            except tk.TclError:
                 print(f"Warning: Treeview item {item_id} no longer exists. Edit display not updated.")

            if self.dataframe is not None:
                try:
                    df_index = self._df_index_map.get(item_id)
                    if df_index is not None and df_index in self.dataframe.index:
                        self.dataframe.loc[df_index, column_name] = new_value
                        self.update_status(f"Cell updated: Row index {df_index}, Col '{column_name}' = '{new_value}'")
                        if self.current_excel_file:
                            self.save_changes_button.config(state=tk.NORMAL)
                    else:
                        print(f"Error: Could not find DataFrame index ({df_index}) for tree item {item_id} during edit save.")
                        self.update_status(f"Error: Failed to update backing data for edit (Index mapping issue).")
                except Exception as e:
                    print(f"Error updating DataFrame at index {df_index}, column {column_name}: {e}")
                    self.update_status(f"Error: Failed to update backing data for edit ({type(e).__name__}).")

    def _cancel_cell_edit(self, event=None):
        """Destroys the inline edit Entry widget and resets edit state."""
        if self._edit_entry:
            try:
                self._edit_entry.destroy()
            except tk.TclError:
                pass # Widget might already be destroyed
            self._edit_entry = None
            self._edit_item_id = None
            self._edit_column_id = None


    # --- Context Menu and Row Operations ---
    def show_context_menu(self, event):
        """Displays the right-click context menu."""
        item_id = self.tree.identify_row(event.y)

        if item_id:
            if item_id not in self.tree.selection():
                self.tree.selection_set(item_id)

        has_selection = bool(self.tree.selection())
        can_paste = bool(self._clipboard)
        can_add = self.dataframe is not None

        self.context_menu.entryconfig("Add Blank Row", state=tk.NORMAL if can_add else tk.DISABLED)
        self.context_menu.entryconfig("Copy Selected Row(s)", state=tk.NORMAL if has_selection else tk.DISABLED)
        self.context_menu.entryconfig("Paste Row(s)", state=tk.NORMAL if can_paste else tk.DISABLED)
        self.context_menu.entryconfig("Delete Selected Row(s)", state=tk.NORMAL if has_selection else tk.DISABLED)

        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def _add_row(self):
        """Adds a new blank row to the DataFrame and reloads the Treeview."""
        if self.dataframe is None:
            self.update_status("Cannot add row: No data loaded.")
            messagebox.showwarning("Add Row Failed", "Load or process data first before adding rows.", parent=self.master)
            return
        if self._edit_entry: self._cancel_cell_edit()

        try:
            new_row_data = {col: "" for col in self.dataframe.columns}
            new_row_data["Source PDF"] = "[MANUALLY ADDED]"
            new_row_data["Full PDF Path"] = ""
            new_row_data["Invoice Folio"] = ""
            new_row_data["Document Type"] = ""
            new_row_data["Reference Number"] = ""

            new_row_df = pd.DataFrame([new_row_data])
            self.dataframe = pd.concat([self.dataframe, new_row_df], ignore_index=True)

            self.load_data_to_treeview()

            children = self.tree.get_children()
            if children:
                new_item_id = children[-1]
                self.tree.see(new_item_id)
                self.tree.selection_set(new_item_id)
                self.tree.focus(new_item_id)

            self.update_status("Added 1 blank row.")
            if self.current_excel_file and not self.dataframe.empty:
                 self.save_changes_button.config(state=tk.NORMAL)

        except Exception as e:
            print(f"Error adding row: {e}")
            self.update_status(f"Error: Could not add row ({type(e).__name__}).")
            import traceback; traceback.print_exc()
            try:
                self.load_data_to_treeview()
            except Exception as reload_e:
                print(f"Error reloading treeview after add row error: {reload_e}")

    def _delete_selected_rows(self):
        """Deletes selected rows from DataFrame and Treeview."""
        selected_item_ids = self.tree.selection()
        if not selected_item_ids:
            self.update_status("No rows selected to delete.")
            return

        num_selected = len(selected_item_ids)
        if not messagebox.askyesno("Confirm Delete",
                                   f"Are you sure you want to permanently delete {num_selected} selected row(s)?\n"
                                   "This action cannot be undone.",
                                   parent=self.master):
            self.update_status("Deletion cancelled.")
            return

        if self._edit_entry: self._cancel_cell_edit()

        try:
            indices_to_drop = []
            failed_map_ids = []
            for item_id in selected_item_ids:
                df_index = self._df_index_map.get(item_id)
                if df_index is not None and df_index in self.dataframe.index:
                    indices_to_drop.append(df_index)
                else:
                    failed_map_ids.append(item_id)
                    print(f"Warning: Could not find valid DataFrame index for tree item '{item_id}' during delete.")

            if failed_map_ids:
                 self.update_status(f"Warning: Could not map {len(failed_map_ids)} selected item(s) to data for deletion.")

            if not indices_to_drop:
                 self.update_status("Error: No valid data found for selected rows to delete.")
                 return

            self.dataframe.drop(index=indices_to_drop, inplace=True)
            self.dataframe.reset_index(drop=True, inplace=True)
            self.load_data_to_treeview()

            self.update_status(f"Deleted {len(indices_to_drop)} row(s).")

            if self.current_excel_file and self.dataframe is not None and not self.dataframe.empty:
                 self.save_changes_button.config(state=tk.NORMAL)
            else:
                 self.save_changes_button.config(state=tk.DISABLED)

        except Exception as e:
            print(f"Error deleting rows: {e}")
            self.update_status(f"Error: Could not delete rows ({type(e).__name__}).")
            import traceback; traceback.print_exc()
            try:
                 self.load_data_to_treeview()
            except Exception as reload_e:
                 print(f"Error reloading treeview after delete rows error: {reload_e}")

    def _copy_selected_rows(self):
        """Copies data of selected rows to the internal clipboard (_clipboard)."""
        selected_item_ids = self.tree.selection()
        if not selected_item_ids:
            self.update_status("No rows selected to copy.")
            return
        if self.dataframe is None:
            self.update_status("Cannot copy: Data not loaded.")
            return

        self._clipboard = []
        copied_count = 0

        try:
            df_indices_to_copy = []
            failed_map_ids = []
            for item_id in selected_item_ids:
                 df_index = self._df_index_map.get(item_id)
                 if df_index is not None and df_index in self.dataframe.index:
                      df_indices_to_copy.append(df_index)
                 else:
                      failed_map_ids.append(item_id)
                      print(f"Warning: Could not find valid DataFrame index for tree item '{item_id}' during copy.")

            if failed_map_ids:
                 self.update_status(f"Warning: Could not map {len(failed_map_ids)} selected item(s) to data for copying.")

            if not df_indices_to_copy:
                 self.update_status("Error: Could not map any selected items to data for copying.")
                 return

            copied_data_df = self.dataframe.loc[df_indices_to_copy]
            self._clipboard = copy.deepcopy(copied_data_df.to_dict('records'))
            copied_count = len(self._clipboard)

            if copied_count > 0:
                self.update_status(f"Copied {copied_count} row(s) to clipboard.")
            else:
                self.update_status("Warning: No row data was copied.")

        except Exception as e:
            print(f"Error copying rows: {e}")
            self.update_status(f"Error: Could not copy rows ({type(e).__name__}).")
            self._clipboard = []
            import traceback; traceback.print_exc()

    def _paste_rows(self):
        """Pastes rows from the clipboard (_clipboard) into the DataFrame and Treeview."""
        if not self._clipboard:
            self.update_status("Clipboard is empty. Nothing to paste.")
            return
        if self.dataframe is None:
            self.update_status("Cannot paste: No data loaded.")
            messagebox.showwarning("Paste Failed", "Load or process data first before pasting rows.", parent=self.master)
            return
        if self._edit_entry: self._cancel_cell_edit()

        try:
            pasted_df = pd.DataFrame(self._clipboard)
            if pasted_df.empty:
                 self.update_status("Clipboard contained no valid data to paste.")
                 return

            pasted_df = pasted_df.reindex(columns=self.dataframe.columns, fill_value="")
            if pasted_df.empty:
                 self.update_status("No matching columns found between clipboard and table. Nothing pasted.")
                 return

            self.dataframe = pd.concat([self.dataframe, pasted_df], ignore_index=True)
            num_pasted = len(pasted_df)
            self.load_data_to_treeview()

            children = self.tree.get_children()
            if len(children) >= num_pasted:
                 first_pasted_item_id = children[-num_pasted]
                 self.tree.selection_set(children[-num_pasted:])
                 self.tree.see(first_pasted_item_id)

            self.update_status(f"Pasted {num_pasted} row(s) from clipboard.")
            if self.current_excel_file and not self.dataframe.empty:
                 self.save_changes_button.config(state=tk.NORMAL)

        except Exception as e:
            print(f"Error pasting rows: {e}")
            self.update_status(f"Error: Could not paste rows ({type(e).__name__}).")
            import traceback; traceback.print_exc()
            try:
                 self.load_data_to_treeview()
            except Exception as reload_e:
                 print(f"Error reloading treeview after paste rows error: {reload_e}")


    # --- Excel Saving Method ---
    def save_changes_to_excel(self):
        """Saves the current DataFrame content to the stored Excel file path."""
        if self._edit_entry: self._cancel_cell_edit()

        if self.dataframe is None or self.dataframe.empty:
            messagebox.showwarning("No Data", "There is no data to save.", parent=self.master)
            return
        if not self.current_excel_file:
            messagebox.showerror("Save Error", "The Excel file path is unknown.\n"
                                             "Cannot save changes. (Did the initial save fail?)", parent=self.master)
            self.update_status("Save failed: Excel file path not set.")
            return
        excel_dir = os.path.dirname(self.current_excel_file)
        if not os.path.isdir(excel_dir):
            messagebox.showerror("Save Error", f"The directory for the Excel file no longer exists:\n{excel_dir}\n\nCannot save changes.", parent=self.master)
            self.update_status(f"Save failed: Directory not found - {excel_dir}")
            return

        if not messagebox.askyesno("Confirm Save",
                                   f"This will overwrite the existing file:\n{self.current_excel_file}\n\n"
                                   "Are you sure you want to save the current data?",
                                   parent=self.master):
            self.update_status("Save cancelled by user.")
            return

        self.update_status(f"Attempting to save changes to: {self.current_excel_file}")

        try:
            excel_columns = ["Source PDF", "Invoice Folio", "Document Type", "Reference Number"]
            cols_to_save = [col for col in excel_columns if col in self.dataframe.columns]
            if not cols_to_save:
                 raise ValueError("No valid columns found in the DataFrame to save.")

            df_to_save = self.dataframe[cols_to_save].copy()
            df_to_save.to_excel(self.current_excel_file, index=False, engine='openpyxl')

            success_msg = f"Changes successfully saved to:\n{self.current_excel_file}"
            self.update_status(success_msg)
            messagebox.showinfo("Save Successful", success_msg, parent=self.master)

        except PermissionError:
            error_message = (f"Error saving changes to:\n{self.current_excel_file}\n\n"
                             f"Permission denied. Is the file open in Excel or another program?\n\n"
                             f"Close the file and try saving again.")
            print(error_message)
            self.update_status(f"Save Error: Permission Denied.")
            messagebox.showerror("Save Error", error_message, parent=self.master)
        except Exception as e:
            error_message = f"An unexpected error occurred while saving changes to Excel:\n\n{e}"
            print(error_message)
            import traceback; traceback.print_exc()
            self.update_status(f"Save Error: {e}. See console.")
            messagebox.showerror("Save Error", error_message, parent=self.master)


    # --- Treeview Selection & PDF Preview ---
    def _store_current_scroll_position(self):
        """ [NEW] Stores the current scroll position of the PDF canvas."""
        if self.current_preview_pdf_path and hasattr(self, 'pdf_canvas') and self.pdf_canvas.winfo_exists():
            try:
                x_frac = self.pdf_canvas.xview()[0]
                y_frac = self.pdf_canvas.yview()[0]
                # Only store if not at the very top-left (avoids storing default)
                # Or store always if you want to remember even the top-left state
                if x_frac > 0.0 or y_frac > 0.0:
                    self.pdf_scroll_positions[self.current_preview_pdf_path] = (x_frac, y_frac)
                elif self.current_preview_pdf_path in self.pdf_scroll_positions:
                    # If scrolled back to top-left, remove stored position to default next time
                    # Or keep it stored as (0.0, 0.0) - choice depends on desired behavior
                    del self.pdf_scroll_positions[self.current_preview_pdf_path]

            except tk.TclError as e:
                print(f"Warning: TclError getting scroll position: {e}")
            except Exception as e:
                print(f"Error getting scroll position: {e}")


    def on_treeview_select(self, event):
        """Handles selection changes in the Treeview to update the PDF preview."""
        # If an edit is active when selection changes, try to save it first
        if self._edit_entry and self.tree.focus() != self._edit_entry:
             if self._edit_item_id and self._edit_column_id:
                  try:
                       col_idx = int(self._edit_column_id.replace('#', '')) - 1
                       col_name = self.tree["columns"][col_idx]
                       self._save_cell_edit(self._edit_item_id, col_idx, col_name)
                  except Exception:
                       self._cancel_cell_edit()
             else:
                  self._cancel_cell_edit()

        # --- [NEW] Store scroll position of the *current* PDF before changing ---
        self._store_current_scroll_position()
        # --- End New Section ---

        selected_items = self.tree.selection()
        if not selected_items:
            # Optionally clear preview or just do nothing
            # self.clear_pdf_preview("No row selected.")
            # self.current_preview_pdf_path = None # Already handled by clear_pdf_preview if called
            return

        selected_item_id = selected_items[0]
        pdf_full_path = self._pdf_path_map.get(selected_item_id)

        if pdf_full_path and isinstance(pdf_full_path, str):
            if os.path.exists(pdf_full_path):
                # Only update if the path is different
                if pdf_full_path != self.current_preview_pdf_path:
                    self.update_pdf_preview(pdf_full_path)
                # else: scroll position will be restored by _update_canvas_scrollregion if needed
            else:
                base_name = os.path.basename(pdf_full_path)
                err_msg = f"File Not Found:\n{base_name}\n(Path: {pdf_full_path})"
                self.clear_pdf_preview(err_msg)
                print(f"Error: File path from selection '{pdf_full_path}' does not exist.")
                self.update_status(f"Preview Error: File not found - {base_name}")
                self.current_preview_pdf_path = None # Clear current path state
        elif pdf_full_path is None and selected_item_id in self._df_index_map:
             self.clear_pdf_preview("No PDF associated with this row.")
             self.current_preview_pdf_path = None
        elif selected_item_id not in self._df_index_map:
             print(f"Error: Selected item ID {selected_item_id} not found in data map.")
             self.clear_pdf_preview("Error: Cannot find data for selected row.")
             self.current_preview_pdf_path = None
        else:
             print(f"Error: Invalid path data associated with selected row ID {selected_item_id}: {pdf_full_path}")
             self.clear_pdf_preview("Error: Invalid file path\nin selected row data.")
             self.current_preview_pdf_path = None

    def clear_pdf_preview(self, message="Select a row to preview PDF"):
        """Clears the PDF preview area and displays a placeholder message."""
        # --- [NEW] Store scroll position before clearing ---
        self._store_current_scroll_position()
        # --- End New Section ---

        if self._canvas_image_id:
             try: self.pdf_canvas.delete(self._canvas_image_id)
             except tk.TclError: pass
             self._canvas_image_id = None
        self.pdf_preview_image = None

        if self._placeholder_window_id:
            try: self.pdf_canvas.delete("placeholder")
            except tk.TclError: pass
            self._placeholder_window_id = None

        if hasattr(self, 'pdf_placeholder_label') and self.pdf_placeholder_label.winfo_exists():
            self.pdf_placeholder_label.config(text=message)
        else:
            self.pdf_placeholder_label = ttk.Label(self.pdf_canvas, text=message, style="Placeholder.TLabel")

        canvas_w = self.pdf_canvas.winfo_width() if self.pdf_canvas.winfo_width() > 1 else 200
        canvas_h = self.pdf_canvas.winfo_height() if self.pdf_canvas.winfo_height() > 1 else 200
        self._placeholder_window_id = self.pdf_canvas.create_window(
            canvas_w//2, canvas_h//2,
            window=self.pdf_placeholder_label,
            anchor=tk.CENTER,
            tags="placeholder"
        )

        # Update scroll region, but don't pass a path to restore
        self.master.after_idle(self._update_canvas_scrollregion, None) # [MODIFIED] Pass None
        self.current_preview_pdf_path = None

    def update_pdf_preview(self, pdf_path):
        """Loads and displays the first page of the specified PDF."""
        if not PIL_AVAILABLE:
            self.clear_pdf_preview("PDF Preview disabled:\nPillow library not found.")
            return

        # --- Clear previous content ---
        if self._placeholder_window_id:
            try: self.pdf_canvas.delete("placeholder")
            except tk.TclError: pass
            self._placeholder_window_id = None
        if self._canvas_image_id:
            try: self.pdf_canvas.delete(self._canvas_image_id)
            except tk.TclError: pass
            self._canvas_image_id = None
        self.pdf_preview_image = None
        # --- End Clear ---

        try:
            doc = fitz.open(pdf_path)
            if len(doc) == 0:
                raise ValueError("PDF has no pages.")

            page = doc.load_page(0)
            page_rect = page.rect
            if page_rect.width == 0 or page_rect.height == 0:
                 raise ValueError("PDF page has zero dimensions.")

            zoom_factor = self.current_zoom_factor
            mat = fitz.Matrix(zoom_factor, zoom_factor)
            pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB, alpha=False)
            doc.close()

            img_bytes = pix.samples
            if not isinstance(img_bytes, bytes):
                 img_bytes = bytes(img_bytes)

            if not img_bytes:
                 raise ValueError("Pixmap samples are empty after rendering.")

            pil_image = Image.frombytes("RGB", [pix.width, pix.height], img_bytes)
            self.pdf_preview_image = ImageTk.PhotoImage(image=pil_image)

            if self.pdf_preview_image:
                # Store the path *before* creating the image and scheduling scroll restore
                self.current_preview_pdf_path = pdf_path
                # Create image
                self._canvas_image_id = self.pdf_canvas.create_image(
                    0, 0, anchor=tk.NW, image=self.pdf_preview_image, tags="pdf_image"
                )
                # Update scroll region AND attempt to restore scroll position *after* layout update
                self.master.after_idle(self._update_canvas_scrollregion, pdf_path) # [MODIFIED] Pass pdf_path
                # Note: Initial scroll to top-left is handled by _update_canvas_scrollregion if no stored position is found.
            else:
                raise ValueError("Failed to create PhotoImage object.")

        except (fitz.fitz.FileNotFoundError, fitz.fitz.PasswordError, ValueError, Exception) as e:
            base_name = os.path.basename(pdf_path) if pdf_path else "Unknown File"
            error_type = type(e).__name__
            error_msg = f"Preview Error ({error_type}):\n{base_name}"
            if isinstance(e, fitz.fitz.PasswordError):
                error_msg += "\n(Password Protected?)"
            elif isinstance(e, ValueError):
                 error_msg += f"\n({e})"
            import traceback
            print(f"--- PDF Preview Exception Traceback for {base_name} ---")
            traceback.print_exc()
            print(f"--- End Traceback ---")
            self.clear_pdf_preview(error_msg) # clear_pdf_preview also resets current_preview_pdf_path


    # --- Zoom Methods ---
    def zoom_in(self):
        """Increases the zoom factor and updates the preview."""
        if self.current_preview_pdf_path:
            # [NEW] Store current scroll position before zooming
            self._store_current_scroll_position()
            new_zoom = min(self.current_zoom_factor * self.zoom_step, self.max_zoom)
            if new_zoom != self.current_zoom_factor:
                self.current_zoom_factor = new_zoom
                self.update_pdf_preview(self.current_preview_pdf_path) # Reload with new zoom

    def zoom_out(self):
        """Decreases the zoom factor and updates the preview."""
        if self.current_preview_pdf_path:
            # [NEW] Store current scroll position before zooming
            self._store_current_scroll_position()
            new_zoom = max(self.current_zoom_factor / self.zoom_step, self.min_zoom)
            if new_zoom != self.current_zoom_factor:
                self.current_zoom_factor = new_zoom
                self.update_pdf_preview(self.current_preview_pdf_path)

    def reset_zoom(self):
        """Resets the zoom factor to 1.0 and updates the preview."""
        if self.current_preview_pdf_path and self.current_zoom_factor != 1.0:
            # [NEW] Store current scroll position before zooming
            self._store_current_scroll_position()
            self.current_zoom_factor = 1.0
            self.update_pdf_preview(self.current_preview_pdf_path)

    def _update_canvas_scrollregion(self, pdf_path_to_restore=None): # [MODIFIED] Added argument
        """Updates the canvas scroll region and restores scroll position if applicable."""
        try:
            scroll_bbox = None # Initialize

            # If an image is currently displayed
            if self._canvas_image_id and self.pdf_canvas.find_withtag(self._canvas_image_id):
                bbox = self.pdf_canvas.bbox(self._canvas_image_id)
                if bbox:
                    scroll_bbox = (bbox[0], bbox[1], bbox[2] + 5, bbox[3] + 5)
            else:
                # If no image, set scroll region to current canvas size
                current_width = self.pdf_canvas.winfo_width()
                current_height = self.pdf_canvas.winfo_height()
                scroll_bbox = (0, 0, max(1, current_width), max(1, current_height))

            # Apply the calculated scroll region
            if scroll_bbox:
                self.pdf_canvas.config(scrollregion=scroll_bbox)

            # --- [NEW] Attempt to restore scroll position ---
            # Check if we were asked to restore a specific path AND that path is still the one currently displayed
            if pdf_path_to_restore and pdf_path_to_restore == self.current_preview_pdf_path:
                # Get stored position, default to (0.0, 0.0) if not found
                pos = self.pdf_scroll_positions.get(pdf_path_to_restore, (0.0, 0.0))
                # Apply the position
                self.pdf_canvas.xview_moveto(pos[0])
                self.pdf_canvas.yview_moveto(pos[1])
                # print(f"Restored scroll for {os.path.basename(pdf_path_to_restore)} to {pos}") # Debug
            elif pdf_path_to_restore is None: # If called from clear_pdf_preview, ensure top-left
                 self.pdf_canvas.xview_moveto(0.0)
                 self.pdf_canvas.yview_moveto(0.0)
            # If pdf_path_to_restore doesn't match current_preview_pdf_path (e.g., rapid selection change),
            # do nothing, let the default top-left remain or wait for the *next* update cycle.
            # --- End New Section ---

        except tk.TclError as e:
            print(f"Warning: TclError updating canvas scrollregion/restoring scroll: {e}")
        except Exception as e:
            print(f"Error updating canvas scrollregion/restoring scroll: {e}")


# --- Main Execution ---
if __name__ == "__main__":
    if not PIL_AVAILABLE:
         root_check = tk.Tk()
         root_check.withdraw()
         messagebox.showwarning("Dependency Warning",
                                "Python Imaging Library (Pillow) not found.\n"
                                "PDF preview will be disabled.\n\n"
                                "Install it using:\npip install Pillow",
                                parent=None)
         root_check.destroy()

    root = tk.Tk()
    app = InvoiceProcessorApp(root)
    root.mainloop()





