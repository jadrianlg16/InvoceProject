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


# --- Regex Patterns ---
# (Keep all existing regex patterns)
REGEX_ESCRITURA_RANGE = r'(?:Escritura|Esc)\s*(?:P[úu]blica)?\s*(?:N[uú]meros?|Nos\.?|N°)\s+(\d+)\s+A\s+(\d+)\b'
REGEX_ACTA_RANGE = r'Acta(?:s)?\s+Fuera\s+de\s+Protocolo\s+(?:N[uú]meros?|Nos\.?|N°)\s+(\d+)\s+A\s+(\d+)\b'
REGEX_ESCRITURA_LIST_Y = r'(?:Escritura|Esc)\s*(?:P[úu]blica)?\s*(?:N[uú]meros?|Nos\.?|N°)\s+(\d+)\s+Y\s+(\d+)\b'
REGEX_ACTA_LIST_Y = r'Acta(?:s)?\s+Fuera\s+de\s+Protocolo\s+(?:N[uú]meros?|Nos\.?|N°)\s+(\d+)\s+Y\s+(\d+)\b'
REGEX_ACTA_SPECIAL = r'Acta\s+Fuera\s+de\s+Protocolo\s+N[uú]mero\s+\d+\/(\d+)\/\d+\b'
REGEX_ESCRITURA_SINGLE = r'(?:Escritura|Esc)\s*(?:P[úu]blica)?\s*(?:N[uú]mero|No\.?|N°)?\s*[-:\s]*(\d+)\b'
REGEX_ACTA_SINGLE = r'Acta\s+Fuera\s+de\s+Protocolo\s+(?:N[uú]mero|No\.?|N°)?\s*[-:\s]*(\d+)\b'
REGEX_FOLIO_DBA = r'(?i)\bSerie\s*(?:RP)?\s*Folio\s*(\d+)\b'
REGEX_FOLIO_DBA_ALT = r'(?i)DATOS\s+CFDI.*?Folio:\s*(\d+)'
REGEX_FOLIO_TOTALNOT = r'(?i)Folio\s+interno:\s*(\w+)\b'


# --- Helper Functions ---
# (Keep existing helper functions: find_unique_output_filename, extract_text_from_pdf)
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
            full_text += page.get_text("text", flags=fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE) + "\n"
        doc.close()
        full_text = re.sub(r'[ \t]{2,}', ' ', full_text)
        full_text = re.sub(r'\n\s*\n', '\n', full_text)
        return full_text
    except Exception as e:
        print(f"Error opening or reading PDF {pdf_path}: {e}")
        return None


# --- Extraction Logic ---
# (Keep existing extraction logic: find_folio, find_references, process_single_pdf)
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
        contpaq_simple_pattern = r'\bFOLIO:\s*(\w+)\b'
        candidate_folio = None
        for match in re.finditer(contpaq_simple_pattern, text, re.IGNORECASE):
            start_index = match.start()
            lookback_chars = 30
            preceding_text = text[max(0, start_index - lookback_chars) : start_index]
            if not re.search(r'\bFolio\s+fiscal\b', preceding_text, re.IGNORECASE):
                candidate_folio = match.group(1)
        folio = candidate_folio

    if folio and len(folio) > 20:
         return "FOLIO_FISCAL_SUSPECTED"
    elif not folio:
        return "NOT_FOUND"
    return folio

def find_references(text):
    """Finds all Escritura and Acta references, handling single, range, list, and special formats."""
    references = []
    if not text: return []
    found_escritura_numbers = set()
    found_acta_numbers = set()
    flags = re.IGNORECASE | re.UNICODE

    # Ranges
    for match in re.finditer(REGEX_ESCRITURA_RANGE, text, flags):
        try:
            start, end = int(match.group(1)), int(match.group(2))
            if start <= end:
                for num in range(start, end + 1):
                    num_str = str(num)
                    if num_str not in found_escritura_numbers:
                        references.append({"Type": "Escritura", "Number": num_str})
                        found_escritura_numbers.add(num_str)
        except ValueError: pass
    for match in re.finditer(REGEX_ACTA_RANGE, text, flags):
        try:
            start, end = int(match.group(1)), int(match.group(2))
            if start <= end:
                for num in range(start, end + 1):
                    num_str = str(num)
                    if num_str not in found_acta_numbers:
                        references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
                        found_acta_numbers.add(num_str)
        except ValueError: pass

    # Lists ('Y')
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

    # Special Acta
    for match in re.finditer(REGEX_ACTA_SPECIAL, text, flags):
        num_str = match.group(1).strip()
        if num_str and num_str not in found_acta_numbers:
            references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
            found_acta_numbers.add(num_str)

    # Singles
    potential_escritura_singles = []
    for match in re.finditer(REGEX_ESCRITURA_SINGLE, text, flags):
         num_str = match.group(1).strip()
         if num_str: potential_escritura_singles.append(num_str)
    potential_acta_singles = []
    for match in re.finditer(REGEX_ACTA_SINGLE, text, flags):
        num_str = match.group(1).strip()
        if num_str: potential_acta_singles.append(num_str)

    for num_str in potential_escritura_singles:
        if num_str not in found_escritura_numbers:
            references.append({"Type": "Escritura", "Number": num_str})
            found_escritura_numbers.add(num_str)
    for num_str in potential_acta_singles:
        if num_str not in found_acta_numbers:
            references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
            found_acta_numbers.add(num_str)

    # Sorting
    def sort_key(item):
        try: num_val = int(item["Number"])
        except ValueError: num_val = float('inf')
        return (item["Type"], num_val, item["Number"])
    references.sort(key=sort_key)
    return references

def process_single_pdf(pdf_path, invoice_type):
    """Processes a single PDF to extract folio and references. Returns full path."""
    text = extract_text_from_pdf(pdf_path)
    if not text:
        return [{"Document Type": "ERROR", "Reference Number": "Text Extraction Failed",
                 "Invoice Folio": "ERROR", "Source PDF": os.path.basename(pdf_path),
                 "Full PDF Path": pdf_path}]

    folio = find_folio(text, invoice_type)
    references = find_references(text)
    output_rows = []

    if not references:
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
# (Modified to store the output filename in the app instance and enable the Save button)
def run_processing(folder_path, invoice_type, app_instance):
    """Iterates through folder, processes PDFs, saves to Excel, and updates GUI."""
    all_data = []
    pdf_files = []
    output_filename = None # Initialize

    # Find all PDF files recursively
    try:
        app_instance.master.after(0, app_instance.update_status, f"Scanning folder for PDF files: {folder_path}")
        for root, _, files in os.walk(folder_path):
            for file in files:
                if not file.startswith('.') and file.lower().endswith('.pdf'):
                     pdf_path = os.path.join(root, file)
                     if os.path.isfile(pdf_path):
                         pdf_files.append(pdf_path)
    except Exception as e:
        # (Error handling remains the same)
        error_msg = f"Error accessing folder structure in {folder_path}: {e}"
        print(error_msg)
        app_instance.master.after(0, app_instance.update_status, f"Error accessing folder. Check permissions or path.")
        app_instance.master.after(0, messagebox.showerror, "Folder Error", error_msg)
        app_instance.master.after(0, app_instance.enable_buttons)
        app_instance.master.after(10, app_instance.clear_pdf_preview, "Error during folder scan.")
        return

    total_files = len(pdf_files)
    if total_files == 0:
        # (Handling no files remains the same)
        app_instance.master.after(0, app_instance.update_status, "No PDF files found in the selected folder or subfolders.")
        app_instance.master.after(0, app_instance.enable_buttons)
        app_instance.master.after(0, messagebox.showinfo, "Information", "No PDF files found.")
        app_instance.master.after(10, app_instance.clear_pdf_preview, "No PDFs found to process.")
        return

    start_time = time.time()
    files_processed_count = 0
    files_with_errors = 0

    # Process each found PDF
    for i, pdf_path in enumerate(pdf_files):
        if i % 5 == 0 or i == total_files - 1:
            status_message = f"Processing file {i+1}/{total_files}: {os.path.basename(pdf_path)}"
            app_instance.master.after(0, app_instance.update_status, status_message)
        try:
            # (Processing logic remains the same)
            results = process_single_pdf(pdf_path, invoice_type)
            if results and results[0].get("Document Type") == "ERROR" and "Extraction Failed" in results[0].get("Reference Number", ""):
                 files_with_errors += 1
                 all_data.extend(results)
            elif results:
                 files_processed_count += 1
                 all_data.extend(results)
            else:
                 files_with_errors += 1
                 print(f"Warning: No data returned by process_single_pdf for {os.path.basename(pdf_path)}")
                 all_data.append({
                      "Document Type": "ERROR", "Reference Number": "Processing Function Failed",
                      "Invoice Folio": "ERROR", "Source PDF": os.path.basename(pdf_path),
                      "Full PDF Path": pdf_path})
        except Exception as e:
            # (Error handling remains the same)
            files_with_errors += 1
            error_msg = f"CRITICAL ERROR processing file {os.path.basename(pdf_path)}: {e}"
            print(error_msg)
            import traceback; traceback.print_exc()
            app_instance.master.after(0, app_instance.update_status, f"ERROR processing {os.path.basename(pdf_path)}. See console.")
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
        # (Handling no data remains the same)
        final_message = f"Processing complete. {final_summary}\nNo data extracted."
        app_instance.master.after(0, app_instance.update_status, final_message)
        app_instance.master.after(0, app_instance.enable_buttons)
        app_instance.master.after(0, messagebox.showinfo, "Information", final_message)
        app_instance.master.after(10, app_instance.clear_pdf_preview, "Processing complete. No data.")
        return

    # Create DataFrame
    try:
        df = pd.DataFrame(all_data)
        excel_columns = ["Source PDF", "Invoice Folio", "Document Type", "Reference Number"]
        all_columns_ordered = ["Source PDF", "Invoice Folio", "Document Type", "Reference Number", "Full PDF Path"]
        for col in all_columns_ordered:
            if col not in df.columns: df[col] = None
        df = df[all_columns_ordered]

        # Sort DataFrame (logic remains the same)
        try:
            df['Reference Number Num'] = pd.to_numeric(df['Reference Number'], errors='coerce')
            df.sort_values(by=["Source PDF", "Document Type", "Reference Number Num", "Reference Number"],
                           inplace=True, na_position='last')
            df.drop(columns=['Reference Number Num'], inplace=True)
        except Exception as sort_e:
            print(f"Warning: Could not perform detailed sort on DataFrame: {sort_e}")
            df.sort_values(by=["Source PDF"], inplace=True, na_position='last')

    except Exception as e:
        # (Error handling remains the same)
        error_msg = f"Error creating or sorting DataFrame: {e}"
        print(error_msg); import traceback; traceback.print_exc()
        app_instance.master.after(0, app_instance.update_status, f"Error preparing data. See console.")
        app_instance.master.after(0, messagebox.showerror, "DataFrame Error", error_msg)
        app_instance.master.after(0, app_instance.enable_buttons)
        app_instance.master.after(10, app_instance.clear_pdf_preview, "Error creating data.")
        return

    # Generate unique output filename
    output_filename = find_unique_output_filename("Extracted_Invoices.xlsx")

    # Save to Excel
    try:
        df_to_save = df[excel_columns].copy()
        df_to_save.to_excel(output_filename, index=False, engine='openpyxl')
        final_message = f"Processing complete. {final_summary}\nData saved to:\n{output_filename}"

        # Schedule GUI updates on the main thread
        app_instance.master.after(0, app_instance.update_status, final_message)
        # *** Store the DataFrame and filename, then load to treeview ***
        app_instance.master.after(0, app_instance.set_data_and_file, df, output_filename)
        app_instance.master.after(10, app_instance.load_data_to_treeview) # Load after setting df
        app_instance.master.after(20, messagebox.showinfo, "Success", final_message)
        app_instance.master.after(30, app_instance.clear_pdf_preview, "Select a row above to preview PDF")

    except PermissionError:
        error_message = f"Error saving Excel file:\n{output_filename}\n\nPermission denied. The file might be open..."
        print(error_message)
        app_instance.master.after(0, app_instance.update_status, f"Error saving file: Permission Denied. Data NOT saved initially.")
        app_instance.master.after(0, messagebox.showerror, "Save Error", error_message)
        # Load data to treeview even if saving failed, but DON'T set the excel filename
        app_instance.master.after(10, app_instance.set_data_and_file, df, None) # Pass None for filename
        app_instance.master.after(20, app_instance.load_data_to_treeview)
        app_instance.master.after(30, app_instance.clear_pdf_preview, "Initial save failed. Edits cannot be saved.")
    except Exception as e:
        error_message = f"Error saving Excel file '{output_filename}':\n{e}"
        print(error_message); import traceback; traceback.print_exc()
        app_instance.master.after(0, app_instance.update_status, f"Error saving file. See console. Data NOT saved initially.")
        app_instance.master.after(0, messagebox.showerror, "Save Error", error_message)
        # Load data to treeview even if saving failed, but DON'T set the excel filename
        app_instance.master.after(10, app_instance.set_data_and_file, df, None) # Pass None for filename
        app_instance.master.after(20, app_instance.load_data_to_treeview)
        app_instance.master.after(30, app_instance.clear_pdf_preview, "Initial save failed. Edits cannot be saved.")

    # Always re-enable buttons
    app_instance.master.after(40, app_instance.enable_buttons)


# --- GUI Class ---
class InvoiceProcessorApp:
    def __init__(self, master):
        self.master = master
        master.title("Invoice PDF Processor v2.2 - Editable Preview") # Version bump
        master.geometry("1400x850")

        self.selected_folder = tk.StringVar(value="No folder selected")
        self.processing_active = False
        self.dataframe = None # Holds the pandas DataFrame
        self.current_excel_file = None # Holds the path to the generated excel
        self.pdf_preview_image = None
        self._canvas_image_id = None
        self._pdf_path_map = {}
        self._placeholder_window_id = None
        self._edit_entry = None # Holds the temporary Entry widget for editing
        self._edit_item_id = None # Holds the item ID being edited
        self._edit_column_id = None # Holds the column # being edited

        # --- Zoom State ---
        self.current_zoom_factor = 1.0
        self.zoom_step = 1.2
        self.min_zoom = 0.1
        self.max_zoom = 5.0
        self.current_preview_pdf_path = None

        # --- Configure Styles ---
        style = ttk.Style(master) # Use master here
        style.theme_use('clam')
        style.configure('TButton', padding=(10, 5), font=('Segoe UI', 10))
        style.map('TButton', background=[('active', '#e0e0e0')])
        style.configure('Zoom.TButton', padding=(5, 2), font=('Segoe UI', 9))
        style.configure('TLabel', padding=5, font=('Segoe UI', 10))
        style.configure('Header.TLabel', font=('Segoe UI', 11, 'bold'))
        style.configure('TEntry', padding=5, font=('Segoe UI', 10))
        style.configure("Treeview.Heading", font=('Segoe UI', 10, 'bold'), padding=5)
        style.configure("Treeview", rowheight=25, font=('Segoe UI', 9))
        style.configure("Placeholder.TLabel", foreground="grey", background="lightgrey", padding=10, anchor=tk.CENTER, font=('Segoe UI', 11, 'italic'))

        # --- Top Bar: Folder Selection & Processing Controls ---
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
        self.dba_button.pack(side=tk.LEFT, padx=5, ipady=2)
        self.totalnot_button = ttk.Button(button_frame, text="TOTALNOT", command=lambda: self.start_processing('TOTALNOT'), width=12)
        self.totalnot_button.pack(side=tk.LEFT, padx=5, ipady=2)
        self.contpaq_button = ttk.Button(button_frame, text="CONTPAQ", command=lambda: self.start_processing('CONTPAQ'), width=12)
        self.contpaq_button.pack(side=tk.LEFT, padx=5, ipady=2)

        # --- Main Content Area: Paned Window ---
        self.content_pane = ttk.PanedWindow(master, orient=tk.HORIZONTAL)
        self.content_pane.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(5, 5))

        # --- Left Panel: Excel Data Preview (Now Editable) ---
        tree_frame = ttk.Frame(self.content_pane, padding=5)
        self.content_pane.add(tree_frame, weight=2)

        # Header Frame for Treeview (includes Save button)
        tree_header_frame = ttk.Frame(tree_frame)
        tree_header_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(tree_header_frame, text="Extracted Data (Double-click cell to edit):", style='Header.TLabel').pack(side=tk.LEFT, anchor=tk.W)
        self.save_changes_button = ttk.Button(tree_header_frame, text="Save Changes to Excel", command=self.save_changes_to_excel, state=tk.DISABLED)
        self.save_changes_button.pack(side=tk.RIGHT, padx=5)

        # Treeview and Scrollbars
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
        self.tree.column("Invoice Folio", anchor=tk.W, width=100) # Allow stretch
        self.tree.column("Document Type", anchor=tk.W, width=150) # Allow stretch
        self.tree.column("Reference Number", anchor=tk.W, width=120) # Allow stretch

        # Define which columns are editable (use the heading text)
        self.editable_columns = ["Invoice Folio", "Document Type", "Reference Number"]

        self.tree.bind('<<TreeviewSelect>>', self.on_treeview_select)
        # *** Bind Double Click for Editing ***
        self.tree.bind('<Double-1>', self.on_tree_double_click)
        self._tree_sort_order = {col: False for col in self.tree["columns"]}

        # --- Right Panel: PDF Preview ---
        # (Layout remains the same)
        pdf_preview_frame = ttk.Frame(self.content_pane, padding=5)
        self.content_pane.add(pdf_preview_frame, weight=3)
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
        self.pdf_placeholder_label = ttk.Label(self.pdf_canvas, text="Select a row above to preview PDF", style="Placeholder.TLabel")
        self.pdf_canvas.bind('<Configure>', self._center_placeholder)

        # --- Bottom Bar: Status Log ---
        # (Layout remains the same)
        status_frame = ttk.Frame(master, padding="10 5 10 10")
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Label(status_frame, text="Status Log:", style='Header.TLabel').pack(anchor=tk.W)
        text_scroll_frame = ttk.Frame(status_frame)
        text_scroll_frame.pack(fill=tk.X, expand=False, pady=(5,0))
        scrollbar_status = ttk.Scrollbar(text_scroll_frame)
        scrollbar_status.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text = tk.Text(text_scroll_frame, height=6, wrap=tk.WORD, state=tk.DISABLED, relief=tk.FLAT, borderwidth=0, yscrollcommand=scrollbar_status.set, font=("Consolas", 9), background="#f0f0f0")
        self.status_text.pack(side=tk.LEFT, fill=tk.X, expand=True)
        scrollbar_status.config(command=self.status_text.yview)

        # --- Initial Setup ---
        self.clear_pdf_preview("Select a row above to preview PDF")
        self.update_status("Ready. Please select a folder and invoice type.")
        if not PIL_AVAILABLE:
             self.update_status("WARNING: Pillow library not found. PDF Preview is disabled.")
             self.clear_pdf_preview("PDF Preview disabled:\nPillow library not installed.")

    # --- GUI Methods ---

    def _center_placeholder(self, event=None):
        # (Remains the same)
        if self._placeholder_window_id and self.pdf_canvas.winfo_exists() and \
           self._placeholder_window_id in self.pdf_canvas.find_withtag("placeholder"):
            canvas_w = self.pdf_canvas.winfo_width()
            canvas_h = self.pdf_canvas.winfo_height()
            self.pdf_canvas.coords(self._placeholder_window_id, canvas_w//2, canvas_h//2)

    def select_folder(self):
        if self.processing_active: return
        # If data exists, maybe ask if they want to save changes? For now, just clear.
        if self.dataframe is not None:
            # Optional: Add a confirmation dialog here if edits might be lost.
            pass

        folder = filedialog.askdirectory()
        if folder:
            normalized_folder = os.path.normpath(folder)
            self.selected_folder.set(normalized_folder)
            self.clear_status()
            self.clear_treeview() # Also clears dataframe and excel file path
            self.clear_pdf_preview("Select a row to preview PDF")
            self.current_zoom_factor = 1.0
            self.current_preview_pdf_path = None
            self.update_status(f"Folder selected: {normalized_folder}")
            self.update_status("Ready to process.")
            # Ensure save button is disabled
            self.save_changes_button.config(state=tk.DISABLED)
        else:
            if self.selected_folder.get() != "No folder selected":
                self.update_status("Folder selection cancelled.")

    def clear_status(self):
         # (Remains the same)
         self.status_text.config(state=tk.NORMAL)
         self.status_text.delete('1.0', tk.END)
         self.status_text.config(state=tk.DISABLED)

    def update_status(self, message):
        # (Remains the same)
        if threading.current_thread() != threading.main_thread():
            self.master.after(0, self._update_status_text, message)
        else:
            self._update_status_text(message)

    def _update_status_text(self, message):
        # (Remains the same)
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
        self.zoom_in_button.config(state=tk.DISABLED)
        self.zoom_out_button.config(state=tk.DISABLED)
        self.reset_zoom_button.config(state=tk.DISABLED)
        # Disable save button during processing
        self.save_changes_button.config(state=tk.DISABLED)
        # Cancel any ongoing cell edit
        self._cancel_cell_edit()


    def enable_buttons(self):
        self.processing_active = False
        self.select_button.config(state=tk.NORMAL)
        self.dba_button.config(state=tk.NORMAL)
        self.totalnot_button.config(state=tk.NORMAL)
        self.contpaq_button.config(state=tk.NORMAL)
        self.zoom_in_button.config(state=tk.NORMAL)
        self.zoom_out_button.config(state=tk.NORMAL)
        self.reset_zoom_button.config(state=tk.NORMAL)
        # Enable save button ONLY if an excel file path exists
        if self.current_excel_file and self.dataframe is not None:
            self.save_changes_button.config(state=tk.NORMAL)
        else:
            self.save_changes_button.config(state=tk.DISABLED)


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
        self.clear_treeview() # Clear previous results and disable save button
        self.clear_pdf_preview(f"Processing {invoice_type} invoices...\nPlease wait.")
        self.update_status(f"Starting recursive processing for {invoice_type} in: {folder}")
        self.update_status("-" * 40)

        process_thread = threading.Thread(target=run_processing,
                                          args=(folder, invoice_type, self),
                                          daemon=True)
        process_thread.start()

    # --- Data Handling Methods ---
    def set_data_and_file(self, df, excel_file_path):
        """Stores the processed DataFrame and the path to the saved Excel file."""
        self.dataframe = df
        self.current_excel_file = excel_file_path
        # Enable save button if path is valid
        if self.current_excel_file and self.dataframe is not None:
             # Schedule button enable on main thread if needed, but likely called via after() already
             self.save_changes_button.config(state=tk.NORMAL)
        else:
             self.save_changes_button.config(state=tk.DISABLED)


    def clear_treeview(self):
        """Clears Treeview, DataFrame, path map, and resets Excel file path."""
        if self._edit_entry: self._cancel_cell_edit() # Cancel edit before clearing
        if hasattr(self, 'tree'):
            for item in self.tree.get_children():
                self.tree.delete(item)
        self.dataframe = None
        self.current_excel_file = None # Reset file path
        self._pdf_path_map.clear()
        # Ensure save button is disabled
        if hasattr(self, 'save_changes_button'): # Check if button exists yet
            self.save_changes_button.config(state=tk.DISABLED)

    def load_data_to_treeview(self):
        """Loads data from the internal DataFrame into the Treeview."""
        # Clear previous items first (without clearing dataframe/excel path this time)
        if self._edit_entry: self._cancel_cell_edit()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._pdf_path_map.clear()

        if self.dataframe is None or self.dataframe.empty:
             # self.update_status("No data available to display in the table.") # Already handled elsewhere
             self.clear_pdf_preview("No data loaded.")
             self.save_changes_button.config(state=tk.DISABLED) # Ensure disabled
             return

        display_columns = list(self.tree["columns"])
        required_cols = display_columns + ["Full PDF Path"]
        if not all(col in self.dataframe.columns for col in required_cols):
            # (Error handling remains the same)
            missing = [col for col in required_cols if col not in self.dataframe.columns]
            errmsg = f"Error: DataFrame is missing required columns: {missing}"
            self.update_status(errmsg)
            messagebox.showerror("Data Loading Error", errmsg, parent=self.master)
            print(f"DataFrame columns available: {self.dataframe.columns.tolist()}")
            self.clear_pdf_preview("Error loading data. See status log.")
            self.dataframe = None # Invalidate data
            self.current_excel_file = None
            self.save_changes_button.config(state=tk.DISABLED)
            return

        self.tree.configure(displaycolumns=display_columns)
        # Use DataFrame index as key for mapping item_id back, more reliable than row order
        self._df_index_map = {}

        for df_index, row in self.dataframe.iterrows():
            try:
                display_values = tuple(str(row[col]) if pd.notna(row[col]) else "" for col in display_columns)
                full_path = row["Full PDF Path"]
                # Insert using df_index as iid for easy lookup later if needed, though we mainly use item_id
                item_id = self.tree.insert("", tk.END, values=display_values, iid=df_index)
                if full_path and isinstance(full_path, str):
                    self._pdf_path_map[item_id] = full_path
                # Store mapping from tree item ID back to DataFrame index
                self._df_index_map[item_id] = df_index
            except Exception as e:
                print(f"Error adding row index {df_index} to treeview: {e}")
                self.update_status(f"Warning: Could not display row for PDF '{row.get('Source PDF', 'Unknown')}' in table.")

        row_count = len(self.tree.get_children())
        # self.update_status(f"Loaded {row_count} rows into the table.") # Can be verbose
        if row_count == 0:
             self.update_status("Loaded data frame, but no rows were added to the table.")
             self.clear_pdf_preview("Data loaded, but no rows to display.")

        # Enable save button if we have data and a file path
        if self.current_excel_file and not self.dataframe.empty:
             self.save_changes_button.config(state=tk.NORMAL)
        else:
             self.save_changes_button.config(state=tk.DISABLED)


    def sort_treeview_column(self, col, reverse):
        """Sorts the treeview column based on the clicked header, updates DataFrame order."""
        if self.dataframe is None or self.dataframe.empty: return
        if self._edit_entry: self._cancel_cell_edit() # Cancel edit before sorting

        try:
            # Sort the DataFrame
            sorted_df = self.dataframe.copy()
            try:
                numeric_col = pd.to_numeric(sorted_df[col], errors='coerce')
                is_numeric = not numeric_col.isna().all()
            except Exception: is_numeric = False

            if is_numeric:
                # Use argsort to get indices, then iloc for stable sort wrt index
                sort_indices = numeric_col.argsort(na_position='last')
                if reverse: sort_indices = sort_indices[::-1]
                sorted_df = sorted_df.iloc[sort_indices]
            else:
                 # Use sort_values for string/mixed types
                 sorted_df = sorted_df.sort_values(by=col, ascending=not reverse, key=lambda x: x.astype(str).str.lower(), na_position='last')

            # Update the internal DataFrame to reflect the new sort order
            self.dataframe = sorted_df

            # Update header indicators
            for c in self.tree["columns"]:
                current_text = self.tree.heading(c, "text").replace(' ▲', '').replace(' ▼', '')
                if c == col:
                    indicator = ' ▲' if not reverse else ' ▼'
                    self.tree.heading(c, text=current_text + indicator)
                else:
                    self.tree.heading(c, text=current_text)

            # Reload the treeview from the now-sorted DataFrame
            self.load_data_to_treeview()
            self._tree_sort_order[col] = not reverse

        except KeyError:
             print(f"Error: Column '{col}' not found in DataFrame for sorting.")
             messagebox.showerror("Sort Error", f"Column '{col}' not found.", parent=self.master)
        except Exception as e:
             print(f"Error sorting treeview column '{col}': {e}")
             import traceback; traceback.print_exc()
             messagebox.showerror("Sort Error", f"Could not sort column '{col}':\n{e}", parent=self.master)

    # --- Treeview Editing Methods ---

    def on_tree_double_click(self, event):
        """Handle double-click on a treeview cell to initiate editing."""
        if self._edit_entry: # If already editing, do nothing
            return

        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return # Click was not on a cell

        item_id = self.tree.identify_row(event.y)
        column_id_str = self.tree.identify_column(event.x) # e.g., '#2'

        if not item_id or not column_id_str: return # Clicked outside actual data rows/cols

        # Convert column ID string to integer index (0-based) and get column header name
        try:
            column_index = int(column_id_str.replace('#', '')) - 1
            column_name = self.tree["columns"][column_index]
        except (ValueError, IndexError):
            print(f"Error identifying column from ID: {column_id_str}")
            return

        # Check if this column is designated as editable
        if column_name not in self.editable_columns:
            # print(f"Column '{column_name}' is not editable.")
            return

        # Get the bounding box of the cell
        try:
             bbox = self.tree.bbox(item_id, column=column_id_str)
             if not bbox: return # Cell not visible or invalid
        except Exception:
             return # Error getting bbox

        # Get current value
        current_values = self.tree.item(item_id, 'values')
        original_value = current_values[column_index]

        # Create and place the Entry widget
        self._edit_item_id = item_id
        self._edit_column_id = column_id_str # Store '#n' format

        self._edit_entry = ttk.Entry(self.tree, font=('Segoe UI', 9))
        self._edit_entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])

        self._edit_entry.insert(0, original_value)
        self._edit_entry.select_range(0, tk.END)
        self._edit_entry.focus_set()

        # Bind events to the Entry widget
        self._edit_entry.bind("<Return>", lambda e: self._save_cell_edit(item_id, column_index, column_name))
        self._edit_entry.bind("<KP_Enter>", lambda e: self._save_cell_edit(item_id, column_index, column_name)) # Numpad Enter
        self._edit_entry.bind("<Escape>", lambda e: self._cancel_cell_edit())
        self._edit_entry.bind("<FocusOut>", lambda e: self._save_cell_edit(item_id, column_index, column_name)) # Save on clicking away

    def _save_cell_edit(self, item_id, column_index, column_name):
        """Save the edited value from the entry widget to Treeview and DataFrame."""
        if not self._edit_entry: return # No edit in progress

        new_value = self._edit_entry.get()
        self._cancel_cell_edit() # Destroy the entry widget first

        # Get current values and update the specific one
        current_values = list(self.tree.item(item_id, 'values'))
        if column_index < len(current_values):
             # Only update if the value actually changed
            if str(current_values[column_index]) != new_value:
                current_values[column_index] = new_value
                self.tree.item(item_id, values=tuple(current_values))

                # --- Update the DataFrame ---
                if self.dataframe is not None:
                    try:
                        # Find the DataFrame index corresponding to the treeview item_id
                        df_index = self._df_index_map.get(item_id)
                        if df_index is not None and df_index in self.dataframe.index:
                            # Use .loc for safe setting
                            self.dataframe.loc[df_index, column_name] = new_value
                            # print(f"DataFrame updated: Index {df_index}, Column '{column_name}', New Value '{new_value}'")
                            self.update_status(f"Cell updated: {column_name} = '{new_value}'")
                             # Re-enable save button if it was disabled (e.g., if data was empty before)
                            if self.current_excel_file:
                                 self.save_changes_button.config(state=tk.NORMAL)
                        else:
                            print(f"Error: Could not find DataFrame index for tree item {item_id}")
                            self.update_status(f"Error: Failed to update backing data for edit.")

                    except KeyError:
                        print(f"Error: Column '{column_name}' or index {df_index} not found in DataFrame during update.")
                        self.update_status(f"Error: Failed to update backing data for edit (KeyError).")
                    except Exception as e:
                         print(f"Error updating DataFrame: {e}")
                         self.update_status(f"Error: Failed to update backing data for edit ({type(e).__name__}).")
            # else:
                # print("No change detected, not updating.")

    def _cancel_cell_edit(self, event=None):
        """Cancel the current cell edit and destroy the entry widget."""
        if self._edit_entry:
            self._edit_entry.destroy()
            self._edit_entry = None
            self._edit_item_id = None
            self._edit_column_id = None
            # Maybe refocus the tree?
            # self.tree.focus_set()

    # --- Excel Saving Method ---
    def save_changes_to_excel(self):
        """Saves the current state of the DataFrame back to the Excel file."""
        if self._edit_entry: self._cancel_cell_edit() # Ensure any active edit is finished/cancelled

        if self.dataframe is None or self.dataframe.empty:
            messagebox.showwarning("No Data", "There is no data to save.", parent=self.master)
            return

        if not self.current_excel_file:
            messagebox.showerror("Save Error", "The original Excel file path is unknown. Cannot save changes.\n(Did the initial save fail?)", parent=self.master)
            return

        if not os.path.exists(os.path.dirname(self.current_excel_file)):
             messagebox.showerror("Save Error", f"The directory for the Excel file no longer exists:\n{os.path.dirname(self.current_excel_file)}", parent=self.master)
             return

        # Confirm overwrite
        if not messagebox.askyesno("Confirm Save", f"This will overwrite the existing file:\n{self.current_excel_file}\n\nAre you sure you want to save the changes?", parent=self.master):
            self.update_status("Save cancelled by user.")
            return

        self.update_status(f"Attempting to save changes to: {self.current_excel_file}")
        try:
            # Select only the columns originally intended for Excel export
            excel_columns = ["Source PDF", "Invoice Folio", "Document Type", "Reference Number"]
            # Ensure these columns exist before trying to slice
            cols_to_save = [col for col in excel_columns if col in self.dataframe.columns]
            if not cols_to_save:
                 raise ValueError("No valid columns found in the DataFrame to save.")

            df_to_save = self.dataframe[cols_to_save].copy()

            # Save to the *stored* Excel file path
            df_to_save.to_excel(self.current_excel_file, index=False, engine='openpyxl')

            success_msg = f"Changes successfully saved to:\n{self.current_excel_file}"
            self.update_status(success_msg)
            messagebox.showinfo("Save Successful", success_msg, parent=self.master)

        except PermissionError:
            error_message = f"Error saving changes to:\n{self.current_excel_file}\n\nPermission denied. The file might be open in another application.\nPlease close it and try saving again."
            print(error_message)
            self.update_status(f"Save Error: Permission Denied.")
            messagebox.showerror("Save Error", error_message, parent=self.master)
        except Exception as e:
            error_message = f"An unexpected error occurred while saving changes to '{self.current_excel_file}':\n{e}"
            print(error_message); import traceback; traceback.print_exc()
            self.update_status(f"Save Error: {e}. See console.")
            messagebox.showerror("Save Error", error_message, parent=self.master)


    def on_treeview_select(self, event):
        """Handles selection changes in the Treeview to update the PDF preview."""
        # Cancel any ongoing edit when selection changes
        if self._edit_entry and self.tree.focus() != self._edit_entry:
             # Check if focus moved away from the entry due to selection change
             # Find the column index/name being edited
             if self._edit_item_id and self._edit_column_id:
                  try:
                       col_idx = int(self._edit_column_id.replace('#', '')) - 1
                       col_name = self.tree["columns"][col_idx]
                       self._save_cell_edit(self._edit_item_id, col_idx, col_name)
                  except: # Fallback if getting details fails
                       self._cancel_cell_edit()
             else: # Fallback
                  self._cancel_cell_edit()


        selected_items = self.tree.selection()
        if not selected_items: return

        selected_item_id = selected_items[0]
        pdf_full_path = self._pdf_path_map.get(selected_item_id)

        if pdf_full_path and isinstance(pdf_full_path, str):
            if os.path.exists(pdf_full_path):
                self.update_pdf_preview(pdf_full_path)
            else:
                 # (Error handling remains the same)
                 base_name = os.path.basename(pdf_full_path)
                 err_msg = f"File Not Found:\n{base_name}\n(Path: {pdf_full_path})"
                 self.clear_pdf_preview(err_msg)
                 print(f"Error: File path from selection '{pdf_full_path}' does not exist.")
                 self.update_status(f"Preview Error: File not found - {base_name}")
                 self.current_preview_pdf_path = None
        elif pdf_full_path is None:
            # (Error handling remains the same)
            print(f"Error: No path found in map for selected item ID: {selected_item_id}")
            self.clear_pdf_preview("Error: Could not retrieve\nfile path for this row.")
            self.current_preview_pdf_path = None
        else:
             # (Error handling remains the same)
             print(f"Error: Invalid path data associated with selected row: {pdf_full_path}")
             self.clear_pdf_preview("Error: Invalid file path\nin selected row data.")
             self.current_preview_pdf_path = None


    # --- PDF Preview and Zoom Methods ---
    # (clear_pdf_preview, update_pdf_preview, zoom_in, zoom_out, reset_zoom, _update_canvas_scrollregion methods remain unchanged)
    def clear_pdf_preview(self, message="Select a row to preview PDF"):
        """Clears the PDF preview area, shows placeholder, and resets path."""
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
            canvas_w//2, canvas_h//2, window=self.pdf_placeholder_label, anchor=tk.CENTER, tags="placeholder"
        )
        self.master.after_idle(self._update_canvas_scrollregion)
        self.current_preview_pdf_path = None


    def update_pdf_preview(self, pdf_path):
        """Renders PDF page using current zoom factor and displays it."""
        if not PIL_AVAILABLE:
             self.clear_pdf_preview("PDF Preview disabled:\nPillow library not found.")
             return

        # Clear existing content
        if self._placeholder_window_id:
            try: self.pdf_canvas.delete("placeholder")
            except tk.TclError: pass
            self._placeholder_window_id = None
        if self._canvas_image_id:
            try: self.pdf_canvas.delete(self._canvas_image_id)
            except tk.TclError: pass
            self._canvas_image_id = None
        self.pdf_preview_image = None

        try:
            doc = fitz.open(pdf_path)
            if len(doc) == 0: raise ValueError("PDF has no pages.")
            page = doc.load_page(0)
            page_rect = page.rect
            if page_rect.width == 0 or page_rect.height == 0: raise ValueError("PDF page has zero dimensions.")

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
                self._canvas_image_id = self.pdf_canvas.create_image(0, 0, anchor=tk.NW, image=self.pdf_preview_image, tags="pdf_image")
                self.current_preview_pdf_path = pdf_path
                self.master.after_idle(self._update_canvas_scrollregion)
                self.pdf_canvas.xview_moveto(0); self.pdf_canvas.yview_moveto(0)
            else:
                 raise ValueError("Failed to create PhotoImage object.")

        except (fitz.fitz.FileNotFoundError, fitz.fitz.PasswordError, ValueError, Exception) as e:
            base_name = os.path.basename(pdf_path) if pdf_path else "Unknown File"
            error_type = type(e).__name__
            error_msg = f"Preview Error ({error_type}):\n{base_name}"
            if isinstance(e, fitz.fitz.PasswordError): error_msg += "\n(Password Protected?)"
            elif isinstance(e, ValueError): error_msg += f"\n({e})"
            import traceback; print(f"--- PDF Preview Exception Traceback for {base_name} ---"); traceback.print_exc(); print(f"--- End Traceback ---")
            self.clear_pdf_preview(error_msg)

    def zoom_in(self):
        if self.current_preview_pdf_path:
            new_zoom = min(self.current_zoom_factor * self.zoom_step, self.max_zoom)
            if new_zoom != self.current_zoom_factor:
                self.current_zoom_factor = new_zoom
                self.update_pdf_preview(self.current_preview_pdf_path)

    def zoom_out(self):
        if self.current_preview_pdf_path:
            new_zoom = max(self.current_zoom_factor / self.zoom_step, self.min_zoom)
            if new_zoom != self.current_zoom_factor:
                self.current_zoom_factor = new_zoom
                self.update_pdf_preview(self.current_preview_pdf_path)

    def reset_zoom(self):
        if self.current_preview_pdf_path and self.current_zoom_factor != 1.0:
            self.current_zoom_factor = 1.0
            self.update_pdf_preview(self.current_preview_pdf_path)

    def _update_canvas_scrollregion(self):
        try:
            if self._canvas_image_id and self.pdf_canvas.find_withtag(self._canvas_image_id):
                bbox = self.pdf_canvas.bbox(self._canvas_image_id)
                if bbox:
                    scroll_bbox = (bbox[0], bbox[1], bbox[2] + 5, bbox[3] + 5)
                    self.pdf_canvas.config(scrollregion=scroll_bbox)
                    return
            current_width = self.pdf_canvas.winfo_width()
            current_height = self.pdf_canvas.winfo_height()
            self.pdf_canvas.config(scrollregion=(0, 0, max(1, current_width), max(1, current_height)))
        except tk.TclError as e: print(f"Warning: TclError updating canvas scrollregion: {e}")
        except Exception as e: print(f"Error updating canvas scrollregion: {e}")


# --- Main Execution ---
if __name__ == "__main__":
    if not PIL_AVAILABLE:
         # (Warning message remains the same)
         root_check = tk.Tk(); root_check.withdraw()
         messagebox.showwarning("Dependency Warning", "Python Imaging Library (Pillow) not found.\nPDF Preview functionality will be disabled.\n\nPlease install it using pip:\npip install Pillow", parent=None)
         root_check.destroy()

    root = tk.Tk()
    app = InvoiceProcessorApp(root)
    root.mainloop()

