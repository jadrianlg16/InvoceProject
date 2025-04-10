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



# (Keep Regex Patterns and Helper Functions as they are)
# --- Regex Patterns ---
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
# (find_folio, find_references, process_single_pdf remain unchanged)
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
    potential_escritura_singles = [m.group(1).strip() for m in re.finditer(REGEX_ESCRITURA_SINGLE, text, flags) if m.group(1)]
    potential_acta_singles = [m.group(1).strip() for m in re.finditer(REGEX_ACTA_SINGLE, text, flags) if m.group(1)]
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
# (run_processing remains unchanged - it just delivers the initial df and filename)
def run_processing(folder_path, invoice_type, app_instance):
    all_data = []
    pdf_files = []
    output_filename = None
    try:
        app_instance.master.after(0, app_instance.update_status, f"Scanning folder for PDF files: {folder_path}")
        for root, _, files in os.walk(folder_path):
            for file in files:
                if not file.startswith('.') and file.lower().endswith('.pdf'):
                     pdf_path = os.path.join(root, file)
                     if os.path.isfile(pdf_path): pdf_files.append(pdf_path)
    except Exception as e:
        error_msg = f"Error accessing folder structure in {folder_path}: {e}"
        print(error_msg)
        app_instance.master.after(0, app_instance.update_status, f"Error accessing folder. Check permissions or path.")
        app_instance.master.after(0, messagebox.showerror, "Folder Error", error_msg)
        app_instance.master.after(0, app_instance.enable_buttons)
        app_instance.master.after(10, app_instance.clear_pdf_preview, "Error during folder scan.")
        return

    total_files = len(pdf_files)
    if total_files == 0:
        app_instance.master.after(0, app_instance.update_status, "No PDF files found in the selected folder or subfolders.")
        app_instance.master.after(0, app_instance.enable_buttons)
        app_instance.master.after(0, messagebox.showinfo, "Information", "No PDF files found.")
        app_instance.master.after(10, app_instance.clear_pdf_preview, "No PDFs found to process.")
        return

    start_time = time.time(); files_processed_count = 0; files_with_errors = 0
    for i, pdf_path in enumerate(pdf_files):
        if i % 5 == 0 or i == total_files - 1:
            status_message = f"Processing file {i+1}/{total_files}: {os.path.basename(pdf_path)}"
            app_instance.master.after(0, app_instance.update_status, status_message)
        try:
            results = process_single_pdf(pdf_path, invoice_type)
            if results and results[0].get("Document Type") == "ERROR" and "Extraction Failed" in results[0].get("Reference Number", ""):
                 files_with_errors += 1; all_data.extend(results)
            elif results:
                 files_processed_count += 1; all_data.extend(results)
            else:
                 files_with_errors += 1
                 print(f"Warning: No data returned by process_single_pdf for {os.path.basename(pdf_path)}")
                 all_data.append({"Document Type": "ERROR", "Reference Number": "Processing Function Failed", "Invoice Folio": "ERROR", "Source PDF": os.path.basename(pdf_path), "Full PDF Path": pdf_path})
        except Exception as e:
            files_with_errors += 1; error_msg = f"CRITICAL ERROR processing file {os.path.basename(pdf_path)}: {e}"
            print(error_msg); import traceback; traceback.print_exc()
            app_instance.master.after(0, app_instance.update_status, f"ERROR processing {os.path.basename(pdf_path)}. See console.")
            all_data.append({"Document Type": "ERROR", "Reference Number": f"Runtime Error: {e}", "Invoice Folio": "ERROR", "Source PDF": os.path.basename(pdf_path), "Full PDF Path": pdf_path})

    end_time = time.time(); processing_time = end_time - start_time
    final_summary = f"{files_processed_count}/{total_files} files processed"
    if files_with_errors > 0: final_summary += f" ({files_with_errors} file(s) encountered errors)"
    final_summary += f" in {processing_time:.2f}s."

    if not all_data:
        final_message = f"Processing complete. {final_summary}\nNo data extracted."
        app_instance.master.after(0, app_instance.update_status, final_message)
        app_instance.master.after(0, app_instance.enable_buttons)
        app_instance.master.after(0, messagebox.showinfo, "Information", final_message)
        app_instance.master.after(10, app_instance.clear_pdf_preview, "Processing complete. No data.")
        return

    try:
        df = pd.DataFrame(all_data)
        excel_columns = ["Source PDF", "Invoice Folio", "Document Type", "Reference Number"]
        all_columns_ordered = ["Source PDF", "Invoice Folio", "Document Type", "Reference Number", "Full PDF Path"]
        for col in all_columns_ordered:
            if col not in df.columns: df[col] = None
        df = df[all_columns_ordered]
        try:
            df['Reference Number Num'] = pd.to_numeric(df['Reference Number'], errors='coerce')
            df.sort_values(by=["Source PDF", "Document Type", "Reference Number Num", "Reference Number"], inplace=True, na_position='last')
            df.drop(columns=['Reference Number Num'], inplace=True)
        except Exception as sort_e:
            print(f"Warning: Could not perform detailed sort on DataFrame: {sort_e}")
            df.sort_values(by=["Source PDF"], inplace=True, na_position='last')
    except Exception as e:
        error_msg = f"Error creating or sorting DataFrame: {e}"
        print(error_msg); import traceback; traceback.print_exc()
        app_instance.master.after(0, app_instance.update_status, f"Error preparing data. See console.")
        app_instance.master.after(0, messagebox.showerror, "DataFrame Error", error_msg)
        app_instance.master.after(0, app_instance.enable_buttons)
        app_instance.master.after(10, app_instance.clear_pdf_preview, "Error creating data.")
        return

    output_filename = find_unique_output_filename("Extracted_Invoices.xlsx")
    try:
        df_to_save = df[excel_columns].copy()
        df_to_save.to_excel(output_filename, index=False, engine='openpyxl')
        final_message = f"Processing complete. {final_summary}\nData saved to:\n{output_filename}"
        app_instance.master.after(0, app_instance.update_status, final_message)
        app_instance.master.after(0, app_instance.set_data_and_file, df, output_filename)
        app_instance.master.after(10, app_instance.load_data_to_treeview)
        app_instance.master.after(20, messagebox.showinfo, "Success", final_message)
        app_instance.master.after(30, app_instance.clear_pdf_preview, "Select a row above to preview PDF")
    except PermissionError:
        error_message = f"Error saving Excel file:\n{output_filename}\n\nPermission denied..."
        print(error_message)
        app_instance.master.after(0, app_instance.update_status, f"Error saving file: Permission Denied. Data NOT saved initially.")
        app_instance.master.after(0, messagebox.showerror, "Save Error", error_message)
        app_instance.master.after(10, app_instance.set_data_and_file, df, None)
        app_instance.master.after(20, app_instance.load_data_to_treeview)
        app_instance.master.after(30, app_instance.clear_pdf_preview, "Initial save failed. Edits cannot be saved.")
    except Exception as e:
        error_message = f"Error saving Excel file '{output_filename}':\n{e}"
        print(error_message); import traceback; traceback.print_exc()
        app_instance.master.after(0, app_instance.update_status, f"Error saving file. See console. Data NOT saved initially.")
        app_instance.master.after(0, messagebox.showerror, "Save Error", error_message)
        app_instance.master.after(10, app_instance.set_data_and_file, df, None)
        app_instance.master.after(20, app_instance.load_data_to_treeview)
        app_instance.master.after(30, app_instance.clear_pdf_preview, "Initial save failed. Edits cannot be saved.")

    app_instance.master.after(40, app_instance.enable_buttons)





# --- GUI Class ---
class InvoiceProcessorApp:
    # ... (keep __init__ and other methods as they were in the previous version) ...
    def __init__(self, master):
        self.master = master
        master.title("Invoice PDF Processor v2.3 - Row Manipulation")
        master.geometry("1400x850")

        self.selected_folder = tk.StringVar(value="No folder selected")
        self.processing_active = False
        self.dataframe = None
        self.current_excel_file = None
        self.pdf_preview_image = None
        self._canvas_image_id = None
        self._pdf_path_map = {}
        self._df_index_map = {} # Maps tree item ID -> df index
        self._placeholder_window_id = None
        self._edit_entry = None
        self._edit_item_id = None
        self._edit_column_id = None
        self._clipboard = [] # For copy/paste rows

        # --- Zoom State ---
        self.current_zoom_factor = 1.0
        self.zoom_step = 1.2
        self.min_zoom = 0.1
        self.max_zoom = 5.0
        self.current_preview_pdf_path = None

        # --- Configure Styles ---
        style = ttk.Style(master)
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


        # --- Top Bar ---
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

        # --- Main Content Area ---
        self.content_pane = ttk.PanedWindow(master, orient=tk.HORIZONTAL)
        self.content_pane.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(5, 5))

        # --- Left Panel ---
        tree_frame = ttk.Frame(self.content_pane, padding=5)
        self.content_pane.add(tree_frame, weight=2)
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
                                 selectmode='extended')
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
        self.tree.column("Invoice Folio", anchor=tk.W, width=100)
        self.tree.column("Document Type", anchor=tk.W, width=150)
        self.tree.column("Reference Number", anchor=tk.W, width=120)
        self.editable_columns = ["Invoice Folio", "Document Type", "Reference Number"]
        self.tree.bind('<<TreeviewSelect>>', self.on_treeview_select)
        self.tree.bind('<Double-1>', self.on_tree_double_click)
        self.tree.bind('<Button-3>', self.show_context_menu)
        self._tree_sort_order = {col: False for col in self.tree["columns"]}

        # --- Context Menu ---
        self.context_menu = tk.Menu(master, tearoff=0)
        self.context_menu.add_command(label="Add Blank Row", command=self._add_row)
        self.context_menu.add_command(label="Copy Selected Row(s)", command=self._copy_selected_rows)
        self.context_menu.add_command(label="Paste Row(s)", command=self._paste_rows)
        self.context_menu.add_command(label="Delete Selected Row(s)", command=self._delete_selected_rows)

        # --- Right Panel ---
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

        # --- Bottom Bar ---
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

    # ... (keep _center_placeholder, select_folder, clear_status, update_status, _update_status_text) ...
    # ... (keep disable_buttons, enable_buttons, start_processing) ...
    # ... (keep set_data_and_file, clear_treeview, load_data_to_treeview, sort_treeview_column) ...
    # ... (keep on_tree_double_click, _save_cell_edit, _cancel_cell_edit) ...
    # ... (keep show_context_menu) ... # show_context_menu logic remains the same

    def _center_placeholder(self, event=None):
        if self._placeholder_window_id and self.pdf_canvas.winfo_exists() and \
           self._placeholder_window_id in self.pdf_canvas.find_withtag("placeholder"):
            canvas_w = self.pdf_canvas.winfo_width()
            canvas_h = self.pdf_canvas.winfo_height()
            self.pdf_canvas.coords(self._placeholder_window_id, canvas_w//2, canvas_h//2)

    def select_folder(self):
        if self.processing_active: return
        if self.dataframe is not None and self.current_excel_file:
             if not messagebox.askokcancel("Confirm Folder Change", "Changing the folder will clear the current data and edits.\nAny unsaved changes will be lost.\n\nProceed?", parent=self.master):
                  return
        folder = filedialog.askdirectory()
        if folder:
            normalized_folder = os.path.normpath(folder)
            self.selected_folder.set(normalized_folder)
            self.clear_status()
            self.clear_treeview()
            self.clear_pdf_preview("Select a row to preview PDF")
            self.current_zoom_factor = 1.0
            self.current_preview_pdf_path = None
            self.update_status(f"Folder selected: {normalized_folder}")
            self.update_status("Ready to process.")
            self.save_changes_button.config(state=tk.DISABLED)
        else:
            if self.selected_folder.get() != "No folder selected":
                self.update_status("Folder selection cancelled.")

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
        self.save_changes_button.config(state=tk.DISABLED)
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
        if self.current_excel_file and self.dataframe is not None and not self.dataframe.empty:
            self.save_changes_button.config(state=tk.NORMAL)
        else:
            self.save_changes_button.config(state=tk.DISABLED)

    def start_processing(self, invoice_type):
        folder = self.selected_folder.get()
        if not folder or folder == "No folder selected": messagebox.showerror("Error", "Please select a folder first.", parent=self.master); return
        if not os.path.isdir(folder): messagebox.showerror("Error", f"Invalid directory selected:\n{folder}", parent=self.master); return
        if self.processing_active: messagebox.showwarning("Busy", "Processing is already in progress.", parent=self.master); return
        self.disable_buttons(); self.clear_status(); self.clear_treeview()
        self.clear_pdf_preview(f"Processing {invoice_type} invoices...\nPlease wait.")
        self.update_status(f"Starting recursive processing for {invoice_type} in: {folder}"); self.update_status("-" * 40)
        process_thread = threading.Thread(target=run_processing, args=(folder, invoice_type, self), daemon=True); process_thread.start()

    def set_data_and_file(self, df, excel_file_path):
        self.dataframe = df
        self.current_excel_file = excel_file_path
        self._clipboard = []
        if self.current_excel_file and self.dataframe is not None and not self.dataframe.empty:
             self.save_changes_button.config(state=tk.NORMAL)
        else:
             self.save_changes_button.config(state=tk.DISABLED)

    def clear_treeview(self):
        if self._edit_entry: self._cancel_cell_edit()
        if hasattr(self, 'tree'):
            self.tree.unbind('<Button-3>')
            for item in self.tree.get_children():
                try: self.tree.delete(item)
                except tk.TclError: pass
            self.tree.bind('<Button-3>', self.show_context_menu)
        self.dataframe = None
        self.current_excel_file = None
        self._pdf_path_map.clear()
        self._df_index_map.clear()
        self._clipboard = []
        if hasattr(self, 'save_changes_button'): self.save_changes_button.config(state=tk.DISABLED)

    def load_data_to_treeview(self):
        if self._edit_entry: self._cancel_cell_edit()
        self.tree.unbind('<Button-3>')
        for item in self.tree.get_children():
            try: self.tree.delete(item)
            except tk.TclError: pass
        self.tree.bind('<Button-3>', self.show_context_menu)
        self._pdf_path_map.clear(); self._df_index_map.clear()

        if self.dataframe is None or self.dataframe.empty:
             self.clear_pdf_preview("No data loaded."); self.save_changes_button.config(state=tk.DISABLED); return

        display_columns = list(self.tree["columns"]); required_cols = display_columns + ["Full PDF Path"]
        if not all(col in self.dataframe.columns for col in required_cols):
            missing = [col for col in required_cols if col not in self.dataframe.columns]; errmsg = f"Error: DataFrame is missing required columns: {missing}"
            self.update_status(errmsg); messagebox.showerror("Data Loading Error", errmsg, parent=self.master); print(f"DataFrame columns available: {self.dataframe.columns.tolist()}")
            self.clear_pdf_preview("Error loading data. See status log."); self.dataframe = None; self.current_excel_file = None; self._clipboard = []
            self.save_changes_button.config(state=tk.DISABLED); return

        self.tree.configure(displaycolumns=display_columns)
        last_item_id = None
        if not self.dataframe.index.is_unique:
             self.update_status("Warning: DataFrame index is not unique. Resetting index.")
             self.dataframe.reset_index(drop=True, inplace=True)

        for df_index, row in self.dataframe.iterrows():
            try:
                display_values = tuple(str(row[col]) if pd.notna(row[col]) else "" for col in display_columns)
                full_path = row["Full PDF Path"]
                item_id = str(df_index)
                self.tree.insert("", tk.END, values=display_values, iid=item_id)
                if full_path and isinstance(full_path, str): self._pdf_path_map[item_id] = full_path
                self._df_index_map[item_id] = df_index
                last_item_id = item_id
            except Exception as e:
                print(f"Error adding row index {df_index} (iid: {item_id}) to treeview: {e}")
                self.update_status(f"Warning: Could not display row for PDF '{row.get('Source PDF', 'Unknown')}' in table.")

        row_count = len(self.tree.get_children())
        if row_count == 0: self.update_status("Loaded data frame, but no rows were added to the table."); self.clear_pdf_preview("Data loaded, but no rows to display.")
        if self.current_excel_file and not self.dataframe.empty: self.save_changes_button.config(state=tk.NORMAL)
        else: self.save_changes_button.config(state=tk.DISABLED)

    def sort_treeview_column(self, col, reverse):
        if self.dataframe is None or self.dataframe.empty: return
        if self._edit_entry: self._cancel_cell_edit()
        try:
            sorted_df = self.dataframe.copy()
            try: numeric_col = pd.to_numeric(sorted_df[col], errors='coerce'); is_numeric = not numeric_col.isna().all()
            except Exception: is_numeric = False
            if is_numeric:
                sort_indices = numeric_col.argsort(na_position='last');
                if reverse: sort_indices = sort_indices[::-1]
                sorted_df = sorted_df.iloc[sort_indices]
            else: sorted_df = sorted_df.sort_values(by=col, ascending=not reverse, key=lambda x: x.astype(str).str.lower(), na_position='last')
            self.dataframe = sorted_df
            for c in self.tree["columns"]:
                current_text = self.tree.heading(c, "text").replace(' ▲', '').replace(' ▼', '')
                if c == col: indicator = ' ▲' if not reverse else ' ▼'; self.tree.heading(c, text=current_text + indicator)
                else: self.tree.heading(c, text=current_text)
            self.load_data_to_treeview()
            self._tree_sort_order[col] = not reverse
        except KeyError: print(f"Error: Column '{col}' not found in DataFrame for sorting."); messagebox.showerror("Sort Error", f"Column '{col}' not found.", parent=self.master)
        except Exception as e: print(f"Error sorting treeview column '{col}': {e}"); import traceback; traceback.print_exc(); messagebox.showerror("Sort Error", f"Could not sort column '{col}':\n{e}", parent=self.master)

    def on_tree_double_click(self, event):
        if self._edit_entry: return
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell": return
        item_id = self.tree.identify_row(event.y); column_id_str = self.tree.identify_column(event.x)
        if not item_id or not column_id_str: return
        try: column_index = int(column_id_str.replace('#', '')) - 1; column_name = self.tree["columns"][column_index]
        except (ValueError, IndexError): print(f"Error identifying column from ID: {column_id_str}"); return
        if column_name not in self.editable_columns: return
        try: bbox = self.tree.bbox(item_id, column=column_id_str);
        except Exception: return
        if not bbox: return
        current_values = self.tree.item(item_id, 'values'); original_value = current_values[column_index]
        self._edit_item_id = item_id; self._edit_column_id = column_id_str
        self._edit_entry = ttk.Entry(self.tree, font=('Segoe UI', 9)); self._edit_entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        self._edit_entry.insert(0, original_value); self._edit_entry.select_range(0, tk.END); self._edit_entry.focus_set()
        self._edit_entry.bind("<Return>", lambda e: self._save_cell_edit(item_id, column_index, column_name))
        self._edit_entry.bind("<KP_Enter>", lambda e: self._save_cell_edit(item_id, column_index, column_name))
        self._edit_entry.bind("<Escape>", lambda e: self._cancel_cell_edit())
        self._edit_entry.bind("<FocusOut>", lambda e: self._save_cell_edit(item_id, column_index, column_name))

    def _save_cell_edit(self, item_id, column_index, column_name):
        if not self._edit_entry: return
        new_value = self._edit_entry.get(); self._cancel_cell_edit()
        current_values = list(self.tree.item(item_id, 'values'))
        if column_index < len(current_values):
            if str(current_values[column_index]) != new_value:
                current_values[column_index] = new_value; self.tree.item(item_id, values=tuple(current_values))
                if self.dataframe is not None:
                    try:
                        df_index = self._df_index_map.get(item_id)
                        if df_index is not None and df_index in self.dataframe.index:
                            self.dataframe.loc[df_index, column_name] = new_value
                            self.update_status(f"Cell updated: Row {df_index}, {column_name} = '{new_value}'")
                            if self.current_excel_file: self.save_changes_button.config(state=tk.NORMAL)
                        else: print(f"Error: Could not find DataFrame index ({df_index}) for tree item {item_id}"); self.update_status(f"Error: Failed to update backing data for edit.")
                    except Exception as e: print(f"Error updating DataFrame: {e}"); self.update_status(f"Error: Failed to update backing data for edit ({type(e).__name__}).")

    def _cancel_cell_edit(self, event=None):
        if self._edit_entry: self._edit_entry.destroy(); self._edit_entry = None; self._edit_item_id = None; self._edit_column_id = None

    def show_context_menu(self, event):
        item_id = self.tree.identify_row(event.y)
        if item_id:
            if item_id not in self.tree.selection(): self.tree.selection_set(item_id)
        can_delete_copy = bool(self.tree.selection()); can_paste = bool(self._clipboard); can_add = self.dataframe is not None
        self.context_menu.entryconfig("Add Blank Row", state=tk.NORMAL if can_add else tk.DISABLED)
        self.context_menu.entryconfig("Copy Selected Row(s)", state=tk.NORMAL if can_delete_copy else tk.DISABLED)
        self.context_menu.entryconfig("Paste Row(s)", state=tk.NORMAL if can_paste else tk.DISABLED)
        self.context_menu.entryconfig("Delete Selected Row(s)", state=tk.NORMAL if can_delete_copy else tk.DISABLED)
        try: self.context_menu.tk_popup(event.x_root, event.y_root)
        finally: self.context_menu.grab_release()

    # --- Row Operations ---

    def _add_row(self):
        """Adds a new blank row to the DataFrame and reloads the Treeview."""
        if self.dataframe is None:
            self.update_status("Cannot add row: No data loaded.")
            return
        if self._edit_entry: self._cancel_cell_edit()

        try:
            # Create a dictionary for the new row
            new_row_data = {col: "" for col in self.dataframe.columns}
            new_row_data["Source PDF"] = "[MANUALLY ADDED]" # Indicate source
            new_row_data["Full PDF Path"] = "" # No path

            # Convert to DataFrame
            new_row_df = pd.DataFrame([new_row_data])

            # Append to the main DataFrame and reset the index
            # ignore_index=True handles creating a new, unique sequential index
            self.dataframe = pd.concat([self.dataframe, new_row_df], ignore_index=True)

            # --- Reload Treeview to reflect the change and new indices ---
            self.load_data_to_treeview()

            # Get the ID of the last item added after reloading
            children = self.tree.get_children()
            if children:
                new_item_id = children[-1]
                self.tree.see(new_item_id) # Scroll to the new row
                self.tree.selection_set(new_item_id) # Optionally select the new row

            self.update_status("Added 1 blank row.")
            # Ensure save button is enabled if possible
            if self.current_excel_file and not self.dataframe.empty:
                 self.save_changes_button.config(state=tk.NORMAL)

        except Exception as e:
            # Catch potential errors during concat or reload
            print(f"Error adding row: {e}")
            self.update_status(f"Error: Could not add row ({type(e).__name__}).")
            import traceback; traceback.print_exc()
            # Attempt to reload treeview even on error to try and resync
            try:
                self.load_data_to_treeview()
            except Exception as reload_e:
                print(f"Error reloading treeview after add error: {reload_e}")


    def _delete_selected_rows(self):
        """Deletes selected rows from DataFrame and Treeview."""
        selected_item_ids = self.tree.selection()
        if not selected_item_ids: self.update_status("No rows selected to delete."); return
        if not messagebox.askyesno("Confirm Delete", f"Are you sure you want to permanently delete {len(selected_item_ids)} selected row(s)?", parent=self.master):
            self.update_status("Deletion cancelled."); return
        if self._edit_entry: self._cancel_cell_edit()

        try:
            indices_to_drop = []
            for item_id in selected_item_ids:
                df_index = self._df_index_map.get(item_id)
                if df_index is not None and df_index in self.dataframe.index: # Check if index still exists
                    indices_to_drop.append(df_index)
                else: print(f"Warning: Could not find DataFrame index for tree item {item_id} during delete.")

            if not indices_to_drop: self.update_status("Error: Could not map selected items to data for deletion."); return

            # Drop rows from DataFrame
            self.dataframe.drop(index=indices_to_drop, inplace=True)
            # IMPORTANT: Reset index after dropping to keep it sequential
            self.dataframe.reset_index(drop=True, inplace=True)

            # --- Reload Treeview to reflect the change and new indices ---
            self.load_data_to_treeview()

            self.update_status(f"Deleted {len(selected_item_ids)} row(s).")
            # Update save button state
            if self.current_excel_file and self.dataframe is not None and not self.dataframe.empty:
                 self.save_changes_button.config(state=tk.NORMAL)
            else: # Data might be empty now
                 self.save_changes_button.config(state=tk.DISABLED)

        except Exception as e:
            print(f"Error deleting rows: {e}")
            self.update_status(f"Error: Could not delete rows ({type(e).__name__}).")
            import traceback; traceback.print_exc()
            # Attempt reload to resync
            try: self.load_data_to_treeview()
            except Exception as reload_e: print(f"Error reloading treeview after delete error: {reload_e}")


    def _copy_selected_rows(self):
        """Copies data of selected rows to the internal clipboard."""
        selected_item_ids = self.tree.selection()
        if not selected_item_ids: self.update_status("No rows selected to copy."); return
        if self.dataframe is None: self.update_status("Cannot copy: Data not loaded."); return

        self._clipboard = []
        copied_count = 0
        try:
            # Get corresponding DataFrame indices, respecting current sort order in Treeview
            df_indices_to_copy = []
            for item_id in selected_item_ids: # Iterate through selection
                 df_index = self._df_index_map.get(item_id)
                 if df_index is not None and df_index in self.dataframe.index:
                      df_indices_to_copy.append(df_index)

            if not df_indices_to_copy:
                 self.update_status("Error: Could not map selected items to data for copying.")
                 return

            # Retrieve rows from DataFrame based on collected indices
            copied_data = self.dataframe.loc[df_indices_to_copy]

            # Convert the copied DataFrame rows to list of dictionaries for clipboard
            self._clipboard = copied_data.to_dict('records')
            copied_count = len(self._clipboard)

            if copied_count > 0: self.update_status(f"Copied {copied_count} row(s) to clipboard.")
            else: self.update_status("Error: Failed to copy selected row data.")

        except Exception as e:
            print(f"Error copying rows: {e}")
            self.update_status(f"Error: Could not copy rows ({type(e).__name__}).")
            self._clipboard = []
            import traceback; traceback.print_exc()


    def _paste_rows(self):
        """Pastes rows from the clipboard to the DataFrame and Treeview."""
        if not self._clipboard: self.update_status("Clipboard is empty. Nothing to paste."); return
        if self.dataframe is None: self.update_status("Cannot paste: No data loaded."); return
        if self._edit_entry: self._cancel_cell_edit()

        try:
            # Create a DataFrame from the clipboard data (list of dicts)
            # Pandas automatically aligns columns, missing columns in clipboard become NaN
            pasted_df = pd.DataFrame(self._clipboard)

            # Ensure pasted data only contains columns present in the target dataframe
            pasted_df = pasted_df.reindex(columns=self.dataframe.columns, fill_value="")

            if pasted_df.empty:
                 self.update_status("No valid rows to paste from clipboard after column alignment.")
                 return

            # Append to the main DataFrame and reset the index
            self.dataframe = pd.concat([self.dataframe, pasted_df], ignore_index=True)

            # --- Reload Treeview to reflect the change and new indices ---
            self.load_data_to_treeview()

            # Try to scroll to the first pasted row
            children = self.tree.get_children()
            num_pasted = len(pasted_df)
            if len(children) >= num_pasted:
                 first_pasted_item_id = children[-num_pasted]
                 self.tree.see(first_pasted_item_id)
                 # Select the pasted rows (optional)
                 # self.tree.selection_set(children[-num_pasted:])


            self.update_status(f"Pasted {len(pasted_df)} row(s) from clipboard.")
            if self.current_excel_file and not self.dataframe.empty:
                 self.save_changes_button.config(state=tk.NORMAL)

        except Exception as e:
            print(f"Error pasting rows: {e}")
            self.update_status(f"Error: Could not paste rows ({type(e).__name__}).")
            import traceback; traceback.print_exc()
            # Attempt reload to resync
            try: self.load_data_to_treeview()
            except Exception as reload_e: print(f"Error reloading treeview after paste error: {reload_e}")

    # --- Excel Saving Method ---
    # (save_changes_to_excel remains the same)
    def save_changes_to_excel(self):
        if self._edit_entry: self._cancel_cell_edit()
        if self.dataframe is None or self.dataframe.empty: messagebox.showwarning("No Data", "There is no data to save.", parent=self.master); return
        if not self.current_excel_file: messagebox.showerror("Save Error", "The original Excel file path is unknown...", parent=self.master); return
        if not os.path.exists(os.path.dirname(self.current_excel_file)): messagebox.showerror("Save Error", f"The directory for the Excel file no longer exists...", parent=self.master); return
        if not messagebox.askyesno("Confirm Save", f"This will overwrite the existing file:\n{self.current_excel_file}\n\nAre you sure?", parent=self.master): self.update_status("Save cancelled by user."); return
        self.update_status(f"Attempting to save changes to: {self.current_excel_file}")
        try:
            excel_columns = ["Source PDF", "Invoice Folio", "Document Type", "Reference Number"]
            cols_to_save = [col for col in excel_columns if col in self.dataframe.columns]
            if not cols_to_save: raise ValueError("No valid columns found to save.")
            df_to_save = self.dataframe[cols_to_save].copy()
            df_to_save.to_excel(self.current_excel_file, index=False, engine='openpyxl')
            success_msg = f"Changes successfully saved to:\n{self.current_excel_file}"; self.update_status(success_msg); messagebox.showinfo("Save Successful", success_msg, parent=self.master)
        except PermissionError: error_message = f"Error saving changes to:\n{self.current_excel_file}\n\nPermission denied..."; print(error_message); self.update_status(f"Save Error: Permission Denied."); messagebox.showerror("Save Error", error_message, parent=self.master)
        except Exception as e: error_message = f"An unexpected error occurred while saving changes...\n{e}"; print(error_message); import traceback; traceback.print_exc(); self.update_status(f"Save Error: {e}. See console."); messagebox.showerror("Save Error", error_message, parent=self.master)


    # --- Treeview Selection & PDF Preview ---
    # (on_treeview_select, clear_pdf_preview, update_pdf_preview, zoom methods remain the same)
    def on_treeview_select(self, event):
        if self._edit_entry and self.tree.focus() != self._edit_entry:
             if self._edit_item_id and self._edit_column_id:
                  try: col_idx = int(self._edit_column_id.replace('#', '')) - 1; col_name = self.tree["columns"][col_idx]; self._save_cell_edit(self._edit_item_id, col_idx, col_name)
                  except: self._cancel_cell_edit()
             else: self._cancel_cell_edit()
        selected_items = self.tree.selection()
        if not selected_items: return
        selected_item_id = selected_items[0]; pdf_full_path = self._pdf_path_map.get(selected_item_id)
        if pdf_full_path and isinstance(pdf_full_path, str):
            if os.path.exists(pdf_full_path): self.update_pdf_preview(pdf_full_path)
            else: base_name = os.path.basename(pdf_full_path); err_msg = f"File Not Found:\n{base_name}\n(Path: {pdf_full_path})"; self.clear_pdf_preview(err_msg); print(f"Error: File path from selection '{pdf_full_path}' does not exist."); self.update_status(f"Preview Error: File not found - {base_name}"); self.current_preview_pdf_path = None
        elif pdf_full_path is None and selected_item_id in self._df_index_map: self.clear_pdf_preview("No PDF associated with this row."); self.current_preview_pdf_path = None
        elif selected_item_id not in self._df_index_map: print(f"Error: Selected item ID {selected_item_id} not found in data map."); self.clear_pdf_preview("Error: Cannot find data for selected row."); self.current_preview_pdf_path = None
        else: print(f"Error: Invalid path data associated with selected row: {pdf_full_path}"); self.clear_pdf_preview("Error: Invalid file path\nin selected row data."); self.current_preview_pdf_path = None

    def clear_pdf_preview(self, message="Select a row to preview PDF"):
        if self._canvas_image_id:
             try: self.pdf_canvas.delete(self._canvas_image_id)
             except tk.TclError: pass
             self._canvas_image_id = None
        self.pdf_preview_image = None
        if self._placeholder_window_id:
            try: self.pdf_canvas.delete("placeholder")
            except tk.TclError: pass
            self._placeholder_window_id = None
        if hasattr(self, 'pdf_placeholder_label') and self.pdf_placeholder_label.winfo_exists(): self.pdf_placeholder_label.config(text=message)
        else: self.pdf_placeholder_label = ttk.Label(self.pdf_canvas, text=message, style="Placeholder.TLabel")
        canvas_w = self.pdf_canvas.winfo_width() if self.pdf_canvas.winfo_width() > 1 else 200; canvas_h = self.pdf_canvas.winfo_height() if self.pdf_canvas.winfo_height() > 1 else 200
        self._placeholder_window_id = self.pdf_canvas.create_window(canvas_w//2, canvas_h//2, window=self.pdf_placeholder_label, anchor=tk.CENTER, tags="placeholder")
        self.master.after_idle(self._update_canvas_scrollregion); self.current_preview_pdf_path = None

    def update_pdf_preview(self, pdf_path):
        if not PIL_AVAILABLE: self.clear_pdf_preview("PDF Preview disabled:\nPillow library not found."); return
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
            doc = fitz.open(pdf_path);
            if len(doc) == 0: raise ValueError("PDF has no pages.")
            page = doc.load_page(0); page_rect = page.rect
            if page_rect.width == 0 or page_rect.height == 0: raise ValueError("PDF page has zero dimensions.")
            zoom_factor = self.current_zoom_factor; mat = fitz.Matrix(zoom_factor, zoom_factor)
            pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB, alpha=False); doc.close()
            img_bytes = pix.samples;
            if not isinstance(img_bytes, bytes): img_bytes = bytes(img_bytes)
            if not img_bytes: raise ValueError("Pixmap samples are empty.")
            pil_image = Image.frombytes("RGB", [pix.width, pix.height], img_bytes)
            self.pdf_preview_image = ImageTk.PhotoImage(image=pil_image)
            if self.pdf_preview_image:
                self._canvas_image_id = self.pdf_canvas.create_image(0, 0, anchor=tk.NW, image=self.pdf_preview_image, tags="pdf_image")
                self.current_preview_pdf_path = pdf_path; self.master.after_idle(self._update_canvas_scrollregion)
                self.pdf_canvas.xview_moveto(0); self.pdf_canvas.yview_moveto(0)
            else: raise ValueError("Failed to create PhotoImage object.")
        except (fitz.fitz.FileNotFoundError, fitz.fitz.PasswordError, ValueError, Exception) as e:
            base_name = os.path.basename(pdf_path) if pdf_path else "Unknown File"; error_type = type(e).__name__
            error_msg = f"Preview Error ({error_type}):\n{base_name}"
            if isinstance(e, fitz.fitz.PasswordError): error_msg += "\n(Password Protected?)"
            elif isinstance(e, ValueError): error_msg += f"\n({e})"
            import traceback; print(f"--- PDF Preview Exception Traceback for {base_name} ---"); traceback.print_exc(); print(f"--- End Traceback ---")
            self.clear_pdf_preview(error_msg)

    def zoom_in(self):
        if self.current_preview_pdf_path:
            new_zoom = min(self.current_zoom_factor * self.zoom_step, self.max_zoom)
            if new_zoom != self.current_zoom_factor: self.current_zoom_factor = new_zoom; self.update_pdf_preview(self.current_preview_pdf_path)

    def zoom_out(self):
        if self.current_preview_pdf_path:
            new_zoom = max(self.current_zoom_factor / self.zoom_step, self.min_zoom)
            if new_zoom != self.current_zoom_factor: self.current_zoom_factor = new_zoom; self.update_pdf_preview(self.current_preview_pdf_path)

    def reset_zoom(self):
        if self.current_preview_pdf_path and self.current_zoom_factor != 1.0:
            self.current_zoom_factor = 1.0; self.update_pdf_preview(self.current_preview_pdf_path)

    def _update_canvas_scrollregion(self):
        try:
            if self._canvas_image_id and self.pdf_canvas.find_withtag(self._canvas_image_id):
                bbox = self.pdf_canvas.bbox(self._canvas_image_id)
                if bbox: scroll_bbox = (bbox[0], bbox[1], bbox[2] + 5, bbox[3] + 5); self.pdf_canvas.config(scrollregion=scroll_bbox); return
            current_width = self.pdf_canvas.winfo_width(); current_height = self.pdf_canvas.winfo_height()
            self.pdf_canvas.config(scrollregion=(0, 0, max(1, current_width), max(1, current_height)))
        except tk.TclError as e: print(f"Warning: TclError updating canvas scrollregion: {e}")
        except Exception as e: print(f"Error updating canvas scrollregion: {e}")


# --- Main Execution ---
if __name__ == "__main__":
    if not PIL_AVAILABLE:
         root_check = tk.Tk(); root_check.withdraw()
         messagebox.showwarning("Dependency Warning", "Python Imaging Library (Pillow) not found...\npip install Pillow", parent=None)
         root_check.destroy()

    root = tk.Tk()
    app = InvoiceProcessorApp(root)
    root.mainloop()