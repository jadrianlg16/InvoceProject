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
    # Optional: Show a warning if Pillow isn't installed, as PPM fallback is less reliable
    # messagebox.showwarning("Dependency Warning", "Pillow library not found (pip install Pillow).\nPDF preview might be less reliable.")
    print("WARNING: Pillow library not found (pip install Pillow). PDF preview might be less reliable.")


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
            full_text += page.get_text("text", flags=fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE) + "\n"
        doc.close()
        full_text = re.sub(r'\s{2,}', ' ', full_text)
        full_text = re.sub(r'\n+', '\n', full_text)
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
        contpaq_simple_pattern = r'\bFOLIO:\s*(\w+)\b'
        for match in re.finditer(contpaq_simple_pattern, text, re.IGNORECASE):
            start_index = match.start()
            lookback_chars = 25
            preceding_text = text[max(0, start_index - lookback_chars) : start_index]
            if not re.search(r'\bFolio\s+fiscal\b', preceding_text, re.IGNORECASE):
                folio = match.group(1)
                break

    if folio and len(folio) > 20:
         print(f"Warning: Found potentially long folio '{folio}' in {invoice_type}. Might be Folio Fiscal. Skipping.")
         return None
    return folio

def find_references(text):
    """Finds all Escritura and Acta references, handling single, range, list, and special formats."""
    references = []
    if not text: return []
    found_escritura_numbers = set()
    found_acta_numbers = set()
    flags = re.IGNORECASE

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
        except ValueError: pass # Ignore parse errors
    for match in re.finditer(REGEX_ACTA_RANGE, text, flags):
        try:
            start, end = int(match.group(1)), int(match.group(2))
            if start <= end:
                for num in range(start, end + 1):
                    num_str = str(num)
                    if num_str not in found_acta_numbers:
                        references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
                        found_acta_numbers.add(num_str)
        except ValueError: pass # Ignore parse errors

    # --- Lists ---
    for match in re.finditer(REGEX_ESCRITURA_LIST_Y, text, flags):
        for num_str in [match.group(1).strip(), match.group(2).strip()]:
             if num_str not in found_escritura_numbers:
                references.append({"Type": "Escritura", "Number": num_str})
                found_escritura_numbers.add(num_str)
    for match in re.finditer(REGEX_ACTA_LIST_Y, text, flags):
         for num_str in [match.group(1).strip(), match.group(2).strip()]:
             if num_str not in found_acta_numbers:
                references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
                found_acta_numbers.add(num_str)

    # --- Special Acta ---
    for match in re.finditer(REGEX_ACTA_SPECIAL, text, flags):
        num_str = match.group(1).strip()
        if num_str not in found_acta_numbers:
            references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
            found_acta_numbers.add(num_str)

    # --- Singles (with context checks) ---
    for match in re.finditer(REGEX_ESCRITURA_SINGLE, text, flags):
        num_str = match.group(1).strip()
        if num_str not in found_escritura_numbers:
            start_pos, end_pos = match.start(0), match.end(0)
            ctx_before = text[max(0, start_pos - 10):start_pos]
            ctx_after = text[end_pos:min(len(text), end_pos + 10)]
            is_multi_ctx = re.search(r'(?:N[uú]meros?|Nos\.?|N°)\s+\d+\s+(?:Y|A)\s*$', ctx_before, flags) or \
                           re.search(r'^\s*(?:Y|A)\s+\d+', ctx_after, flags)
            if not is_multi_ctx:
                references.append({"Type": "Escritura", "Number": num_str})
                found_escritura_numbers.add(num_str)

    for match in re.finditer(REGEX_ACTA_SINGLE, text, flags):
        num_str = match.group(1).strip()
        if num_str not in found_acta_numbers:
            start_pos, end_pos = match.start(0), match.end(0)
            ctx_before = text[max(0, start_pos - 10):start_pos]
            ctx_after = text[end_pos:min(len(text), end_pos + 10)]
            is_multi_ctx = re.search(r'(?:N[uú]meros?|Nos\.?|N°)\s+\d+\s+(?:Y|A)\s*$', ctx_before, flags) or \
                           re.search(r'^\s*(?:Y|A)\s+\d+', ctx_after, flags)
            is_special_ctx = re.search(r'\d+\/\s*$', ctx_before) or re.search(r'^\/\d+', ctx_after)
            if not is_multi_ctx and not is_special_ctx:
                 references.append({"Type": "Acta Fuera de Protocolo", "Number": num_str})
                 found_acta_numbers.add(num_str)

    # --- Sorting ---
    def sort_key(item):
        try: return (item["Type"], int(item["Number"]))
        except ValueError: return (item["Type"], float('inf'))
    references.sort(key=sort_key)
    return references

def process_single_pdf(pdf_path, invoice_type):
    """Processes a single PDF to extract folio and references. Returns full path."""
    print(f"Processing: {os.path.basename(pdf_path)}")
    text = extract_text_from_pdf(pdf_path)
    if not text:
        print(f"Warning: Could not extract text from {os.path.basename(pdf_path)}. Skipping.")
        # Return structure indicating failure but including path for potential reporting
        return [{"Document Type": "ERROR", "Reference Number": "Text Extraction Failed",
                 "Invoice Folio": "ERROR", "Source PDF": os.path.basename(pdf_path),
                 "Full PDF Path": pdf_path}] # Include full path even on error

    folio = find_folio(text, invoice_type)
    if not folio:
        folio = "NOT_FOUND" # Use explicit marker

    references = find_references(text)
    output_rows = []

    if not references:
         output_rows.append({
            "Document Type": "N/A",
            "Reference Number": "N/A",
            "Invoice Folio": folio,
            "Source PDF": os.path.basename(pdf_path), # Keep basename for display/Excel
            "Full PDF Path": pdf_path # Add full path
        })
    else:
        for ref in references:
            output_rows.append({
                "Document Type": ref["Type"],
                "Reference Number": ref["Number"],
                "Invoice Folio": folio,
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
        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith('.pdf'):
                    pdf_files.append(os.path.join(root, file)) # Collect full paths
    except Exception as e:
        # Handle errors during directory walk
        error_msg = f"Error accessing folder structure in {folder_path}: {e}"
        print(error_msg)
        app_instance.master.after(0, app_instance.update_status, f"Error accessing folder. Check permissions or path.")
        app_instance.master.after(0, messagebox.showerror, "Folder Error", error_msg)
        app_instance.master.after(0, app_instance.enable_buttons)
        return # Stop processing if folder walk fails

    total_files = len(pdf_files)
    if total_files == 0:
        app_instance.master.after(0, app_instance.update_status, "No PDF files found in the selected folder or subfolders.")
        app_instance.master.after(0, app_instance.enable_buttons)
        app_instance.master.after(0, messagebox.showinfo, "Information", "No PDF files found.")
        return

    start_time = time.time()
    files_processed_count = 0
    files_with_errors = 0

    # Process each found PDF
    for i, pdf_path in enumerate(pdf_files):
        status_message = f"Processing file {i+1}/{total_files}: {os.path.basename(pdf_path)}"
        app_instance.master.after(0, app_instance.update_status, status_message)
        try:
            results = process_single_pdf(pdf_path, invoice_type)
            if results:
                all_data.extend(results)
                # Check if the first result indicates an error from process_single_pdf itself
                if results[0].get("Document Type") == "ERROR":
                    files_with_errors += 1
                else:
                    # Only count as processed if no error was reported by process_single_pdf
                    files_processed_count += 1 # Increment here only if process_single_pdf succeeded
            else: # Should ideally not happen if process_single_pdf returns []
                 files_with_errors += 1 # Count as error if nothing is returned
                 app_instance.master.after(0, app_instance.update_status, f"Warning: No data returned for {os.path.basename(pdf_path)}")

        except Exception as e:
            files_with_errors += 1
            error_msg = f"CRITICAL ERROR processing file {os.path.basename(pdf_path)}: {e}"
            print(error_msg) # Log detailed error to console
            import traceback
            traceback.print_exc()
            # Update status briefly about the file error
            app_instance.master.after(0, app_instance.update_status, f"ERROR processing {os.path.basename(pdf_path)}. See console.")
            # Add error entry to data to potentially show in excel/treeview
            all_data.append({
                 "Document Type": "ERROR", "Reference Number": f"Processing Error: {e}",
                 "Invoice Folio": "ERROR", "Source PDF": os.path.basename(pdf_path),
                 "Full PDF Path": pdf_path})

    end_time = time.time()
    processing_time = end_time - start_time
    # Adjust summary message for clarity on "processed" vs "encountered"
    final_summary = f"{files_processed_count}/{total_files} files processed successfully"
    if files_with_errors > 0:
        error_files_count = files_with_errors
        total_encountered = files_processed_count + error_files_count
        # Ensure total doesn't exceed original file count if errors happened during file finding
        if total_encountered != total_files:
             print(f"Warning: Processed count ({files_processed_count}) + error count ({error_files_count}) != total files ({total_files}). Check file listing.")
        final_summary += f" ({error_files_count} file(s) encountered errors)"


    if not all_data:
        final_message = f"Processing complete ({final_summary} in {processing_time:.2f}s).\nNo data extracted."
        app_instance.master.after(0, app_instance.update_status, final_message)
        app_instance.master.after(0, app_instance.enable_buttons)
        app_instance.master.after(0, messagebox.showinfo, "Information", final_message)
        return

    # Create DataFrame - including the Full PDF Path
    try:
        df = pd.DataFrame(all_data)
        # Define desired columns for Excel output (exclude Full PDF Path if not wanted in the file)
        excel_columns = ["Source PDF", "Invoice Folio", "Document Type", "Reference Number"]
        # Define order for DataFrame (keep Full PDF Path for Treeview use)
        all_columns_ordered = ["Source PDF", "Invoice Folio", "Document Type", "Reference Number", "Full PDF Path"]
        # Ensure all expected columns exist, fill missing ones if necessary
        for col in all_columns_ordered:
            if col not in df.columns:
                df[col] = None # Or appropriate default like "N/A"
        df = df[all_columns_ordered] # Reorder/select

        # Optional: Sort DataFrame before saving/displaying
        try:
            # Convert Reference Number to numeric temporarily for sorting, handle errors
            df['Reference Number Int'] = pd.to_numeric(df['Reference Number'], errors='coerce').fillna(float('inf'))
            # Sort by multiple criteria
            df.sort_values(by=["Source PDF", "Invoice Folio", "Document Type", "Reference Number Int"], inplace=True, na_position='last')
            df.drop(columns=['Reference Number Int'], inplace=True) # Remove temporary sort column
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
        return

    # Generate unique output filename
    output_filename = find_unique_output_filename("Extracted_Invoices.xlsx")

    # Save to Excel (only the selected columns)
    try:
        # Ensure only the excel_columns exist in the DataFrame slice being saved
        df_to_save = df[excel_columns].copy()
        df_to_save.to_excel(output_filename, index=False, engine='openpyxl') # Save only desired cols
        final_message = f"Processing complete ({final_summary} in {processing_time:.2f}s).\nData saved to:\n{output_filename}"

        # Schedule GUI updates (status, load to treeview, enable buttons) on the main thread
        app_instance.master.after(0, app_instance.update_status, final_message)
        app_instance.master.after(0, app_instance.load_data_to_treeview, df) # Pass the full DataFrame
        app_instance.master.after(0, messagebox.showinfo, "Success", final_message)

    except PermissionError:
        error_message = f"Error saving Excel file '{output_filename}':\nPermission denied. The file might be open.\nPlease close it and try again."
        print(error_message)
        app_instance.master.after(0, app_instance.update_status, f"Error saving file: Permission Denied.")
        app_instance.master.after(0, messagebox.showerror, "Save Error", error_message)
        # Load data to treeview even if saving failed, so user can see results
        app_instance.master.after(10, app_instance.load_data_to_treeview, df)
    except Exception as e:
        error_message = f"Error saving Excel file '{output_filename}':\n{e}"
        print(error_message)
        import traceback
        traceback.print_exc()
        app_instance.master.after(0, app_instance.update_status, f"Error saving file. See console.")
        app_instance.master.after(0, messagebox.showerror, "Save Error", error_message)
        # Load data to treeview even if saving failed
        app_instance.master.after(10, app_instance.load_data_to_treeview, df)

    # Always re-enable buttons, scheduled after other updates
    app_instance.master.after(20, app_instance.enable_buttons) # Small delay


# --- GUI Class ---
class InvoiceProcessorApp:
    def __init__(self, master):
        self.master = master
        master.title("Invoice PDF Processor v1.3 - Preview Enabled")
        master.geometry("1200x750") # Slightly taller for comfort

        self.selected_folder = tk.StringVar(value="No folder selected")
        self.processing_active = False
        self.dataframe = None # To store the loaded DataFrame
        self.pdf_preview_image = None # To hold the PhotoImage reference
        self._canvas_image_id = None # Store canvas image item ID

        # Configure styles
        style = ttk.Style(root)
        style.theme_use('clam') # Or 'alt', 'default', 'classic', 'vista', 'xpnative'
        style.configure('TButton', padding=6, relief="flat", background="#ccc")
        style.map('TButton', background=[('active', '#eee')])
        style.configure('TLabel', padding=5)
        style.configure('TEntry', padding=5)
        style.configure("Treeview.Heading", font=('TkDefaultFont', 10, 'bold'))
        style.configure("Placeholder.TLabel", foreground="grey", padding=10) # Style for placeholder

        # --- Main Layout: Paned Window (Horizontal) ---
        main_paned_window = ttk.PanedWindow(master, orient=tk.HORIZONTAL)
        main_paned_window.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # --- Left Pane: Controls ---
        left_frame = ttk.Frame(main_paned_window, padding="5 5 5 5", width=450) # Set initial width
        left_frame.pack_propagate(False) # Prevent frame from shrinking
        main_paned_window.add(left_frame, weight=1) # Allow resizing

        # Folder Selection Frame (inside left_frame)
        folder_frame = ttk.Frame(left_frame, padding="10 10 10 5")
        folder_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(folder_frame, text="Invoice Folder:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(folder_frame, textvariable=self.selected_folder, state="readonly", width=40).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        self.select_button = ttk.Button(folder_frame, text="Select...", command=self.select_folder)
        self.select_button.pack(side=tk.LEFT)

        # Invoice Type and Processing Frame (inside left_frame)
        process_frame = ttk.Frame(left_frame, padding="10 5 10 5")
        process_frame.pack(fill=tk.X, pady=5)
        ttk.Label(process_frame, text="Select Invoice Type and Process:").pack(pady=(0, 10), anchor=tk.W)
        button_frame = ttk.Frame(process_frame)
        button_frame.pack(pady=5)
        self.dba_button = ttk.Button(button_frame, text="Process DBA", command=lambda: self.start_processing('DBA'), width=15)
        self.dba_button.pack(side=tk.LEFT, padx=10, pady=5, ipady=4)
        self.totalnot_button = ttk.Button(button_frame, text="Process TOTALNOT", command=lambda: self.start_processing('TOTALNOT'), width=15)
        self.totalnot_button.pack(side=tk.LEFT, padx=10, pady=5, ipady=4)
        self.contpaq_button = ttk.Button(button_frame, text="Process CONTPAQ", command=lambda: self.start_processing('CONTPAQ'), width=15)
        self.contpaq_button.pack(side=tk.LEFT, padx=10, pady=5, ipady=4)

        # Status Area Frame (inside left_frame)
        status_frame = ttk.Frame(left_frame, padding="10 5 10 10")
        status_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        ttk.Label(status_frame, text="Status Log:").pack(anchor=tk.W)
        text_scroll_frame = ttk.Frame(status_frame)
        text_scroll_frame.pack(fill=tk.BOTH, expand=True, pady=(5,0))
        scrollbar_status = ttk.Scrollbar(text_scroll_frame)
        scrollbar_status.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text = tk.Text(text_scroll_frame, height=10, wrap=tk.WORD, state=tk.DISABLED,
                                   relief=tk.SUNKEN, borderwidth=1, yscrollcommand=scrollbar_status.set,
                                   font=("Consolas", 9))
        self.status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_status.config(command=self.status_text.yview)

        # --- Right Pane: Previews ---
        right_paned_window = ttk.PanedWindow(main_paned_window, orient=tk.VERTICAL)
        main_paned_window.add(right_paned_window, weight=3) # More space for previews

        # Top-Right Pane: Excel Treeview
        tree_frame = ttk.Frame(right_paned_window, padding=5)
        right_paned_window.add(tree_frame, weight=2) # Treeview gets less height than preview

        ttk.Label(tree_frame, text="Excel Data Preview:", font=('TkDefaultFont', 10, 'bold')).pack(anchor=tk.NW, pady=(0, 5))
        # Treeview scrollbars
        tree_scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        # Treeview widget
        self.tree = ttk.Treeview(tree_frame,
                                 columns=("Source PDF", "Invoice Folio", "Document Type", "Reference Number"), # Display Columns
                                 show='headings',
                                 yscrollcommand=tree_scroll_y.set,
                                 xscrollcommand=tree_scroll_x.set,
                                 selectmode='browse') # Single selection
        # Define headings
        self.tree.heading("Source PDF", text="Source PDF", command=lambda: self.sort_treeview_column("Source PDF", False))
        self.tree.heading("Invoice Folio", text="Invoice Folio", command=lambda: self.sort_treeview_column("Invoice Folio", False))
        self.tree.heading("Document Type", text="Document Type", command=lambda: self.sort_treeview_column("Document Type", False))
        self.tree.heading("Reference Number", text="Reference Number", command=lambda: self.sort_treeview_column("Reference Number", False))
        # Define column widths
        self.tree.column("Source PDF", anchor=tk.W, width=200, stretch=tk.NO)
        self.tree.column("Invoice Folio", anchor=tk.W, width=100, stretch=tk.NO)
        self.tree.column("Document Type", anchor=tk.W, width=150, stretch=tk.NO)
        self.tree.column("Reference Number", anchor=tk.W, width=120, stretch=tk.NO)
        self.tree.pack(fill=tk.BOTH, expand=True)
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)
        # Bind selection event
        self.tree.bind('<<TreeviewSelect>>', self.on_treeview_select)
        # Dictionary to track sort order for columns
        self._tree_sort_order = {col: False for col in ("Source PDF", "Invoice Folio", "Document Type", "Reference Number")}


        # Bottom-Right Pane: PDF Preview
        pdf_preview_frame = ttk.Frame(right_paned_window, padding=5)
        right_paned_window.add(pdf_preview_frame, weight=3) # PDF preview gets more height

        ttk.Label(pdf_preview_frame, text="PDF Preview (First Page):", font=('TkDefaultFont', 10, 'bold')).pack(anchor=tk.NW, pady=(0, 5))
        # Canvas for PDF Preview
        self.pdf_canvas = tk.Canvas(pdf_preview_frame, bg="lightgrey", relief=tk.SUNKEN, borderwidth=1)
        self.pdf_canvas.pack(fill=tk.BOTH, expand=True)
        # Placeholder label initially inside the canvas
        self.pdf_placeholder_label = ttk.Label(self.pdf_canvas, text="Select a row above to preview PDF", style="Placeholder.TLabel")
        # Place label initially (will be hidden/managed by clear/update)
        self.pdf_canvas.create_window(10, 10, window=self.pdf_placeholder_label, anchor=tk.NW, tags="placeholder")


        # --- Initial Status ---
        self.update_status("Ready. Please select a folder and invoice type.")


    def select_folder(self):
        if self.processing_active: return
        folder = filedialog.askdirectory()
        if folder:
            normalized_folder = os.path.normpath(folder)
            self.selected_folder.set(normalized_folder)
            self.clear_status()
            self.clear_treeview() # Clear previous results
            self.clear_pdf_preview("Select a row to preview PDF") # Reset preview
            self.update_status(f"Folder selected: {normalized_folder}\nReady to process.")
        else:
             if self.selected_folder.get() != "No folder selected":
                 self.selected_folder.set("No folder selected")
                 self.update_status("Folder selection cancelled.")

    def clear_status(self):
         self.status_text.config(state=tk.NORMAL)
         self.status_text.delete('1.0', tk.END)
         self.status_text.config(state=tk.DISABLED)

    def update_status(self, message):
        # Ensure GUI updates are on the main thread
        if threading.current_thread() != threading.main_thread():
            self.master.after(0, lambda msg=message: self._update_status_text(msg))
        else:
            self._update_status_text(message)

    def _update_status_text(self, message):
        """Internal method to update the status text widget safely."""
        current_state = self.status_text.cget('state')
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state=current_state) # Restore state

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
            messagebox.showerror("Error", "Please select a folder first.")
            return
        if not os.path.isdir(folder):
             messagebox.showerror("Error", f"Invalid directory:\n{folder}")
             return
        if self.processing_active:
            messagebox.showwarning("Busy", "Processing is already in progress.")
            return

        self.disable_buttons()
        self.clear_status()
        self.clear_treeview() # Clear previous results
        self.clear_pdf_preview("Processing... Select row after completion.") # Update preview placeholder
        self.update_status(f"Starting recursive processing for {invoice_type} in:\n{folder}")
        self.update_status("-" * 30)

        process_thread = threading.Thread(target=run_processing,
                                          args=(folder, invoice_type, self),
                                          daemon=True)
        process_thread.start()

    def clear_treeview(self):
        """Clears all items from the Treeview."""
        if hasattr(self, 'tree'):
            for item in self.tree.get_children():
                self.tree.delete(item)
        self.dataframe = None # Clear associated dataframe

    def load_data_to_treeview(self, df):
        """Loads data from the DataFrame into the Treeview."""
        self.clear_treeview() # Clear existing data first
        if df is None or df.empty:
             self.update_status("No data to display in the table.")
             return
        self.dataframe = df # Store the dataframe

        # Define the columns to display in the Treeview
        display_columns = list(self.tree["columns"]) # Get columns from tree definition

        # Check if essential columns exist in the DataFrame
        required_cols = display_columns + ["Full PDF Path"]
        if not all(col in df.columns for col in required_cols):
            missing = [col for col in required_cols if col not in df.columns]
            self.update_status(f"Error: DataFrame is missing required columns for display: {missing}")
            messagebox.showerror("Error", f"Could not display data: Missing columns {missing}.")
            print(f"DataFrame columns: {df.columns.tolist()}") # Debugging
            return

        # Insert data row by row
        self.tree.configure(displaycolumns=display_columns) # Ensure correct columns are shown
        for index, row in df.iterrows():
            try:
                # Prepare values tuple FOR DISPLAY - must match self.tree["columns"] order
                # Handle potential NaN or None values gracefully for display
                display_values = tuple(str(row[col]) if pd.notna(row[col]) else "" for col in display_columns)

                # Store the full path along with the display values when inserting
                item_values_with_path = list(display_values)
                item_values_with_path.append(row["Full PDF Path"]) # Add full path at the end

                # Insert into treeview. The `values` param takes the list/tuple.
                self.tree.insert("", tk.END, values=item_values_with_path, tags=(index,)) # Use index as tag

            except Exception as e:
                print(f"Error adding row {index} to treeview: {e}")
                self.update_status(f"Warning: Could not display row {index} in table.")

        self.update_status(f"Loaded {len(df)} rows into the table.")


    def sort_treeview_column(self, col, reverse):
        """Sorts the treeview column based on the clicked header."""
        if self.dataframe is None or self.dataframe.empty:
            return # No data to sort

        try:
            # Use the stored DataFrame for sorting
            # Decide if sorting should be numeric
            is_numeric = pd.api.types.is_numeric_dtype(self.dataframe[col])

            # Make a copy to avoid modifying the original DataFrame directly unless intended
            sorted_df = self.dataframe.copy()

            if is_numeric:
                 # Convert to numeric, coercing errors for safe sorting
                 sorted_df[col] = pd.to_numeric(sorted_df[col], errors='coerce')
                 sorted_df.sort_values(by=col, ascending=not reverse, inplace=True, na_position='last')
            else:
                 # String sort, handle potential None/NaN as empty strings for sorting
                 sorted_df.sort_values(by=col, ascending=not reverse, inplace=True, key=lambda x: x.astype(str).str.lower(), na_position='last')

            # Reload the sorted data into the Treeview
            self.load_data_to_treeview(sorted_df)

            # Update the sort direction for the next click
            self._tree_sort_order[col] = not reverse
            # Optional: Add indicator to header (e.g., ▲▼) - more complex
            # self.tree.heading(col, text=f"{col} {'▲' if not reverse else '▼'}")

        except KeyError:
             print(f"Error: Column '{col}' not found in DataFrame for sorting.")
        except Exception as e:
             print(f"Error sorting treeview column '{col}': {e}")
             messagebox.showerror("Sort Error", f"Could not sort column '{col}':\n{e}")


    def on_treeview_select(self, event):
        """Handles selection changes in the Treeview to update the PDF preview."""
        selected_items = self.tree.selection()
        if not selected_items: return

        selected_item_id = selected_items[0]
        item_values = self.tree.item(selected_item_id, 'values')

        # --- DEBUGGING ---
        print(f"\n--- Treeview Selection ---")
        print(f"Selected Item ID: {selected_item_id}")
        print(f"Retrieved Item Values: {item_values} (Type: {type(item_values)})")
        # --- END DEBUGGING ---

        # Expecting tuple/list of strings from treeview + the appended path
        if not item_values or len(item_values) != len(self.tree["columns"]) + 1:
            print(f"Warning: Retrieved data has unexpected length ({len(item_values)}) for item {selected_item_id}. Expected {len(self.tree['columns'])+1}")
            self.clear_pdf_preview("Error retrieving file path from selection")
            return

        raw_path_data = item_values[-1] # Path is the last element
        pdf_full_path = None

        print(f"Raw path data extracted: '{raw_path_data}' (Type: {type(raw_path_data)})") # DEBUG

        # Ensure it's a non-empty string before proceeding
        if isinstance(raw_path_data, str) and raw_path_data.strip():
            pdf_full_path = raw_path_data.strip() # Basic whitespace strip
        else:
            print(f"Error: Invalid or empty path data retrieved: {raw_path_data}")
            self.clear_pdf_preview("Invalid or missing file path in selected row")
            return

        # --- DEBUGGING ---
        print(f"Cleaned PDF Path Candidate: {pdf_full_path}")
        # --- END DEBUGGING ---

        if pdf_full_path and len(pdf_full_path) > 3: # Basic sanity check
            if os.path.exists(pdf_full_path):
                print(f"Path exists. Calling update_pdf_preview for: {pdf_full_path}")
                self.update_pdf_preview(pdf_full_path)
            else:
                 err_msg = f"File not found:\n{os.path.basename(pdf_full_path)}"
                 self.clear_pdf_preview(err_msg)
                 print(f"Error: File path '{pdf_full_path}' does not exist.")
        else:
             # Should have been caught by the string check above, but as a fallback
             err_msg = "Invalid file path format"
             self.clear_pdf_preview(err_msg)
             print(f"Error: Invalid path after checks: {pdf_full_path}")


    def clear_pdf_preview(self, message="Select a row to preview PDF"):
        """Clears the PDF preview area and shows a placeholder message."""
        # Delete previous image item if it exists
        if self._canvas_image_id:
             try: # Add try-except in case ID is somehow invalid
                 self.pdf_canvas.delete(self._canvas_image_id)
                 #print(f"Deleted canvas image item ID: {self._canvas_image_id}") # DEBUG (verbose)
             except tk.TclError as e:
                 print(f"Warning: Could not delete canvas item {self._canvas_image_id}: {e}")
             self._canvas_image_id = None
        self.pdf_preview_image = None # Clear reference to PhotoImage

        # Ensure the placeholder label is visible and updated
        if hasattr(self, 'pdf_placeholder_label') and self.pdf_placeholder_label.winfo_exists():
             self.pdf_placeholder_label.config(text=message)
             # Move placeholder to front
             self.pdf_canvas.lift("placeholder")
             #print(f"Cleared PDF preview. Displaying message: '{message}'") # DEBUG (verbose)
        else:
            # Recreate label if it was somehow destroyed (shouldn't happen normally)
            self.pdf_placeholder_label = ttk.Label(self.pdf_canvas, text=message, style="Placeholder.TLabel")
            self.pdf_canvas.create_window(10, 10, window=self.pdf_placeholder_label, anchor=tk.NW, tags="placeholder")
            print(f"Recreated placeholder label. Displaying message: '{message}'") # DEBUG

        # Reset scroll region when clearing
        self.master.after_idle(self._update_canvas_scrollregion)


    def update_pdf_preview(self, pdf_path):
        """Opens a PDF, renders the first page, and displays it in the preview canvas."""
        print(f"--- Attempting to update PDF Preview for: {pdf_path} ---") # DEBUG
        if not PIL_AVAILABLE:
             self.clear_pdf_preview("PDF Preview disabled:\nPillow library not found.")
             print("Error: Cannot preview PDF, Pillow library is missing.")
             return

        # Hide placeholder before attempting to load image
        if hasattr(self, 'pdf_placeholder_label') and self.pdf_placeholder_label.winfo_exists():
             self.pdf_canvas.lower("placeholder") # Hide placeholder

        try:
            #print("Opening document with fitz...") # DEBUG (verbose)
            doc = fitz.open(pdf_path)
            if len(doc) == 0:
                self.clear_pdf_preview(f"PDF has no pages:\n{os.path.basename(pdf_path)}")
                if doc: doc.close()
                print("PDF has no pages.") # DEBUG
                return

            #print(f"Loading page 0...") # DEBUG (verbose)
            page = doc.load_page(0) # Load the first page

            # --- Render the page to a pixmap ---
            canvas_width = self.pdf_canvas.winfo_width()
            canvas_height = self.pdf_canvas.winfo_height()
            #print(f"Canvas dimensions: {canvas_width}x{canvas_height}")# DEBUG (verbose)
            if canvas_width <= 1: canvas_width = 500 # Default width
            if canvas_height <= 1: canvas_height = 600 # Default height

            page_rect = page.rect
            target_width = canvas_width - 20 # Padding
            target_height = canvas_height - 20

            zoom_x = target_width / page_rect.width if page_rect.width > 0 else 1
            zoom_y = target_height / page_rect.height if page_rect.height > 0 else 1
            zoom_factor = min(zoom_x, zoom_y, 1.5) # Fit, limit max zoom
            zoom_factor = max(zoom_factor, 0.1) # Prevent tiny zoom

            #print(f"Page rect: {page_rect}, Target: {target_width}x{target_height}, Zoom: {zoom_factor:.2f}") # DEBUG (verbose)
            mat = fitz.Matrix(zoom_factor, zoom_factor)

            #print("Getting pixmap (RGB)...") # DEBUG (verbose)
            pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB, alpha=False) # Force RGB
            #print(f"Pixmap generated: {pix.width}x{pix.height}, Colorspace: {pix.colorspace.name}, Alpha: {pix.alpha}") # DEBUG (verbose)

            #print("Closing fitz document.") # DEBUG (verbose)
            doc.close() # Close promptly

            # --- Convert pixmap to Tkinter PhotoImage using Pillow ---
            self.pdf_preview_image = None # Clear previous reference
            try:
                #print("Converting pixmap samples to PIL Image...") # DEBUG (verbose)
                img_bytes = pix.samples if isinstance(pix.samples, bytes) else bytes(pix.samples)
                if not img_bytes: raise ValueError("Pixmap samples are empty.")

                pil_image = Image.frombytes("RGB", [pix.width, pix.height], img_bytes)
                #print("Converting PIL Image to ImageTk.PhotoImage...") # DEBUG (verbose)
                self.pdf_preview_image = ImageTk.PhotoImage(image=pil_image)
                #print("ImageTk.PhotoImage created successfully.") # DEBUG (verbose)
            except Exception as pil_e:
                print(f"ERROR during Pillow conversion: {pil_e}")
                import traceback
                traceback.print_exc()
                self.clear_pdf_preview(f"Error converting image:\n{pil_e}")
                return # Stop if conversion fails

            # --- Display image on canvas ---
            if self.pdf_preview_image:
                #print("Clearing previous canvas content...") # DEBUG (verbose)
                # Clear previous image item ONLY (leave placeholder hidden)
                if self._canvas_image_id:
                    try: self.pdf_canvas.delete(self._canvas_image_id)
                    except tk.TclError: pass # Ignore if ID invalid
                    self._canvas_image_id = None

                #print("Creating image on canvas...") # DEBUG (verbose)
                self._canvas_image_id = self.pdf_canvas.create_image(10, 10, anchor=tk.NW, image=self.pdf_preview_image, tags="pdf_image")
                #print(f"Canvas image item created with ID: {self._canvas_image_id}") # DEBUG (verbose)

                # Update scroll region after image creation
                self.master.after_idle(self._update_canvas_scrollregion)
                #print("Canvas scrollregion update scheduled via after_idle.") # DEBUG (verbose)
            else:
                 print("ERROR: self.pdf_preview_image is None after conversion attempts.")
                 self.clear_pdf_preview("Failed to create image object")


        except fitz.fitz.FileNotFoundError: # Specific exception
             error_msg = f"Preview Error: File not found\n{pdf_path}"
             print(error_msg)
             self.clear_pdf_preview(error_msg)
             self.pdf_preview_image = None
        except Exception as e:
            error_msg = f"Error previewing PDF:\n{os.path.basename(pdf_path)}\n{type(e).__name__}: {e}"
            import traceback
            print(f"--- PDF Preview Exception Traceback ---")
            traceback.print_exc()
            print(f"--- End Traceback ---")
            self.clear_pdf_preview(error_msg) # Show error in GUI
            self.pdf_preview_image = None # Ensure reference is cleared on error

    def _update_canvas_scrollregion(self):
        """ Safely update canvas scroll region based on the pdf image. """
        try:
            bbox = None
            if self._canvas_image_id:
                # Check if canvas item still exists
                if self.pdf_canvas.find_withtag(self._canvas_image_id):
                    bbox = self.pdf_canvas.bbox(self._canvas_image_id)
                else:
                    # Item might have been deleted between preview calls
                    self._canvas_image_id = None # Reset ID
                    #print("Canvas item ID was invalid during scrollregion update.") # DEBUG

            if bbox:
                # Add padding to the bounding box
                padded_bbox = (bbox[0]-5, bbox[1]-5, bbox[2] + 10, bbox[3] + 10)
                self.pdf_canvas.config(scrollregion=padded_bbox)
                #print(f"Canvas scrollregion updated to: {padded_bbox}") # DEBUG (verbose)
            else:
                # If no image or bbox fails, fit to current canvas size
                current_width = self.pdf_canvas.winfo_width()
                current_height = self.pdf_canvas.winfo_height()
                self.pdf_canvas.config(scrollregion=(0, 0, current_width, current_height))
                #print("Canvas scrollregion reset to canvas size (no valid image bbox).") # DEBUG (verbose)
        except Exception as e:
            print(f"Error updating canvas scrollregion: {e}")


# --- Main Execution ---
if __name__ == "__main__":
    # Optional: Check Pillow on startup and inform user if missing
    if not PIL_AVAILABLE:
         root_check = tk.Tk()
         root_check.withdraw() # Hide the temporary window
         messagebox.showwarning("Dependency Warning",
                                "Python Imaging Library (Pillow) not found.\n"
                                "PDF Preview functionality will be disabled.\n\n"
                                "Please install it using:\n"
                                "pip install Pillow")
         root_check.destroy()

    root = tk.Tk()
    app = InvoiceProcessorApp(root)
    root.mainloop()





