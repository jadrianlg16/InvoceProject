# gui.py
# Contains the Tkinter GUI class (InvoiceProcessorApp).

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
import pandas as pd # Needed for DataFrame checks and manipulation within GUI methods
import copy # For deep copying clipboard data
import fitz # Needed for preview rendering

# --- Import logic functions ---
import logic # Import the logic module

# --- Pillow Import and Check ---
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("WARNING: Pillow library not found (pip install Pillow). PDF preview will be disabled.")


# --- GUI Class ---
class InvoiceProcessorApp:
    def __init__(self, master):
        self.master = master
        master.title("Invoice PDF Processor v2.5 - Interactive Sort & Scroll Memory")
        master.geometry("1400x850")

        # --- Pass messagebox module to instance for logic thread ---
        # This allows the logic thread to show message boxes in the main GUI thread
        self.messagebox = messagebox

        self.selected_folder = tk.StringVar(value="No folder selected")
        self.processing_active = False
        self.dataframe = None
        self.current_excel_file = None
        self.pdf_preview_image = None
        self._canvas_image_id = None
        self._pdf_path_map = {}
        self._df_index_map = {}
        self._placeholder_window_id = None
        self._edit_entry = None
        self._edit_item_id = None
        self._edit_column_id = None
        self._clipboard = [] # List of dictionaries for copied rows

        # Zoom State
        self.current_zoom_factor = 1.0
        self.zoom_step = 1.2
        self.min_zoom = 0.1
        self.max_zoom = 5.0
        self.current_preview_pdf_path = None

        # Scroll Position Memory
        self.pdf_scroll_positions = {} # {pdf_path: (x_fraction, y_fraction)}

        # Sorting State
        self._tree_current_sort_col = None
        self._tree_current_sort_ascending = True

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

        # --- Main Content Area (Paned Window) ---
        self.content_pane = ttk.PanedWindow(master, orient=tk.HORIZONTAL)
        self.content_pane.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(5, 5))

        # --- Left Panel (Treeview and Controls) ---
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

        # Define Treeview headings and columns
        self.tree.heading("Source PDF", text="Source PDF", command=lambda c="Source PDF": self.sort_treeview_column(c))
        self.tree.heading("Invoice Folio", text="Invoice Folio", command=lambda c="Invoice Folio": self.sort_treeview_column(c))
        self.tree.heading("Document Type", text="Document Type", command=lambda c="Document Type": self.sort_treeview_column(c))
        self.tree.heading("Reference Number", text="Reference Number", command=lambda c="Reference Number": self.sort_treeview_column(c))

        self.tree.column("Source PDF", anchor=tk.W, width=220, stretch=tk.NO)
        self.tree.column("Invoice Folio", anchor=tk.W, width=100)
        self.tree.column("Document Type", anchor=tk.W, width=150)
        self.tree.column("Reference Number", anchor=tk.W, width=120)

        self.editable_columns = ["Invoice Folio", "Document Type", "Reference Number"]

        # Bind events
        self.tree.bind('<<TreeviewSelect>>', self.on_treeview_select)
        self.tree.bind('<Double-1>', self.on_tree_double_click)
        self.tree.bind('<Button-3>', self.show_context_menu)

        # Context Menu
        self.context_menu = tk.Menu(master, tearoff=0)
        self.context_menu.add_command(label="Add Blank Row", command=self._add_row)
        self.context_menu.add_command(label="Copy Selected Row(s)", command=self._copy_selected_rows)
        self.context_menu.add_command(label="Paste Row(s)", command=self._paste_rows)
        self.context_menu.add_command(label="Delete Selected Row(s)", command=self._delete_selected_rows)

        # --- Right Panel (PDF Preview) ---
        pdf_preview_frame = ttk.Frame(self.content_pane, padding=5)
        self.content_pane.add(pdf_preview_frame, weight=3)

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

        # Canvas Frame
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

        # Placeholder Label
        self.pdf_placeholder_label = ttk.Label(self.pdf_canvas, text="", style="Placeholder.TLabel")
        self.pdf_canvas.bind('<Configure>', self._center_placeholder)

        # --- Bottom Bar (Status Log) ---
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
        self.clear_pdf_preview("Select a row above to preview PDF")
        self.update_status("Ready. Please select a folder and invoice type.")
        if not PIL_AVAILABLE:
             self.update_status("WARNING: Pillow library not found. PDF Preview is disabled.")
             self.clear_pdf_preview("PDF Preview disabled:\nPillow library not installed.")

    def _center_placeholder(self, event=None):
        """Centers the placeholder label within the PDF canvas."""
        if self._placeholder_window_id and self.pdf_canvas.winfo_exists() and \
           self._placeholder_window_id in self.pdf_canvas.find_withtag("placeholder"):
            canvas_w = self.pdf_canvas.winfo_width()
            canvas_h = self.pdf_canvas.winfo_height()
            self.pdf_canvas.coords(self._placeholder_window_id, canvas_w//2, canvas_h//2)

    def select_folder(self):
        """Opens a dialog to select a folder."""
        if self.processing_active: return

        if self.dataframe is not None and self.current_excel_file:
             if not messagebox.askokcancel("Confirm Folder Change",
                                        "Changing the folder will clear the current data and edits.\n"
                                        "Any unsaved changes will be lost.\n\nProceed?", parent=self.master):
                  return

        folder = filedialog.askdirectory()
        if folder:
            normalized_folder = os.path.normpath(folder)
            self.selected_folder.set(normalized_folder)
            self.clear_status()
            self.clear_treeview()
            self.pdf_scroll_positions.clear()
            self.clear_pdf_preview("Select a row to preview PDF")
            self.current_zoom_factor = 1.0
            self.current_preview_pdf_path = None
            self.update_status(f"Folder selected: {normalized_folder}")
            self.update_status("Ready to process.")
            self.save_changes_button.config(state=tk.DISABLED)
            self._tree_current_sort_col = None
            self._tree_current_sort_ascending = True
        else:
            if self.selected_folder.get() != "No folder selected":
                self.update_status("Folder selection cancelled.")

    def clear_status(self):
         """Clears the status log."""
         self.status_text.config(state=tk.NORMAL)
         self.status_text.delete('1.0', tk.END)
         self.status_text.config(state=tk.DISABLED)

    def update_status(self, message):
        """Appends a message to the status log (thread-safe)."""
        if threading.current_thread() != threading.main_thread():
            self.master.after(0, self._update_status_text, message)
        else:
            self._update_status_text(message)

    def _update_status_text(self, message):
        """Internal helper to update the status text widget."""
        current_state = self.status_text.cget('state')
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state=current_state)

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
        self.save_changes_button.config(state=tk.DISABLED)
        self._cancel_cell_edit()

    def enable_buttons(self):
        """Enables buttons after processing."""
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
        """Initiates PDF processing in a background thread."""
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
        self.pdf_scroll_positions.clear()
        self.clear_pdf_preview(f"Processing {invoice_type} invoices...\nPlease wait.")
        self.update_status(f"Starting recursive processing for {invoice_type} in: {folder}")
        self.update_status("-" * 40)

        self._tree_current_sort_col = None
        self._tree_current_sort_ascending = True

        # --- Call the logic function in a thread ---
        # Pass 'self' (the GUI instance) to the logic function
        process_thread = threading.Thread(target=logic.run_processing, args=(folder, invoice_type, self), daemon=True)
        process_thread.start()

    def set_data_and_file(self, df, excel_file_path):
        """Callback for logic thread to set data and file path."""
        self.dataframe = df
        self.current_excel_file = excel_file_path
        self._clipboard = []
        if self.current_excel_file and self.dataframe is not None and not self.dataframe.empty:
             self.save_changes_button.config(state=tk.NORMAL)
        else:
             self.save_changes_button.config(state=tk.DISABLED)

    def clear_treeview(self):
        """Clears Treeview, maps, DataFrame, and resets state."""
        if self._edit_entry: self._cancel_cell_edit()

        if hasattr(self, 'tree'):
            try: self.tree.unbind('<Button-3>')
            except tk.TclError: pass
            for item in self.tree.get_children():
                try: self.tree.delete(item)
                except tk.TclError: pass
            try: self.tree.bind('<Button-3>', self.show_context_menu)
            except tk.TclError: pass

        self.dataframe = None
        self.current_excel_file = None
        self._pdf_path_map.clear()
        self._df_index_map.clear()
        self._clipboard = []

        if hasattr(self, 'tree'):
             for col in self.tree["columns"]:
                 try:
                      current_text = self.tree.heading(col, "text").replace(' ▲', '').replace(' ▼', '')
                      self.tree.heading(col, text=current_text)
                 except tk.TclError: pass

        self._tree_current_sort_col = None
        self._tree_current_sort_ascending = True

        if hasattr(self, 'save_changes_button'):
            self.save_changes_button.config(state=tk.DISABLED)

    def load_data_to_treeview(self):
        """Populates the Treeview from self.dataframe."""
        if self._edit_entry: self._cancel_cell_edit()

        # Clear existing items and maps
        self.tree.unbind('<Button-3>')
        for item in self.tree.get_children():
            try: self.tree.delete(item)
            except tk.TclError: pass
        self.tree.bind('<Button-3>', self.show_context_menu)
        self._pdf_path_map.clear()
        self._df_index_map.clear()

        if self.dataframe is None or self.dataframe.empty:
             self.update_status("No data available to display.")
             self.clear_pdf_preview("No data loaded.")
             self.save_changes_button.config(state=tk.DISABLED)
             return

        display_columns = list(self.tree["columns"])
        required_cols = display_columns + ["Full PDF Path"]
        if not all(col in self.dataframe.columns for col in required_cols):
            missing = [col for col in required_cols if col not in self.dataframe.columns]
            errmsg = f"Error: DataFrame is missing required columns: {missing}"
            self.update_status(errmsg)
            messagebox.showerror("Data Loading Error", errmsg, parent=self.master)
            self.clear_pdf_preview("Error loading data. See status log.")
            self.dataframe = None
            self.current_excel_file = None
            self._clipboard = []
            self.save_changes_button.config(state=tk.DISABLED)
            return

        self.tree.configure(displaycolumns=display_columns)

        # Reset index if not unique (important for using index as iid)
        if not self.dataframe.index.is_unique:
             self.update_status("Warning: DataFrame index is not unique. Resetting index.")
             self.dataframe.reset_index(drop=True, inplace=True)

        # Populate Treeview
        for df_index, row in self.dataframe.iterrows():
            try:
                display_values = tuple(str(row[col]) if pd.notna(row[col]) else "" for col in display_columns)
                full_path = row["Full PDF Path"]
                item_id = str(df_index) # Use DataFrame index as Treeview item ID

                self.tree.insert("", tk.END, values=display_values, iid=item_id)

                if full_path and isinstance(full_path, str):
                    self._pdf_path_map[item_id] = full_path
                self._df_index_map[item_id] = df_index

            except Exception as e:
                print(f"Error adding row with DataFrame index {df_index} to treeview: {e}")
                self.update_status(f"Warning: Could not display row for PDF '{row.get('Source PDF', 'Unknown')}' in table.")

        row_count = len(self.tree.get_children())
        if row_count == 0 and not self.dataframe.empty:
            self.update_status("Loaded data frame, but no rows were added to the table. Check data integrity.")
            self.clear_pdf_preview("Data loaded, but no rows to display.")
        elif row_count == 0 and self.dataframe.empty:
             self.update_status("Data loaded, but it contains no rows.")
             self.clear_pdf_preview("No data loaded.")

        # Update save button state
        if self.current_excel_file and not self.dataframe.empty:
            self.save_changes_button.config(state=tk.NORMAL)
        else:
             self.save_changes_button.config(state=tk.DISABLED if self.dataframe.empty else tk.NORMAL if self.current_excel_file else tk.DISABLED)

    # --- Treeview Sorting ---
    def sort_treeview_column(self, col):
        """Sorts the Treeview based on the clicked column header."""
        if self.dataframe is None or self.dataframe.empty: return
        if self._edit_entry: self._cancel_cell_edit()

        try:
            sort_ascending = not self._tree_current_sort_ascending if self._tree_current_sort_col == col else True
            sorted_df = self.dataframe.copy()
            numeric_col_check = pd.to_numeric(sorted_df[col], errors='coerce')
            is_numeric_type = pd.api.types.is_numeric_dtype(numeric_col_check)
            sort_numerically = not numeric_col_check.isna().all() and is_numeric_type

            if sort_numerically:
                temp_sort_col = f"__{col}_numeric_sort"
                sorted_df[temp_sort_col] = numeric_col_check
                sorted_df = sorted_df.sort_values(by=temp_sort_col, ascending=sort_ascending, na_position='last')
                sorted_df.drop(columns=[temp_sort_col], inplace=True)
            else:
                sorted_df = sorted_df.sort_values(
                    by=col, ascending=sort_ascending,
                    key=lambda x: x.map(lambda s: str(s).lower() if pd.notna(s) else ''),
                    na_position='last'
                )

            self.dataframe = sorted_df
            self._tree_current_sort_col = col
            self._tree_current_sort_ascending = sort_ascending

            for c in self.tree["columns"]:
                current_text = self.tree.heading(c, "text").replace(' ▲', '').replace(' ▼', '')
                indicator = ' ▲' if sort_ascending else ' ▼'
                self.tree.heading(c, text=current_text + (indicator if c == col else ''))

            self.load_data_to_treeview()

        except Exception as e:
            errmsg = f"Error sorting treeview column '{col}': {e}"
            print(errmsg)
            import traceback; traceback.print_exc()
            messagebox.showerror("Sort Error", f"Could not sort column '{col}':\n{e}", parent=self.master)
            self._tree_current_sort_col = None
            self._tree_current_sort_ascending = True

    # --- Inline Editing ---
    def on_tree_double_click(self, event):
        """Handles double-click for inline editing."""
        if self._edit_entry: return
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell": return
        item_id = self.tree.identify_row(event.y)
        column_id_str = self.tree.identify_column(event.x)
        if not item_id or not column_id_str: return

        try:
            column_index = int(column_id_str.replace('#', '')) - 1
            if column_index < 0 or column_index >= len(self.tree["columns"]): return
            column_name = self.tree["columns"][column_index]
        except (ValueError, IndexError): return

        if column_name not in self.editable_columns: return
        try: bbox = self.tree.bbox(item_id, column=column_id_str)
        except Exception: return
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
        """Saves the edited cell value."""
        if not self._edit_entry: return
        new_value = self._edit_entry.get()
        self._cancel_cell_edit()

        try: current_values = list(self.tree.item(item_id, 'values'))
        except tk.TclError: return # Item deleted

        if column_index < len(current_values) and str(current_values[column_index]) != new_value:
            current_values[column_index] = new_value
            try: self.tree.item(item_id, values=tuple(current_values))
            except tk.TclError: pass # Item deleted

            if self.dataframe is not None:
                try:
                    df_index = self._df_index_map.get(item_id)
                    if df_index is not None and df_index in self.dataframe.index:
                        self.dataframe.loc[df_index, column_name] = new_value
                        self.update_status(f"Cell updated: Row index {df_index}, Col '{column_name}' = '{new_value}'")
                        if self.current_excel_file: self.save_changes_button.config(state=tk.NORMAL)
                    else:
                        print(f"Error: Could not find DataFrame index ({df_index}) for tree item {item_id} during edit save.")
                        self.update_status(f"Error: Failed to update backing data for edit (Index mapping issue).")
                except Exception as e:
                    print(f"Error updating DataFrame at index {df_index}, column {column_name}: {e}")
                    self.update_status(f"Error: Failed to update backing data for edit ({type(e).__name__}).")

    def _cancel_cell_edit(self, event=None):
        """Cancels the current cell edit."""
        if self._edit_entry:
            try: self._edit_entry.destroy()
            except tk.TclError: pass
            self._edit_entry = None
            self._edit_item_id = None
            self._edit_column_id = None

    # --- Context Menu and Row Operations ---
    def show_context_menu(self, event):
        """Displays the right-click context menu."""
        item_id = self.tree.identify_row(event.y)
        if item_id and item_id not in self.tree.selection():
            self.tree.selection_set(item_id)

        has_selection = bool(self.tree.selection())
        can_paste = bool(self._clipboard)
        can_add = self.dataframe is not None

        self.context_menu.entryconfig("Add Blank Row", state=tk.NORMAL if can_add else tk.DISABLED)
        self.context_menu.entryconfig("Copy Selected Row(s)", state=tk.NORMAL if has_selection else tk.DISABLED)
        self.context_menu.entryconfig("Paste Row(s)", state=tk.NORMAL if can_paste else tk.DISABLED)
        self.context_menu.entryconfig("Delete Selected Row(s)", state=tk.NORMAL if has_selection else tk.DISABLED)

        try: self.context_menu.tk_popup(event.x_root, event.y_root)
        finally: self.context_menu.grab_release()

    def _add_row(self):
        """Adds a blank row."""
        if self.dataframe is None:
            messagebox.showwarning("Add Row Failed", "Load data first.", parent=self.master)
            return
        if self._edit_entry: self._cancel_cell_edit()

        try:
            new_row_data = {col: "" for col in self.dataframe.columns}
            # Sensible defaults for core columns
            new_row_data["Source PDF"] = "[MANUALLY ADDED]"
            new_row_data["Full PDF Path"] = ""
            new_row_data["Invoice Folio"] = ""
            new_row_data["Document Type"] = ""
            new_row_data["Reference Number"] = ""

            new_row_df = pd.DataFrame([new_row_data])
            self.dataframe = pd.concat([self.dataframe, new_row_df], ignore_index=True)
            self.load_data_to_treeview() # Reloads and resets index maps

            children = self.tree.get_children()
            if children: # Select and focus the newly added row
                new_item_id = children[-1]
                self.tree.see(new_item_id)
                self.tree.selection_set(new_item_id)
                self.tree.focus(new_item_id)

            self.update_status("Added 1 blank row.")
            if self.current_excel_file: self.save_changes_button.config(state=tk.NORMAL)

        except Exception as e:
            print(f"Error adding row: {e}")
            self.update_status(f"Error: Could not add row ({type(e).__name__}).")
            import traceback; traceback.print_exc()
            try: self.load_data_to_treeview() # Attempt recovery
            except Exception as reload_e: print(f"Error reloading treeview after add row error: {reload_e}")

    def _delete_selected_rows(self):
        """Deletes selected rows."""
        selected_item_ids = self.tree.selection()
        if not selected_item_ids: return
        num_selected = len(selected_item_ids)
        if not messagebox.askyesno("Confirm Delete", f"Delete {num_selected} selected row(s)?", parent=self.master):
            return
        if self._edit_entry: self._cancel_cell_edit()

        try:
            indices_to_drop = []
            failed_map_ids = []
            for item_id in selected_item_ids:
                df_index = self._df_index_map.get(item_id)
                if df_index is not None and df_index in self.dataframe.index:
                    indices_to_drop.append(df_index)
                else: failed_map_ids.append(item_id)

            if failed_map_ids: self.update_status(f"Warning: Could not map {len(failed_map_ids)} item(s) for deletion.")
            if not indices_to_drop:
                 self.update_status("Error: No valid data found for selected rows to delete.")
                 return

            self.dataframe.drop(index=indices_to_drop, inplace=True)
            self.dataframe.reset_index(drop=True, inplace=True) # Reset index after drop
            self.load_data_to_treeview() # Reload to update tree and maps
            self.update_status(f"Deleted {len(indices_to_drop)} row(s).")

            if self.current_excel_file and self.dataframe is not None and not self.dataframe.empty:
                 self.save_changes_button.config(state=tk.NORMAL)
            else: # Disable save if empty or no file path
                 self.save_changes_button.config(state=tk.DISABLED)

        except Exception as e:
            print(f"Error deleting rows: {e}")
            self.update_status(f"Error: Could not delete rows ({type(e).__name__}).")
            import traceback; traceback.print_exc()
            try: self.load_data_to_treeview() # Attempt recovery
            except Exception as reload_e: print(f"Error reloading treeview after delete rows error: {reload_e}")

    def _copy_selected_rows(self):
        """Copies selected rows' data to internal clipboard."""
        selected_item_ids = self.tree.selection()
        if not selected_item_ids or self.dataframe is None: return

        try:
            df_indices_to_copy = []
            failed_map_ids = []
            for item_id in selected_item_ids:
                 df_index = self._df_index_map.get(item_id)
                 if df_index is not None and df_index in self.dataframe.index:
                      df_indices_to_copy.append(df_index)
                 else: failed_map_ids.append(item_id)

            if failed_map_ids: self.update_status(f"Warning: Could not map {len(failed_map_ids)} item(s) for copying.")
            if not df_indices_to_copy:
                 self.update_status("Error: Could not map any selected items to data for copying.")
                 self._clipboard = []
                 return

            copied_data_df = self.dataframe.loc[df_indices_to_copy]
            # Use deepcopy to prevent modifying original data if clipboard dicts are changed later
            self._clipboard = copy.deepcopy(copied_data_df.to_dict('records'))
            copied_count = len(self._clipboard)
            self.update_status(f"Copied {copied_count} row(s) to clipboard.")

        except Exception as e:
            print(f"Error copying rows: {e}")
            self.update_status(f"Error: Could not copy rows ({type(e).__name__}).")
            self._clipboard = []
            import traceback; traceback.print_exc()

    def _paste_rows(self):
        """Pastes rows from clipboard."""
        if not self._clipboard or self.dataframe is None:
            self.update_status("Clipboard empty or no data loaded. Cannot paste.")
            return
        if self._edit_entry: self._cancel_cell_edit()

        try:
            pasted_df = pd.DataFrame(self._clipboard)
            if pasted_df.empty: return

            # Align columns, filling missing ones with ""
            pasted_df = pasted_df.reindex(columns=self.dataframe.columns, fill_value="")
            if pasted_df.empty: return

            self.dataframe = pd.concat([self.dataframe, pasted_df], ignore_index=True)
            num_pasted = len(pasted_df)
            self.load_data_to_treeview() # Reloads tree and resets maps

            # Select the newly pasted rows
            children = self.tree.get_children()
            if len(children) >= num_pasted:
                 first_pasted_item_id = children[-num_pasted]
                 self.tree.selection_set(children[-num_pasted:])
                 self.tree.see(first_pasted_item_id)

            self.update_status(f"Pasted {num_pasted} row(s) from clipboard.")
            if self.current_excel_file: self.save_changes_button.config(state=tk.NORMAL)

        except Exception as e:
            print(f"Error pasting rows: {e}")
            self.update_status(f"Error: Could not paste rows ({type(e).__name__}).")
            import traceback; traceback.print_exc()
            try: self.load_data_to_treeview() # Attempt recovery
            except Exception as reload_e: print(f"Error reloading treeview after paste rows error: {reload_e}")

    # --- Excel Saving Method ---
    def save_changes_to_excel(self):
        """Saves the current DataFrame to the Excel file."""
        if self._edit_entry: self._cancel_cell_edit()
        if self.dataframe is None or self.dataframe.empty:
            messagebox.showwarning("No Data", "Nothing to save.", parent=self.master)
            return
        if not self.current_excel_file:
            messagebox.showerror("Save Error", "Excel file path unknown. Cannot save.", parent=self.master)
            return
        excel_dir = os.path.dirname(self.current_excel_file)
        if not os.path.isdir(excel_dir):
            messagebox.showerror("Save Error", f"Directory not found:\n{excel_dir}", parent=self.master)
            return

        if not messagebox.askyesno("Confirm Save", f"Overwrite:\n{self.current_excel_file}?", parent=self.master):
            return

        self.update_status(f"Saving changes to: {self.current_excel_file}")
        try:
            excel_columns = ["Source PDF", "Invoice Folio", "Document Type", "Reference Number"]
            cols_to_save = [col for col in excel_columns if col in self.dataframe.columns]
            if not cols_to_save: raise ValueError("No valid columns to save.")

            df_to_save = self.dataframe[cols_to_save].copy()
            df_to_save.to_excel(self.current_excel_file, index=False, engine='openpyxl')

            success_msg = f"Changes saved to:\n{self.current_excel_file}"
            self.update_status(success_msg)
            messagebox.showinfo("Save Successful", success_msg, parent=self.master)

        except PermissionError:
            error_message = (f"Error saving:\n{self.current_excel_file}\n\nPermission denied. Is it open?")
            self.update_status(f"Save Error: Permission Denied.")
            messagebox.showerror("Save Error", error_message, parent=self.master)
        except Exception as e:
            error_message = f"Error saving to Excel:\n\n{e}"
            print(error_message)
            import traceback; traceback.print_exc()
            self.update_status(f"Save Error: {e}. See console.")
            messagebox.showerror("Save Error", error_message, parent=self.master)

    # --- Treeview Selection & PDF Preview ---
    def _store_current_scroll_position(self):
        """Stores the current PDF canvas scroll position."""
        if self.current_preview_pdf_path and hasattr(self, 'pdf_canvas') and self.pdf_canvas.winfo_exists():
            try:
                x_frac, y_frac = self.pdf_canvas.xview()[0], self.pdf_canvas.yview()[0]
                # Store position, even if 0,0 to remember it was viewed
                self.pdf_scroll_positions[self.current_preview_pdf_path] = (x_frac, y_frac)
            except tk.TclError as e: print(f"Warning: TclError getting scroll position: {e}")
            except Exception as e: print(f"Error getting scroll position: {e}")

    def on_treeview_select(self, event):
        """Handles Treeview selection changes to update PDF preview."""
        if self._edit_entry and self.tree.focus() != self._edit_entry:
             try: # Attempt to save edit if focus moves away
                 col_idx = int(self._edit_column_id.replace('#', '')) - 1
                 col_name = self.tree["columns"][col_idx]
                 self._save_cell_edit(self._edit_item_id, col_idx, col_name)
             except Exception: self._cancel_cell_edit()

        self._store_current_scroll_position() # Store scroll of *previous* PDF

        selected_items = self.tree.selection()
        if not selected_items: return

        selected_item_id = selected_items[0] # Preview first selected item
        pdf_full_path = self._pdf_path_map.get(selected_item_id)

        if pdf_full_path and isinstance(pdf_full_path, str):
            if os.path.exists(pdf_full_path):
                if pdf_full_path != self.current_preview_pdf_path:
                    self.update_pdf_preview(pdf_full_path)
                # If same path, scroll position is handled by _update_canvas_scrollregion
            else:
                err_msg = f"File Not Found:\n{os.path.basename(pdf_full_path)}"
                self.clear_pdf_preview(err_msg)
                self.current_preview_pdf_path = None
        elif pdf_full_path is None and selected_item_id in self._df_index_map:
             # Row exists but has no associated PDF path (e.g., manually added)
             self.clear_pdf_preview("No PDF associated with this row.")
             self.current_preview_pdf_path = None
        else: # Item ID not in map or path is invalid
             self.clear_pdf_preview("Error: Cannot find data or path\nfor selected row.")
             self.current_preview_pdf_path = None

    def clear_pdf_preview(self, message="Select a row to preview PDF"):
        """Clears PDF preview and shows a message."""
        self._store_current_scroll_position() # Store scroll before clearing

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
            canvas_w//2, canvas_h//2, window=self.pdf_placeholder_label,
            anchor=tk.CENTER, tags="placeholder"
        )

        self.master.after_idle(self._update_canvas_scrollregion, None) # Update scroll, pass None path
        self.current_preview_pdf_path = None

    def update_pdf_preview(self, pdf_path):
        """Loads and displays the first page of the PDF."""
        if not PIL_AVAILABLE:
            self.clear_pdf_preview("PDF Preview disabled:\nPillow library not found.")
            return

        # Clear previous content
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
            if page.rect.width == 0 or page.rect.height == 0: raise ValueError("PDF page has zero dimensions.")

            mat = fitz.Matrix(self.current_zoom_factor, self.current_zoom_factor)
            pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB, alpha=False)
            doc.close()

            img_bytes = pix.samples
            if not isinstance(img_bytes, bytes): img_bytes = bytes(img_bytes)
            if not img_bytes: raise ValueError("Pixmap samples are empty.")

            pil_image = Image.frombytes("RGB", [pix.width, pix.height], img_bytes)
            self.pdf_preview_image = ImageTk.PhotoImage(image=pil_image)

            if self.pdf_preview_image:
                self.current_preview_pdf_path = pdf_path # Set path *before* scheduling scroll update
                self._canvas_image_id = self.pdf_canvas.create_image(0, 0, anchor=tk.NW, image=self.pdf_preview_image, tags="pdf_image")
                self.master.after_idle(self._update_canvas_scrollregion, pdf_path) # Schedule scroll update/restore
            else: raise ValueError("Failed to create PhotoImage.")

        except Exception as e:
            base_name = os.path.basename(pdf_path) if pdf_path else "Unknown"
            error_type = type(e).__name__
            error_msg = f"Preview Error ({error_type}):\n{base_name}"
            if isinstance(e, fitz.fitz.PasswordError): error_msg += "\n(Password Protected?)"
            elif isinstance(e, ValueError): error_msg += f"\n({e})"
            import traceback; print(f"--- PDF Preview Exception: {base_name} ---"); traceback.print_exc(); print("--- End Traceback ---")
            self.clear_pdf_preview(error_msg) # Also resets current_preview_pdf_path

    # --- Zoom Methods ---
    def zoom_in(self):
        if self.current_preview_pdf_path:
            self._store_current_scroll_position()
            new_zoom = min(self.current_zoom_factor * self.zoom_step, self.max_zoom)
            if new_zoom != self.current_zoom_factor:
                self.current_zoom_factor = new_zoom
                self.update_pdf_preview(self.current_preview_pdf_path)

    def zoom_out(self):
        if self.current_preview_pdf_path:
            self._store_current_scroll_position()
            new_zoom = max(self.current_zoom_factor / self.zoom_step, self.min_zoom)
            if new_zoom != self.current_zoom_factor:
                self.current_zoom_factor = new_zoom
                self.update_pdf_preview(self.current_preview_pdf_path)

    def reset_zoom(self):
        if self.current_preview_pdf_path and self.current_zoom_factor != 1.0:
            self._store_current_scroll_position()
            self.current_zoom_factor = 1.0
            self.update_pdf_preview(self.current_preview_pdf_path)

    def _update_canvas_scrollregion(self, pdf_path_to_restore=None):
        """Updates canvas scroll region and restores scroll if applicable."""
        try:
            scroll_bbox = None
            if self._canvas_image_id and self.pdf_canvas.find_withtag(self._canvas_image_id):
                bbox = self.pdf_canvas.bbox(self._canvas_image_id)
                if bbox: scroll_bbox = (bbox[0], bbox[1], bbox[2] + 5, bbox[3] + 5)
            else: # No image, use canvas size
                scroll_bbox = (0, 0, max(1, self.pdf_canvas.winfo_width()), max(1, self.pdf_canvas.winfo_height()))

            if scroll_bbox: self.pdf_canvas.config(scrollregion=scroll_bbox)

            # Restore scroll position if this path matches the one we intended to restore
            if pdf_path_to_restore and pdf_path_to_restore == self.current_preview_pdf_path:
                pos = self.pdf_scroll_positions.get(pdf_path_to_restore, (0.0, 0.0))
                self.pdf_canvas.xview_moveto(pos[0])
                self.pdf_canvas.yview_moveto(pos[1])
            elif pdf_path_to_restore is None: # Called from clear_pdf_preview
                 self.pdf_canvas.xview_moveto(0.0)
                 self.pdf_canvas.yview_moveto(0.0)

        except tk.TclError as e: print(f"Warning: TclError updating canvas scrollregion: {e}")
        except Exception as e: print(f"Error updating canvas scrollregion: {e}")


