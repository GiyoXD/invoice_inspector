import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import sys
import pandas as pd
from pathlib import Path
import json

# Import the pipeline logic
# Ensure directory is in path if needed, but since we are in root, it should be fine
try:
    from extract_invoice_data import run_pipeline
except ImportError:
    # Fallback if run from different dir
    sys.path.append(str(Path(__file__).parent))
    from extract_invoice_data import run_pipeline

class InvoiceInspectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Inspector")
        self.root.geometry("1100x700")
        
        # Style
        style = ttk.Style()
        style.theme_use('clam')
        
        # Variables
        self.folder_path = tk.StringVar()
        self.master_path = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")
        
        # Auto-detect defaults
        self.ensure_defaults()

        # --- UI Layout ---
        
        # 1. Top Frame: Controls
        control_frame = ttk.LabelFrame(self.root, text="Configuration", padding=10)
        control_frame.pack(fill="x", padx=10, pady=5)
        
        # Folder Selection
        ttk.Label(control_frame, text="Invoice Folder:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(control_frame, textvariable=self.folder_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(control_frame, text="Browse", command=self.browse_folder).grid(row=0, column=2, padx=5, pady=5)
        
        # Master List Selection
        ttk.Label(control_frame, text="Master List:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(control_frame, textvariable=self.master_path, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(control_frame, text="Browse", command=self.browse_master).grid(row=1, column=2, padx=5, pady=5)
        ttk.Button(control_frame, text="Load / Edit", command=self.load_master_editor).grid(row=1, column=3, padx=5, pady=5)

        # Run Button
        self.run_btn = ttk.Button(control_frame, text="RUN INSPECTION", command=self.start_inspection_thread)
        self.run_btn.grid(row=2, column=1, pady=10, sticky="ew")

        # 2. Main Content: Split Pane
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Tab 1: Inspection Results
        results_frame = ttk.Frame(self.notebook)
        self.notebook.add(results_frame, text="Inspection Results")
        
        # PanedWindow for Master-Detail View
        self.paned = ttk.PanedWindow(results_frame, orient=tk.VERTICAL)
        self.paned.pack(fill="both", expand=True)
        
        # Top Pane: List (Treeview)
        list_frame = ttk.Frame(self.paned)
        self.paned.add(list_frame, weight=3) # Give more weight to list
        
        # Treeview for Results
        cols = ("ID", "Status", "Pcs", "Qty(Sqft)", "Amount", "Pallets", "Net W", "Gross W", "CBM", "Details")
        self.tree = ttk.Treeview(list_frame, columns=cols, show='headings')
        
        for col in cols:
            self.tree.heading(col, text=col)
            if col == "Details":
                self.tree.heading(col, text=col, anchor="w")
                self.tree.column(col, width=600, minwidth=250, stretch=True, anchor="w")
            elif col == "ID":
                self.tree.column(col, width=120)
            else:
                self.tree.column(col, width=80)
            
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        x_scrollbar = ttk.Scrollbar(list_frame, orient="horizontal", command=self.tree.xview)
        
        self.tree.configure(yscroll=scrollbar.set, xscroll=x_scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        x_scrollbar.pack(side="bottom", fill="x")

        # Bottom Pane: Detail View
        detail_frame = ttk.LabelFrame(self.paned, text="Full Details (Text Wrapped)")
        self.paned.add(detail_frame, weight=1)
        
        self.details_text = tk.Text(detail_frame, wrap="word", height=6)
        self.details_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Bind Selection
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        
        # Status Bar
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor="w")
        status_bar.pack(side="bottom", fill="x")

        # Tab 2: Master List Editor
        self.editor_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.editor_frame, text="Master List Editor")
        
        # Editor Controls
        editor_controls = ttk.Frame(self.editor_frame)
        editor_controls.pack(fill="x", padx=5, pady=5)
        
        ttk.Button(editor_controls, text="Save Changes", command=self.save_master_list).pack(side="left", padx=5)
        ttk.Button(editor_controls, text="Reload File", command=self.load_master_editor).pack(side="left", padx=5)
        ttk.Button(editor_controls, text="Add New Row", command=self.add_master_row).pack(side="left", padx=5)
        ttk.Button(editor_controls, text="Delete Row", command=self.delete_master_row).pack(side="left", padx=5)
        ttk.Button(editor_controls, text="Paste Data", command=self.parse_and_preview_paste).pack(side="left", padx=5)
        
        # Editor Grid
        tree_frame = ttk.Frame(self.editor_frame)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.editor_cols = ["Invoice No", "Amount", "Quantity", "Pallets", "PCS", "NetWeight", "GrossWeight", "CBM", "ID", "VERIFY STATE", "DIFF_PALLET", "DIFF_SQFT", "DIFF_AMOUNT", "DIFF_PCS", "DIFF_NET", "DIFF_GROSS", "DIFF_CBM"]
        self.editor_tree = ttk.Treeview(tree_frame, columns=self.editor_cols, show='headings')
        
        for col in self.editor_cols:
            self.editor_tree.heading(col, text=col)
            self.editor_tree.column(col, width=100)
            
        e_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.editor_tree.yview)
        e_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.editor_tree.xview)
        
        self.editor_tree.configure(yscroll=e_scroll.set, xscroll=e_scroll_x.set)
        
        self.editor_tree.grid(row=0, column=0, sticky="nsew")
        e_scroll.grid(row=0, column=1, sticky="ns")
        e_scroll_x.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Bind Paste
        self.editor_tree.bind("<Control-v>", lambda e: self.parse_and_preview_paste())
        
        self.editor_tree.bind("<Double-1>", self.on_editor_double_click)
        self.editor_tree.bind("<ButtonRelease-1>", self.on_editor_click)
        self.master_df = None
        
        # Paste Logic Variables
        self.selected_col_idx = 0
        self.selected_row_id = None

    def on_tree_select(self, event):
        """Updates the Detail View with the full text of the selected row's Details."""
        selected_items = self.tree.selection()
        if not selected_items:
            return
            
        # Clear existing text
        self.details_text.delete(1.0, tk.END)
        
        # Get item values
        item = self.tree.item(selected_items[0])
        values = item['values']
        
        # Details column is the last one (index 9)
        if len(values) >= 10:
             details_val = values[9]
             # Formatting: Break lines for readability
             formatted_val = str(details_val).replace('; ', '\n').replace(';', '\n')
             self.details_text.insert(tk.END, formatted_val)

    def ensure_defaults(self):
        """Checks for default paths (MasterList.csv, process_file_dir) and sets them."""
        root_path = Path.cwd()
        
        # 1. Master List
        default_master = root_path / "MasterList.csv"
        if not default_master.exists():
            try:
                headers = ["Invoice No", "Amount", "Quantity", "Pallets", "PCS", "NetWeight", "GrossWeight", "CBM", "ID", "VERIFY STATE", "DIFF_PALLET", "DIFF_SQFT", "DIFF_AMOUNT", "DIFF_PCS", "DIFF_NET", "DIFF_GROSS", "DIFF_CBM"]
                with open(default_master, 'w', encoding='utf-8', newline='') as f:
                    import csv
                    writer = csv.writer(f)
                    writer.writerow(headers)
                self.status_var.set("Created default MasterList.csv")
            except Exception as e:
                print(f"Failed to create master list: {e}")
        
        if default_master.exists():
            self.master_path.set(str(default_master.resolve()))

        # 2. Process Directory
        default_process_dir = root_path / "process_file_dir"
        if not default_process_dir.exists():
            try:
                default_process_dir.mkdir(exist_ok=True)
                self.status_var.set("Created default process_file_dir/")
            except Exception as e:
                print(f"Failed to create process dir: {e}")
        
        if default_process_dir.exists():
             self.folder_path.set(str(default_process_dir.resolve()))

    def browse_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_path.set(path)

    def browse_master(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx")])
        if path:
            self.master_path.set(path)

    def start_inspection_thread(self):
        folder = self.folder_path.get()
        if not folder:
            messagebox.showerror("Error", "Please select an invoice folder.")
            return
            
        self.run_btn.config(state="disabled")
        self.status_var.set("Running inspection...")
        self.tree.delete(*self.tree.get_children())
        
        t = threading.Thread(target=self.run_inspection)
        t.start()

    def run_inspection(self):
        try:
            folder = self.folder_path.get()
            master = self.master_path.get() or None
            
            # Run Pipeline
            # Redirect stdout? For now just run
            extracted_data = run_pipeline(folder, master)
            
            # Update UI on main thread
            self.root.after(0, self.update_results, extracted_data, master)
            
        except Exception as e:
            err_msg = str(e)
            print(f"Inspection Error: {err_msg}") # Also print to console
            self.root.after(0, lambda: messagebox.showerror("Error", err_msg))
        finally:
            self.root.after(0, lambda: self.run_btn.config(state="normal"))
            self.root.after(0, lambda: self.status_var.set("Inspection Complete"))

    def update_results(self, data, master_path):
        # Clear
        self.tree.delete(*self.tree.get_children())
        
        # If we have a master path, we can read it to get status
        # Since logic writes back to master list, we should read master list again to get verification status?
        # OR run_pipeline could return enriched data.
        # Currently run_pipeline returns extracted data list.
        # But `verify_against_master` modifies the FILE, not the returned list.
        # So we should re-read the Master List to get the "VERIFY STATE".
        
        verification_map = {}
        if master_path:
             try:
                 df = pd.read_csv(master_path) if master_path.endswith('.csv') else pd.read_excel(master_path)
                 # Normalize and map ID -> State
                 # We need to find ID column and State column
                 # Assuming standard naming from my own logic
                 for _, row in df.iterrows():
                     # Find ID
                     # ID column might be NaN, which is truthy!
                     rid = None
                     
                     def is_valid_id(val):
                         if pd.isna(val): return False
                         s = str(val).strip().lower()
                         return s and s != 'nan' and s != 'none'

                     if 'ID' in row and is_valid_id(row['ID']):
                         rid = row['ID']
                     elif 'Invoice No' in row and is_valid_id(row['Invoice No']):
                         rid = row['Invoice No']
                         
                     state = row.get('VERIFY STATE')
                     if rid:
                         verification_map[str(rid).strip()] = str(state)
                 print(f"Debug: Verification Map Keys: {list(verification_map.keys())}")
             except Exception as e:
                 print(f"Debug: Error reading verification map: {e}")

        for item in data:
            inv_id = item.get('invoice_id', 'Unknown')
            
            # Determine Status
            status = "Extracted"
            tags = ()
            
            # 1. Check Previous Verification State (from Master List)
            if inv_id in verification_map:
                v_state = verification_map[inv_id]
                if v_state.lower() == 'true':
                    status = "Passed"
                    tags = ('passed',)
                else:
                    status = "Failed"
                    tags = ('failed',)
            
            # 2. Check for details from verify_against_master injection (Preferred)
            details = item.get('verification_details', '')

            # 3. Fallback: Internal Mismatch checks if not already failed?
            # actually if failed, details might be set.
            # If internal mismatch but verified true? Unlikely with new logic.
            if not details and item.get('contract_mismatch'):
                 details = item.get('mismatch_details', '')
                 if status != "Failed": # If it wasn't failed by master list (maybe not in master list?), warn
                    status = "Contract Mismatch"
                    tags = ('warning',)

            values = (
                inv_id,
                status,
                item.get('col_qty_pcs', ''),
                item.get('col_qty_sf', ''),
                item.get('col_amount', ''),
                item.get('col_pallet_count', ''),
                item.get('col_net', ''),
                item.get('col_gross', ''),
                item.get('col_cbm', ''),
                details  # New Column
            )
            self.tree.insert('', 'end', values=values, tags=tags)
            
        self.tree.tag_configure('passed', background='lightgreen')
        self.tree.tag_configure('failed', background='salmon')
        self.tree.tag_configure('warning', background='orange')
        
        # Also switch to results tab
        self.notebook.select(0)
        
        # Auto-reload editor if master list changed, but DONT switch tabs
        if master_path:
            self.refresh_master_editor_data()

    # --- Master List Editor Logic ---
    
    def refresh_master_editor_data(self):
        """Reloads master list data into the editor without switching tabs."""
        path = self.master_path.get()
        if not path: return
        
        try:
            if path.lower().endswith('.csv'):
                self.master_df = pd.read_csv(path)
            else:
                self.master_df = pd.read_excel(path)
            
            # Populate Tree
            self.editor_tree.delete(*self.editor_tree.get_children())
            
            # Update columns
            current_cols = list(self.master_df.columns)
            self.editor_tree['columns'] = current_cols
            for col in current_cols:
                self.editor_tree.heading(col, text=col)
                self.editor_tree.column(col, width=100)
                
            for index, row in self.master_df.iterrows():
                # FILTER EMPTY ROWS
                # Check key columns: Invoice No, ID, or Amount
                def is_empty(v): return pd.isna(v) or str(v).strip() == '' or str(v).lower() == 'nan'
                
                # If Invoice No AND ID are empty, skip row
                if is_empty(row.get('Invoice No')) and is_empty(row.get('ID')):
                    continue

                vals = [row[c] for c in current_cols]
                self.editor_tree.insert('', 'end', iid=index, values=vals)
                
        except Exception as e:
            print(f"Background refresh failed: {e}")

    def load_master_editor(self):
        path = self.master_path.get()
        if not path:
            return
            
        try:
            if path.lower().endswith('.csv'):
                self.master_df = pd.read_csv(path)
            else:
                self.master_df = pd.read_excel(path)
                
            # Populate Tree
            self.editor_tree.delete(*self.editor_tree.get_children())
            
            # Dynamically update columns if needed, but let's try to stick to fixed set if possible 
            # or update tree columns based on DF
            current_cols = list(self.master_df.columns)
            self.editor_tree['columns'] = current_cols
            for col in current_cols:
                self.editor_tree.heading(col, text=col)
                self.editor_tree.column(col, width=100)
                
            for index, row in self.master_df.iterrows():
                vals = [row[c] for c in current_cols]
                # Store index in iid to map back to DF
                self.editor_tree.insert('', 'end', iid=index, values=vals)
                
            self.notebook.select(1) # Switch to tab
            
        except Exception as e:
            messagebox.showerror("Error loading master list", str(e))

    def on_editor_double_click(self, event):
        selection = self.editor_tree.selection()
        if not selection:
            return
            
        item_id = selection[0]
        col = self.editor_tree.identify_column(event.x)
        # col is like '#1', '#2'
        if not col: return # Clicked outside columns?
        
        col_idx = int(col.replace('#', '')) - 1
        
        # Get column name
        col_name = self.editor_tree['columns'][col_idx]
        
        # Get current value
        current_vals = self.editor_tree.item(item_id, 'values')
        val = current_vals[col_idx]
        
        # Spawn small edit popup
        self.spawn_edit_popup(item_id, col_name, val)

    def on_editor_click(self, event):
        """Track the selected cell (row/col) for pasting."""
        region = self.editor_tree.identify("region", event.x, event.y)
        if region == "cell":
            col = self.editor_tree.identify_column(event.x)
            # col is like '#1', '#2'
            self.selected_col_idx = int(col.replace('#', '')) - 1
            
            item_id = self.editor_tree.identify_row(event.y)
            self.selected_row_id = item_id
            
            # Optional: Feedback in status bar?
            # col_name = self.editor_tree['columns'][self.selected_col_idx]
            # self.status_var.set(f"Selected: Row {item_id}, Col {col_name}")

    def spawn_edit_popup(self, item_id, col_name, current_val):
        top = tk.Toplevel(self.root)
        top.title(f"Edit {col_name}")
        
        tk.Label(top, text="Value:").pack(padx=5, pady=5)
        entry = tk.Entry(top)
        entry.pack(padx=5, pady=5)
        entry.insert(0, str(current_val))
        entry.focus()
        
        def save():
            new_val = entry.get()
            # Update Tree
            self.editor_tree.set(item_id, col_name, new_val)
            # Update DF
            # Try to convert type if needed (float/int)
            # For now string is safe
            self.master_df.at[int(item_id), col_name] = new_val
            top.destroy()
            
        tk.Button(top, text="Set", command=save).pack(pady=5)
        top.bind('<Return>', lambda e: save())

    def add_master_row(self):
        if self.master_df is None:
            return
            
        # Robust Index generation
        if not self.master_df.index.empty:
            new_row_idx = self.master_df.index.max() + 1
        else:
            new_row_idx = 0
            
        # Add empty row to DF
        self.master_df.loc[new_row_idx] = [None] * len(self.master_df.columns)
        
        # Add to Tree
        self.editor_tree.insert('', 'end', iid=new_row_idx, values=[""]*len(self.master_df.columns))

    def delete_master_row(self):
        selected_item = self.editor_tree.selection()
        if not selected_item:
            messagebox.showwarning("Selection Required", "Please select a row to delete.")
            return

        idx = int(selected_item[0]) # iid is the index
        
        # Remove from DF
        if idx in self.master_df.index:
            self.master_df.drop(index=idx, inplace=True)
            
        # Remove from Tree
        self.editor_tree.delete(selected_item)

    def save_master_list(self):
        if self.master_df is None:
            return
        path = self.master_path.get()
        try:
             if path.lower().endswith('.csv'):
                self.master_df.to_csv(path, index=False)
             else:
                self.master_df.to_excel(path, index=False)
             messagebox.showinfo("Success", "Master List saved successfully.")
        except Exception as e:
             messagebox.showerror("Error saving", str(e))

    # --- Paste Logic ---

    def parse_and_preview_paste(self):
        try:
            if self.master_df is None:
                # Try to load if path exists
                if self.master_path.get():
                    self.load_master_editor()
                
                if self.master_df is None:
                    messagebox.showerror("Error", "Please load a Master List first.")
                    return

            # 1. Get Clipboard
            try:
                clipboard_data = self.root.clipboard_get()
            except tk.TclError:
                messagebox.showwarning("Clipboard Empty", "Could not read text from clipboard.\nPlease copy cells from Excel or text and try again.")
                return
                
            if not clipboard_data:
                return
            
            # 2. Parse (assume TSV from Excel)
            rows = [line.split('\t') for line in clipboard_data.strip().split('\n')]
            if not rows:
                return

            # 3. Detect Headers & Aliases
            # Define Aliases
            aliases = {
                'net': 'NetWeight',
                'gross': 'GrossWeight',
                'pcs': 'PCS',
                'qty': 'Quantity',
                'sqft': 'Quantity',
                'amount': 'Amount',
                'invoice no': 'Invoice No',
                'id': 'ID',
                'cbm': 'CBM'
            }
            
            header_candidates = [str(cell).strip() for cell in rows[0]]
            master_cols_lower = {c.lower(): c for c in self.master_df.columns}
            
            match_count = 0
            for h in header_candidates:
                h_lower = h.lower()
                if h_lower in master_cols_lower:
                    match_count += 1
                elif h_lower in aliases:
                    # Check if alias target exists in master
                    if aliases[h_lower].lower() in master_cols_lower:
                        match_count += 1

            # If we have at least 1 match, default to Header Mode
            default_header_mode = (match_count > 0)
            
            # 4. Show Preview Dialog
            self.show_paste_preview(rows, list(self.master_df.columns), self.selected_col_idx, default_header_mode, aliases)
            
        except Exception as e:
            messagebox.showerror("Paste Error", f"Failed to read clipboard: {e}")

    def show_paste_preview(self, data_rows, master_cols, initial_start_col_idx, default_header_mode, aliases):
        top = tk.Toplevel(self.root)
        top.title("Paste Preview")
        top.geometry("1000x600")
        
        # State
        header_mode_var = tk.BooleanVar(value=default_header_mode)
        col_var = tk.StringVar()
        
        # --- Top: Controls ---
        ctrl_frame = ttk.LabelFrame(top, text="Paste Configuration", padding=10)
        ctrl_frame.pack(fill="x", padx=10, pady=5)
        
        # Header Mode Checkbox
        chk_header = ttk.Checkbutton(ctrl_frame, text="First Row contains Headers (Match by Name)", variable=header_mode_var)
        chk_header.grid(row=0, column=0, sticky="w", padx=5)
        
        # Positional Controls (Visible only if Header Mode is OFF)
        pos_frame = ttk.Frame(ctrl_frame)
        pos_frame.grid(row=0, column=1, padx=20)
        
        ttk.Label(pos_frame, text="Start Pasting At Column:").pack(side="left")
        col_combo = ttk.Combobox(pos_frame, textvariable=col_var, values=master_cols, state="readonly")
        if 0 <= initial_start_col_idx < len(master_cols):
             col_combo.current(initial_start_col_idx)
        else:
             col_combo.current(0)
        col_combo.pack(side="left", padx=5)
        
        # Status Label
        status_lbl = ttk.Label(ctrl_frame, text="", font=("Arial", 9, "italic"))
        status_lbl.grid(row=1, column=0, columnspan=2, sticky="w", pady=5)
        
        ttk.Label(ctrl_frame, text="Tip: Click column headers below to change mapping.", font=("Arial", 8)).grid(row=2, column=0, columnspan=2, sticky="w", padx=5)

        # --- Middle: Preview Grid ---
        preview_frame = ttk.Frame(top, padding=5)
        preview_frame.pack(fill="both", expand=True)

        max_paste_cols = max(len(r) for r in data_rows)
        preview_cols = [str(i) for i in range(max_paste_cols)]
        
        preview_tree = ttk.Treeview(preview_frame, columns=preview_cols, show='headings')
        
        vs = ttk.Scrollbar(preview_frame, orient="vertical", command=preview_tree.yview)
        hs = ttk.Scrollbar(preview_frame, orient="horizontal", command=preview_tree.xview)
        preview_tree.configure(yscroll=vs.set, xscroll=hs.set)
        
        preview_tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)

        # Mapping State: index -> target_col_name
        current_mapping = {}
        # Manual overrides: index -> target_col_name (or None for Ignore)
        manual_overrides = {} 

        def on_header_click(col_id):
            """Show popup menu to select target column for this source column."""
            idx = int(col_id)
            
            menu = tk.Menu(top, tearoff=0)
            
            # Option to Ignore
            menu.add_command(label="(Ignore Column)", command=lambda: set_manual_mapping(idx, None))
            menu.add_separator()
            
            # Master Columns
            for mc in master_cols:
                menu.add_command(label=mc, command=lambda c=mc: set_manual_mapping(idx, c))
                
            try:
                x, y = top.winfo_pointerxy()
                menu.tk_popup(x, y)
            finally:
                menu.grab_release()

        def set_manual_mapping(idx, target_name):
            manual_overrides[idx] = target_name
            update_ui() # Refresh

        # Bind Header Clicks
        for col_id in preview_cols:
            preview_tree.heading(col_id, command=lambda c=col_id: on_header_click(c))

        def update_ui(event=None):
            is_header_mode = header_mode_var.get()
            
            # 1. Toggle Positional Controls
            if is_header_mode:
                col_combo.state(['disabled'])
            else:
                col_combo.state(['!disabled'])

            # 2. Compute Mapping & Headers
            current_mapping.clear()
            master_cols_lower = {c.lower(): c for c in master_cols}
            
            if is_header_mode:
                # Use first row as headers
                headers = [str(c).strip() for c in data_rows[0]]
                start_data_row = 1
                
                for i in range(len(preview_cols)):
                    col_id = preview_cols[i]
                    
                    # 1. Manual Override
                    if i in manual_overrides:
                        target = manual_overrides[i]
                        if target is None:
                            # Ignored
                            preview_tree.heading(col_id, text=f"Col {i+1} -> (Ignored)")
                            # Remove from mapping
                        else:
                            current_mapping[i] = target
                            preview_tree.heading(col_id, text=f"Col {i+1} -> {target}")
                        continue

                    # 2. Auto-Detect
                    if i < len(headers):
                        h_text = headers[i]
                        h_lower = h_text.lower()
                        target = None
                        
                        if h_lower in master_cols_lower:
                            target = master_cols_lower[h_lower]
                        elif h_lower in aliases and aliases[h_lower].lower() in master_cols_lower:
                            target = aliases[h_lower]
                            # verify case
                            target = master_cols_lower[target.lower()]
                            
                        if target:
                            current_mapping[i] = target
                            preview_tree.heading(col_id, text=f"{h_text} -> {target}")
                        else:
                            preview_tree.heading(col_id, text=f"{h_text} (No Match)")
                    else:
                         preview_tree.heading(col_id, text=f"Col {i+1} (No Match)")
                        
                # Update Status
                key_cols = [c for c in current_mapping.values() if c in ("Invoice No", "ID")]
                if key_cols:
                    status_lbl.config(text=f"Mode: Merge to existing rows by '{key_cols[0]}'.", foreground="blue")
                else:
                    status_lbl.config(text="Mode: Append new rows (No ID matched).", foreground="black")
                    
            else:
                # Positional Mode
                start_data_row = 0
                start_name = col_var.get()
                try:
                     start_index = master_cols.index(start_name)
                except ValueError: 
                     start_index = 0
                
                for i, col_id in enumerate(preview_cols):
                    # Manual Override priority? Yes, why not.
                    if i in manual_overrides:
                        target = manual_overrides[i]
                        if target is None:
                             preview_tree.heading(col_id, text=f"Col {i+1} -> (Ignored)")
                        else:
                             current_mapping[i] = target
                             preview_tree.heading(col_id, text=f"Col {i+1} -> {target}")
                        continue
                        
                    target_idx = start_index + i
                    if target_idx < len(master_cols):
                        target = master_cols[target_idx]
                        current_mapping[i] = target
                        preview_tree.heading(col_id, text=f"-> {target}")
                    else:
                        preview_tree.heading(col_id, text="(Skip)")
                
                status_lbl.config(text=f"Mode: Insert at selection (Row {self.selected_row_id or 'New'}).", foreground="black")

            # 3. Refresh Data Preview
            preview_tree.delete(*preview_tree.get_children())
            
            # Show first 50 rows
            limit = 50
            for idx, row in enumerate(data_rows):
                if is_header_mode and idx == 0: continue 
                
                padded = row + ([''] * (max_paste_cols - len(row)))
                preview_tree.insert('', 'end', values=padded)
                
                if idx > limit: break

        col_combo.bind("<<ComboboxSelected>>", update_ui)
        chk_header.config(command=update_ui)
        
        # Init
        update_ui()

        
        # --- Bottom: Actions ---
        btn_frame = ttk.Frame(top, padding=10)
        btn_frame.pack(fill="x")
        
        def commit_paste():
            is_header_mode = header_mode_var.get()
            self.apply_paste(data_rows, current_mapping, is_header_mode)
            top.destroy()
            
        ttk.Button(btn_frame, text="Confirm Paste", command=commit_paste).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="Cancel", command=top.destroy).pack(side="right", padx=5)

    def apply_paste(self, rows, mapping, is_header_mode):
        if self.master_df is None: return
        
        # Identify Key Column for Merge
        # We look for "ID" or "Invoice No" in the TARGETS (values of mapping)
        key_col_target = None
        key_col_source_idx = None
        
        # Invert mapping to find source index for key
        for src_idx, target_name in mapping.items():
            if target_name in ("Invoice No", "ID"):
                key_col_target = target_name
                key_col_source_idx = src_idx
                break
        
        # Start index for loop
        start_idx = 1 if is_header_mode else 0
        
        count_updated = 0
        count_inserted = 0
        
        # Pre-calculate Index Map if merging
        # Map ID_Value -> DF_Index
        id_map = {}
        if key_col_target:
            for idx, val in self.master_df[key_col_target].items():
                if pd.notna(val):
                    id_map[str(val).strip()] = idx
        
        current_append_idx = self.master_df.index.max() + 1 if not self.master_df.empty else 0
        
        for i in range(start_idx, len(rows)):
            row_data = rows[i]
            
            target_row_idx = None
            
            # 1. Try to find existing row
            if key_col_target and len(row_data) > key_col_source_idx:
                raw_id = row_data[key_col_source_idx]
                clean_id = str(raw_id).strip()
                if clean_id in id_map:
                    target_row_idx = id_map[clean_id]
            
            # 2. If no match (or no key matching enabled), determine insert/append
            if target_row_idx is None:
                # Fallback to positional insertion if NOT in header mode?
                # Actually user asked for "fill out row of invoice id". 
                # If key not found, we should APPEND new row.
                # Use Positional logic only if NOT Is_Header_Mode and we have a selected row?
                
                if not is_header_mode and self.selected_row_id is not None:
                     # Positional insert over existing?
                     # This is tricky mixing modes.
                     # Let's verify: Positional Mode -> simply overwrite blindly from start row.
                     pass 
                
                # Simplified Logic:
                # If Header Mode -> Merge by ID. If ID not found -> Append.
                # If Positional Mode -> Overwrite blindly starting from selected_row_id.
                
                if is_header_mode:
                    target_row_idx = current_append_idx
                    current_append_idx += 1
                    count_inserted += 1
                else:
                    # Blind overwrite
                    # Need to track current offset
                    offset = i # 0-indexed relative to start of data
                    
                    start_base = int(self.selected_row_id) if (self.selected_row_id is not None) else current_append_idx
                    target_row_idx = start_base + offset
                    
                    if target_row_idx not in self.master_df.index:
                        count_inserted += 1
                        current_append_idx = max(current_append_idx, target_row_idx + 1)
            else:
                count_updated += 1

            # Ensure row exists
            if target_row_idx not in self.master_df.index:
                 self.master_df.loc[target_row_idx] = [None] * len(self.master_df.columns)
                 self.editor_tree.insert('', 'end', iid=target_row_idx, values=[""]*len(self.master_df.columns))
            
            # Write Value
            for src_idx, val in enumerate(row_data):
                if src_idx in mapping:
                    target_col = mapping[src_idx]
                    
                    # Update DF
                    self.master_df.at[target_row_idx, target_col] = val.strip()
                    # Update Tree
                    try:
                         self.editor_tree.set(target_row_idx, target_col, val.strip())
                    except: pass
                    
        messagebox.showinfo("Paste Complete", f"Merged/Updated: {count_updated}\nInserted: {count_inserted}")

if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceInspectorApp(root)
    root.mainloop()
