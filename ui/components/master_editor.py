
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from services.master_data_service import MasterDataService

class MasterEditor(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.service = None
        
        # UI Setup
        self._setup_ui()
        
        # State
        self.selected_col_idx = 0
        self.selected_row_id = None

    def _setup_ui(self):
        # Controls
        ctrl_frame = ttk.Frame(self)
        ctrl_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Button(ctrl_frame, text="Save Changes", command=self.save_data).pack(side="left", padx=5)
        ttk.Button(ctrl_frame, text="Reload File", command=self.reload_data).pack(side="left", padx=5)
        ttk.Button(ctrl_frame, text="Add Row", command=self.add_row).pack(side="left", padx=5)
        ttk.Button(ctrl_frame, text="Delete Row", command=self.delete_row).pack(side="left", padx=5)
        ttk.Button(ctrl_frame, text="Paste Data", command=self.paste_data).pack(side="left", padx=5)
        
        # Treeview
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.tree = ttk.Treeview(tree_frame, show='headings')
        vs = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=vs.set, xscroll=hs.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Events
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<ButtonRelease-1>", self.on_click)
        self.tree.bind("<Control-v>", lambda e: self.paste_data())

    def set_service(self, service: MasterDataService):
        self.service = service
        self.reload_data()

    def reload_data(self):
        if not self.service: return
        try:
            if self.service.load():
                self._populate_tree()
            else:
                messagebox.showerror("Error", "Failed to load master file.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _populate_tree(self):
        self.tree.delete(*self.tree.get_children())
        if self.service.df is None: return
        
        cols = list(self.service.df.columns)
        self.tree['columns'] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=100)
            
        for idx, row in self.service.df.iterrows():
            vals = [row[c] for c in cols]
            # Use index as iid
            self.tree.insert('', 'end', iid=idx, values=vals)

    def save_data(self):
        if not self.service: return
        try:
            if self.service.master_path.suffix == '.csv':
                self.service.df.to_csv(self.service.master_path, index=False)
            else:
                self.service.df.to_excel(self.service.master_path, index=False)
            messagebox.showinfo("Success", "Master List Saved.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def add_row(self):
        if not self.service or self.service.df is None: return
        new_idx = self.service.df.index.max() + 1 if not self.service.df.empty else 0
        self.service.df.loc[new_idx] = [None] * len(self.service.df.columns)
        self._populate_tree() # Refresh simplest

    def delete_row(self):
        sel = self.tree.selection()
        if not sel: return
        idx = int(sel[0])
        if idx in self.service.df.index:
            self.service.df.drop(index=idx, inplace=True)
            self.tree.delete(sel)

    def on_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            col = self.tree.identify_column(event.x)
            self.selected_col_idx = int(col.replace('#', '')) - 1
            self.selected_row_id = self.tree.identify_row(event.y)

    def on_double_click(self, event):
        sel = self.tree.selection()
        if not sel: return
        idx = int(sel[0])
        col = self.tree.identify_column(event.x)
        if not col: return
        col_idx = int(col.replace('#', '')) - 1
        col_name = self.tree['columns'][col_idx]
        
        current_val = self.service.df.at[idx, col_name]
        
        new_val = simpledialog.askstring("Edit", f"Enter new value for {col_name}:", initialvalue=str(current_val))
        if new_val is not None:
             self.service.df.at[idx, col_name] = new_val
             self.tree.set(sel[0], col_name, new_val)

    def paste_data(self):
        if not self.service: return 
        try:
            txt = self.clipboard_get()
        except:
             messagebox.showerror("Error", "Empty Clipboard")
             return
             
        rows, mapping, is_header = self.service.parse_paste_data(txt)
        if not rows: return
        
        self._show_paste_preview(rows, mapping, is_header)

    def _show_paste_preview(self, rows, initial_mapping, is_header_mode):
        top = tk.Toplevel(self)
        top.title("Paste Preview - Select column mappings")
        top.geometry("1100x650")
        
        # State
        header_var = tk.BooleanVar(value=is_header_mode)
        current_mapping = initial_mapping.copy()
        
        # Available columns for dropdown (exclude system columns)
        master_cols = list(self.service.df.columns) if self.service.df is not None else []
        # Filter out DIFF_* and VERIFY STATE - those are system-calculated
        master_cols = [c for c in master_cols if not c.startswith('DIFF_') and c != 'VERIFY STATE']
        std_cols = ['invoice_id', 'col_qty_sf', 'col_amount', 'col_pallet_count', 
                    'col_qty_pcs', 'col_net', 'col_gross', 'col_cbm']
        all_cols = ['(skip)'] + sorted(set(master_cols + std_cols))
        
        # Top controls
        c_frame = ttk.LabelFrame(top, text="Options")
        c_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Checkbutton(c_frame, text="First row is header", variable=header_var, 
                        command=lambda: update_preview()).pack(side="left", padx=10)
        
        # Column mapping frame (horizontal scrollable)
        map_frame = ttk.LabelFrame(top, text="Column Mappings (click to change)")
        map_frame.pack(fill="x", padx=5, pady=5)
        
        # Canvas for horizontal scroll
        canvas = tk.Canvas(map_frame, height=50)
        h_scroll = ttk.Scrollbar(map_frame, orient="horizontal", command=canvas.xview)
        combo_frame = ttk.Frame(canvas)
        
        canvas.pack(side="top", fill="x", expand=True)
        h_scroll.pack(side="bottom", fill="x")
        canvas.configure(xscrollcommand=h_scroll.set)
        
        canvas_window = canvas.create_window((0, 0), window=combo_frame, anchor="nw")
        
        # Store combobox references
        combo_vars = {}  # col_idx -> StringVar
        
        max_c = max(len(r) for r in rows) if rows else 0
        
        for i in range(max_c):
            frame = ttk.Frame(combo_frame)
            frame.pack(side="left", padx=2, pady=2)
            
            # Label (show original header if exists)
            if is_header_mode and i < len(rows[0]):
                lbl_text = f"Col {i}: {rows[0][i][:15]}" if len(str(rows[0][i])) > 15 else f"Col {i}: {rows[0][i]}"
            else:
                lbl_text = f"Column {i}"
            ttk.Label(frame, text=lbl_text, width=18).pack()
            
            # Combobox
            var = tk.StringVar()
            combo_vars[i] = var
            
            # Set initial value
            if i in current_mapping:
                var.set(current_mapping[i])
            else:
                var.set("(skip)")
            
            cb = ttk.Combobox(frame, textvariable=var, values=all_cols, width=15, state="readonly")
            cb.pack()
            
            # Update mapping when changed
            def on_combo_change(event, idx=i, v=var):
                val = v.get()
                if val == "(skip)":
                    if idx in current_mapping:
                        del current_mapping[idx]
                else:
                    current_mapping[idx] = val
            
            cb.bind("<<ComboboxSelected>>", on_combo_change)
        
        # Update canvas scroll region
        combo_frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
        
        # Preview Tree
        tree_frame = ttk.Frame(top)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        p_tree = ttk.Treeview(tree_frame, show='headings')
        vs = ttk.Scrollbar(tree_frame, orient="vertical", command=p_tree.yview)
        hs = ttk.Scrollbar(tree_frame, orient="horizontal", command=p_tree.xview)
        p_tree.configure(yscroll=vs.set, xscroll=hs.set)
        
        p_tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        def update_preview():
            p_tree.delete(*p_tree.get_children())
            
            # Setup cols
            p_cols = [str(i) for i in range(max_c)]
            p_tree['columns'] = p_cols
            
            for i in range(max_c):
                # Header text
                if i in current_mapping:
                    p_tree.heading(str(i), text=f"→ {current_mapping[i]}")
                else:
                    p_tree.heading(str(i), text=f"(skipped)")
                p_tree.column(str(i), width=100)
            
            start = 1 if header_var.get() else 0
            for r in rows[start:]:
                p_tree.insert('', 'end', values=r)
        
        update_preview()
        
        # Actions
        b_frame = ttk.Frame(top)
        b_frame.pack(fill="x", padx=5, pady=10)
        
        def commit():
            self.service.apply_paste(rows, current_mapping, header_var.get())
            self._populate_tree()
            top.destroy()
        
        ttk.Button(b_frame, text="Cancel", command=top.destroy).pack(side="right", padx=5)
        ttk.Button(b_frame, text="Commit Merge", command=commit).pack(side="right", padx=5)

