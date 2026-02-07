
import tkinter as tk
from tkinter import ttk, messagebox
from typing import List
from core.utils import open_file, delete_file

class ResultsPanel(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self._setup_ui()
        
    def _setup_ui(self):
        paned = ttk.PanedWindow(self, orient=tk.VERTICAL)
        paned.pack(fill="both", expand=True)
        
        list_frame = ttk.Frame(paned)
        paned.add(list_frame, weight=3)
        
        # List
        cols = ("ID", "Status", "Pcs", "Qty(Sqft)", "Amount", "Pallets", "Net W", "Gross W", "CBM", "Details")
        self.tree = ttk.Treeview(list_frame, columns=cols, show='headings')
        
        for col in cols:
            self.tree.heading(col, text=col)
            width = 80
            if col == "Details": width = 400
            if col == "ID": width = 120
            self.tree.column(col, width=width)
            
        vs = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        # self.tree.configure(yscroll=vs.set) # paned layout issue if side-by-side
        
        self.tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        
        # Details
        detail_frame = ttk.LabelFrame(paned, text="Details")
        paned.add(detail_frame, weight=1)
        
        # Button bar above detail text
        btn_frame = ttk.Frame(detail_frame)
        btn_frame.pack(fill="x", pady=(0, 5))
        
        self.delete_btn = ttk.Button(btn_frame, text="Delete Selected File", command=self.delete_selected)
        self.delete_btn.pack(side="left", padx=5)
        
        self.open_btn = ttk.Button(btn_frame, text="Open File", command=self.open_selected)
        self.open_btn.pack(side="left", padx=5)

        self.details_text = tk.Text(detail_frame, wrap="word", height=6)
        self.details_text.pack(fill="both", expand=True)
        
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Double-1>", self.on_double_click)  # Double-click to open file
        
        # Tags
        self.tree.tag_configure('passed', background='lightgreen')
        self.tree.tag_configure('failed', background='salmon')
        self.tree.tag_configure('warning', background='orange')

    def update_results(self, data: List[dict], master_service=None):
        self.tree.delete(*self.tree.get_children())
        self._details_map = {}  # Map tree item iid -> full details text
        self._file_paths_map = {}  # Map tree item iid -> file path

        # Create Verify Map if service available
        verify_map = {}
        if master_service and master_service.df is not None:
             # This duplicates logic slightly but is efficient enough
             # We need to know if an ID passed verification
             ids, verified = master_service.get_known_ids()
             # Wait, get_known_ids returns sets.
             # We need actual state per ID.
             # Let's rely on the 'data' passed in, assuming it's fully populated by pipeline
             pass
        
        for item in data:
            inv_id = item.get('invoice_id', 'Unknown')
            status = "Extracted"
            tags = ()
            
            # Determine Status from 'verification_details' or inferred
            details = item.get('verification_details', '')
            
            # New Model: We should check 'status' enum if available, or infer
            # If verification ran, Master List has the truth. 
            # Pipeline service should theoretically enrich the returned data with status.
            # For now, let's use the details presence as fail indicator?
            # Or if master_service loaded, check df?
            
            # Check dataframe directly if available (Source of Truth)
            if master_service and master_service.df is not None:
                # Find row
                # This is O(N*M) but N is small (gui results)
                # Optimization: Build map once?
                pass
                
            # Heuristic:
            if "Mismatch" in details or "[Fail]" in details:
                status = "Failed"
                tags = ('failed',)
            elif master_service:
                 # If we have master service but no mismatch details, assume Passed?
                 # Only if ID exists in master
                 pass
            
            # Better: read from item['status'] if we added it to model
            # core/models.py has .status field!
            
            st_val = item.get('status', 'Extracted')
            status = st_val  # Use actual status value
            
            # Apply color tags based on status
            if st_val == 'Verified':
                tags = ('passed',)
            elif st_val == 'Mismatch':
                tags = ('failed',)
            elif st_val == 'Missing from Master':
                tags = ('warning',)
            
            # Format row
            # Truncate details for table display (full text in detail panel)
            details_summary = details.replace('\n', ' ').strip()
            if len(details_summary) > 50:
                details_summary = details_summary[:47] + "..."
            
            vals = (
                inv_id, status,
                item.get('col_qty_pcs'), item.get('col_qty_sf'), item.get('col_amount'),
                item.get('col_pallet_count'), item.get('col_net'), item.get('col_gross'), item.get('col_cbm'),
                details_summary
            )
            # Store full details and file path for retrieval on select/double-click
            iid = self.tree.insert('', 'end', values=vals, tags=tags)
            self._details_map[iid] = details
            self._file_paths_map[iid] = item.get('file_path', '')


    def on_select(self, event):
        sel = self.tree.selection()
        if not sel: return
        iid = sel[0]
        # Get full details from our map (not truncated table value)
        full_details = self._details_map.get(iid, '')
        txt = str(full_details).replace('; ', '\n').replace(';', '\n')
        self.details_text.delete(1.0, tk.END)
        self.details_text.insert(tk.END, txt)

    def on_double_click(self, event):
        """Opens the source Excel file when a row is double-clicked."""
        self.open_selected()

    def open_selected(self):
        """Opens the selected file."""
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("No Selection", "Please select a row first.")
            return
        iid = sel[0]
        file_path = self._file_paths_map.get(iid, '')
        if file_path:
            open_file(file_path)
        else:
            messagebox.showwarning("No File", "No file path available for this item.")

    def delete_selected(self):
        """Deletes the selected file after confirmation."""
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("No Selection", "Please select a row first.")
            return
        
        iid = sel[0]
        file_path = self._file_paths_map.get(iid, '')
        
        if not file_path:
            messagebox.showwarning("No File", "No file path available for this item.")
            return
        
        # Get filename for display
        from pathlib import Path
        filename = Path(file_path).name
        
        # Confirm deletion
        confirm = messagebox.askyesno(
            "Confirm Delete",
            f"Are you sure you want to delete:\n\n{filename}?"
        )
        
        if confirm:
            if delete_file(file_path):
                # Remove from tree
                self.tree.delete(iid)
                # Clear details
                self.details_text.delete(1.0, tk.END)
                messagebox.showinfo("Deleted", f"{filename} has been deleted.")
            else:
                messagebox.showerror("Error", f"Failed to delete {filename}")
