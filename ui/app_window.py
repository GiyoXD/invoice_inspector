
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from pathlib import Path

from services.pipeline_service import PipelineService
from services.master_data_service import MasterDataService
from ui.components.master_editor import MasterEditor
from ui.components.results_panel import ResultsPanel

class AppWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Inspector (Refactored)")
        self.root.geometry("1200x800")
        
        # State
        self.folder_path = tk.StringVar()
        self.master_path = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")
        
        self.master_service = None
        
        self._ensure_defaults()
        self._setup_ui()

    def _ensure_defaults(self):
        cwd = Path.cwd()
        # Default Master
        def_master = cwd / "MasterList.csv"
        if not def_master.exists():
            try:
                with open(def_master, 'w', encoding='utf-8') as f:
                    f.write("Invoice No,ID,Amount,Quantity,Pallets,PCS,NetWeight,GrossWeight,CBM,VERIFY STATE,DIFF_PALLET,DIFF_SQFT,DIFF_AMOUNT,DIFF_PCS,DIFF_NET,DIFF_GROSS,DIFF_CBM\n")
            except: pass
        if def_master.exists(): self.master_path.set(str(def_master.resolve()))
        
        # Default Dir
        def_dir = cwd / "process_file_dir"
        if not def_dir.exists():
            try: def_dir.mkdir()
            except: pass
        if def_dir.exists(): self.folder_path.set(str(def_dir.resolve()))

    def _setup_ui(self):
        # 1. Controls
        c_frame = ttk.LabelFrame(self.root, text="Configuration", padding=10)
        c_frame.pack(fill="x", padx=10, pady=5)
        
        # Grid layout
        ttk.Label(c_frame, text="Invoices:").grid(row=0, column=0, sticky="e")
        ttk.Entry(c_frame, textvariable=self.folder_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(c_frame, text="Browse", command=self._browse_folder).grid(row=0, column=2)
        
        ttk.Label(c_frame, text="Master List:").grid(row=1, column=0, sticky="e")
        ttk.Entry(c_frame, textvariable=self.master_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(c_frame, text="Browse", command=self._browse_master).grid(row=1, column=2)
        ttk.Button(c_frame, text="Load/Refresh", command=self._load_master).grid(row=1, column=3)
        
        # Import Path
        self.import_path = tk.StringVar()
        ttk.Label(c_frame, text="Import File:").grid(row=2, column=0, sticky="e")
        ttk.Entry(c_frame, textvariable=self.import_path, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(c_frame, text="Import", command=self._import_file).grid(row=2, column=2)
        
        self.run_btn = ttk.Button(c_frame, text="RUN INSPECTION", command=self._start_run)
        self.run_btn.grid(row=3, column=1, pady=10, sticky="ew")

        # 2. Notebook
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Tabs
        self.results_panel = ResultsPanel(self.notebook)
        self.notebook.add(self.results_panel, text="Inspection Results")
        
        self.master_editor = MasterEditor(self.notebook)
        self.notebook.add(self.master_editor, text="Master List Editor")
        
        # 3. Status
        ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN).pack(side="bottom", fill="x")
        
        # Initial Load
        if self.master_path.get():
            self._load_master()

    def _browse_folder(self):
        p = filedialog.askdirectory()
        if p: self.folder_path.set(p)

    def _browse_master(self):
        p = filedialog.askopenfilename(filetypes=[("CSV/Excel", "*.csv *.xlsx")])
        if p:
            self.master_path.set(p)
            self._load_master()

    def _import_file(self):
        """Import a file from the pasted path into process_file_dir."""
        import shutil
        
        source_path = self.import_path.get().strip()
        if not source_path:
            messagebox.showwarning("Warning", "Paste a file path first.")
            return
        
        # Clean up path (remove quotes if pasted from explorer)
        source_path = source_path.strip('"').strip("'")
        source = Path(source_path)
        
        if not source.exists():
            messagebox.showerror("Error", f"File not found:\n{source_path}")
            return
        
        if not source.is_file():
            messagebox.showerror("Error", "Path is not a file.")
            return
        
        # Get destination
        dest_dir = Path(self.folder_path.get())
        if not dest_dir.exists():
            messagebox.showerror("Error", f"Process directory not found:\n{dest_dir}")
            return
        
        dest = dest_dir / source.name
        
        try:
            shutil.copy2(source, dest)
            self.status_var.set(f"Imported: {source.name}")
            self.import_path.set("")  # Clear input
            messagebox.showinfo("Success", f"Imported:\n{source.name}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import:\n{e}")

    def _load_master(self):
        p = self.master_path.get()
        if not p: return
        try:
            self.master_service = MasterDataService(Path(p))
            if self.master_service.load():
                self.master_editor.set_service(self.master_service)
                self.status_var.set(f"Loaded Master: {Path(p).name}")
            else:
                messagebox.showerror("Error", "Could not load master file.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _start_run(self):
        if not self.folder_path.get():
            messagebox.showwarning("Warning", "Select folder first.")
            return
            
        self.run_btn.config(state="disabled")
        self.status_var.set("Running pipeline...")
        # Clear results
        
        threading.Thread(target=self._run_thread).start()

    def _run_thread(self):
        try:
            f_path = self.folder_path.get()
            m_path = self.master_path.get()
            
            pipeline = PipelineService(f_path, m_path)
            results = pipeline.run()
            
            # Update UI
            self.root.after(0, lambda: self._on_inpect_complete(results))
            
        except Exception as e:
            print(f"Run Error: {e}")
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
            self.root.after(0, lambda: self.run_btn.config(state="normal"))

    def _on_inpect_complete(self, output):
        results = output.get('results', [])
        missing = output.get('missing', [])
        
        if missing:
            messagebox.showwarning("Missing Invoices", f"Found {len(missing)} invoices in Master List that were not found in folder.\nCheck missing_invoices.csv")

        self.results_panel.update_results(results, self.master_service)
        # Refresh master editor too as verification might have updated it
        if self.master_service:
            self.master_editor.reload_data()
            
        self.run_btn.config(state="normal")
        self.status_var.set("Inspection Complete.")
        self.notebook.select(0)
