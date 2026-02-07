
import json
import csv
from pathlib import Path
from typing import List, Optional
from core.config import REPORTS_DIR
from core.models import ExtractedInvoice, VerificationStatus
from services.extraction_service import scan_invoice_files, excel_data_extractor
from services.master_data_service import MasterDataService

class PipelineService:
    def __init__(self, folder_path: str, master_path: Optional[str] = None):
        self.folder_path = Path(folder_path)
        self.master_path = Path(master_path) if master_path else None
        self.master_service = MasterDataService(self.master_path) if self.master_path else None
        
        # Ensure reports dir exists
        self.reports_dir = self.folder_path / "reports"
        self.reports_dir.mkdir(exist_ok=True)

    def run(self) -> List[dict]:
        """Runs the full inspection pipeline."""
        print(f"Starting Pipeline on: {self.folder_path}")
        
        # 1. Load Master Data
        master_ids = set()
        verified_ids = set()
        if self.master_service:
            print(f"Using Master List: {self.master_path.name}")
            if self.master_service.load():
                master_ids, verified_ids = self.master_service.get_known_ids()
            else:
                print("Warning: Failed to load Master List.")
        
        # 2. Key Step: Master List Fallback Search if not provided?
        # The original script searched for "Master*.xlsx" if not provided.
        # For simplicity in this Service, we assume it's passed in. 
        # The UI should handle auto-detection.

        # 3. Scan Files
        print("Scanning files...")
        scanned_files = scan_invoice_files(self.folder_path)
        
        # 4. Reconcile (Identify valid vs rejected)
        matched = []
        rejected = []
        failed_parse = []
        
        found_master_ids = set()
        
        for file_dat in scanned_files:
            ext_id = file_dat.get('extracted_id')
            if not ext_id:
                failed_parse.append(file_dat)
                continue
                
            if not master_ids:
                # If no master list, receive all parsed as matched (Extraction Mode only)
                matched.append(file_dat)
            elif ext_id in master_ids:
                matched.append(file_dat)
                found_master_ids.add(ext_id)
            else:
                rejected.append(file_dat)
                
        missing_ids = master_ids - found_master_ids
        
        print(f"Match: {len(matched)}, Reject: {len(rejected)}, Missing: {len(missing_ids)}")
        
        # 5. Generate Reports (Rejection/Missing)
        self._generate_rejection_report(rejected, failed_parse, list(missing_ids))
        
        # 6. Extract Data
        final_results = []
        for file_dat in matched:
            path = file_dat['original_path']
            print(f"Extracting: {path.name}...")
            
            data_obj = excel_data_extractor(path)
            
            # Enforce ID
            known_id = file_dat.get('extracted_id')
            if known_id: data_obj.invoice_id = known_id
            
            final_results.append(data_obj.to_dict())
            
        # 7. Write Final JSON
        output_json = self.reports_dir / "final_invoice_data.json"
        try:
            with open(output_json, 'w', encoding='utf-8') as f:
                json.dump(final_results, f, indent=4)
        except Exception as e:
            print(f"Error saving JSON: {e}")
            
        # 8. Verify against Master
        if self.master_service and final_results:
            print("Verifying against Master List...")
            self.master_service.verify_and_update(final_results)
            
        return {
            "results": final_results,
            "missing": list(missing_ids)
        }

    def _generate_rejection_report(self, rejected, failed, missing):
        # Missing
        if missing:
            p = self.reports_dir / "missing_invoices.csv"
            try:
                with open(p, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.DictWriter(f, fieldnames=['Missing Invoice ID'])
                    writer.writeheader()
                    for m in missing: writer.writerow({'Missing Invoice ID': m})
            except Exception: pass
            
        # Rejected/Failed
        rows = []
        for item in rejected:
            rows.append({'Filename': item['original_name'], 'ID': item['extracted_id'], 'Status': 'Unknown ID'})
        for item in failed:
            rows.append({'Filename': item['original_name'], 'ID': 'N/A', 'Status': 'Parse Error'})
            
        if rows:
            p = self.reports_dir / "rejection_report.csv"
            try:
                with open(p, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.DictWriter(f, fieldnames=['Filename', 'ID', 'Status'])
                    writer.writeheader()
                    writer.writerows(rows)
            except Exception: pass
