import argparse
import sys
from pathlib import Path

# Add project root to sys path if needed, though relative imports might tricky if run as script.
# Standard way: user runs `python -m sheet_verifier.main` from root.
# Or we fix path here.
sys.path.append(str(Path(__file__).parent.parent))

from sheet_verifier.master_loader import load_master_list
from sheet_verifier.verifier import InvoiceVerifier
from sheet_verifier.reporter import generate_report
from sheet_verifier.extractor import SheetExtractor # Used implicitly by verifier, but good to check import

def parse_filename_for_id(filename: str, master_ids: list) -> str:
    # 1. Check exact match
    for mid in master_ids:
        if mid in filename:
            return mid
    
    # 2. Simple regex fallback (optional)
    import re
    match = re.search(r'([A-Z]+[-_]?\d+)', filename)
    if match:
        return match.group(1)
        
    return None

def main():
    parser = argparse.ArgumentParser(description="Multi-Sheet Invoice Verifier")
    parser.add_argument("--folder", type=Path, default=Path.cwd(), help="Folder containing invoices")
    parser.add_argument("--master", type=Path, required=True, help="Path to Master List (Excel/CSV)")
    
    args = parser.parse_args()
    
    print("--- Multi-Sheet Verification Tool ---")
    
    # 1. Load Master List
    print(f"Loading Master List: {args.master}")
    master_data = load_master_list(args.master)
    if not master_data:
        print("Failed to load master list. Exiting.")
        return

    # 2. Init Verifier
    verifier = InvoiceVerifier()
    
    # 3. Scan Files
    folder = args.folder
    if not folder.exists():
        print(f"Folder not found: {folder}")
        return
        
    files = list(folder.glob("*.xlsx")) + list(folder.glob("*.xls"))
    print(f"Found {len(files)} Excel files in {folder}")
    
    results = []
    
    for f in files:
        if f.name.startswith("~"): continue # Skip temp files
        if f.resolve() == args.master.resolve(): continue # Skip master itself
        
        print(f"Verifying: {f.name}...")
        
        # Identify ID
        inv_id = parse_filename_for_id(f.name, list(master_data.keys()))
        
        if not inv_id:
            results.append({
                'file_name': f.name,
                'status': 'SKIPPED',
                'details': ['Could not identify Invoice ID in filename']
            })
            continue
            
        if inv_id not in master_data:
             results.append({
                'file_name': f.name,
                'status': 'REJECTED',
                'details': [f"ID {inv_id} not found in Master List"]
            })
             continue
             
        # Verify
        master_record = master_data[inv_id]
        res = verifier.verify_file(f, master_record)
        results.append(res)
        
    # 4. Report
    generate_report(results, folder)
    print("Done.")

if __name__ == "__main__":
    main()
