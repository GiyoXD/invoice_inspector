
import sys
import os
import pandas as pd
from pathlib import Path

# Add project root to sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from services.master_data_service import MasterDataService

def run_test():
    # Setup
    base_dir = Path("tests/temp_ui_test")
    base_dir.mkdir(parents=True, exist_ok=True)
    master_path = base_dir / "MasterList.csv"
    
    # Create Dummy Master
    data = {
        'Invoice No': ['INV-001'],
        'Amount': [1000.0]
    }
    pd.DataFrame(data).to_csv(master_path, index=False)
    
    # Init Service
    service = MasterDataService(master_path)
    service.load()
    
    # Simulate Extracted Data (Partitioned)
    extracted_item = {
        'invoice_id': 'INV-001',
        'file_name': 'TestFile.xlsx',
        'sheets': {
            'Invoice': {'col_amount': 1000.0, 'detection_info': {'status': 'ok', 'header_row': 20}},
            'Contract': {'col_amount': 1000.0},
            'PackingList': {
                'col_net': 500.0,
                'col_gross': 400.0, # LOGIC ERROR
                'col_qty_sf': 500.0
            }
        },
        'verification_details': ''
    }
    
    # Run Verification
    print("Running verify_and_update...")
    service.verify_and_update([extracted_item])
    
    # Check Result
    details = extracted_item.get('verification_details', '')
    print("\n--- Result Details ---")
    print(details)
    
    # Verifications
    failures = []
    if "1. INVOICE (Row 20)" not in details: failures.append("Missing Sheet Header")
    if "Field      Current    Master     Variance" not in details: failures.append("Missing Table Header")
    if "Amount     1000.0     1000.0     -" not in details: failures.append("Missing/Incorrect Amount Row")
    
    # Logic Error Check
    if "[!] Critical Logic Error: Net Weight (500.0) > Gross Weight (400.0)" not in details: 
        failures.append("Missing Logic Check")
        
    # VERIFY "NO FAKE NUMBERS": Contract should NOT show pallets/net/gross as 0.0
    if "Pallets    0.0" in details: failures.append("Fake 0.0 Pallet found in Invoice/Contract")
    if "Net        0.0" in details: failures.append("Fake 0.0 Net found")

    if not failures:
        print("\nSUCCESS: Report format, logic checks, and NO FAKE NUMBERS verified.")
        sys.exit(0)
    else:
        print(f"\nFAIL: {failures}")
        sys.exit(1)

if __name__ == "__main__":
    run_test()
