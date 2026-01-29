from extract_invoice_data import excel_data_extractor
from pathlib import Path
import json

filename = "CT&INV&PL JF26001 DAF(1).xlsx"
folder = Path("process_file_dir")
path = folder / filename

if not path.exists():
    # Try finding it
    files = list(folder.glob("*.xlsx"))
    if files:
        path = files[0]
    else:
        print("File not found")
        exit()

print(f"Testing Extraction on: {path.name}")
data = excel_data_extractor(path)

print(json.dumps(data, indent=4))

# Check specific fields
if data['pcs'] != "N/A" and data['pcs'] > 0:
    print("SUCCESS: PCS Extracted")
else:
    print("FAIL: PCS Missing")

if data['cbm'] != "N/A" and data['cbm'] > 0:
    print("SUCCESS: CBM Extracted")
else:
    print("FAIL: CBM Missing")
