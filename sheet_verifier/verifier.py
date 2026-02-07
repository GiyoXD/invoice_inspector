from .extractor import SheetExtractor
import openpyxl
from pathlib import Path
from typing import Dict, Any, List

class InvoiceVerifier:
    def __init__(self):
        self.extractor = SheetExtractor()
        
    def verify_file(self, file_path: Path, master_record: Dict[str, float]) -> Dict[str, Any]:
        """
        Verifies all sheets in the file against the master record.
        master_record: { 'amount': X, 'quantity': Y, 'pallets': Z }
        """
        result = {
            'file_name': file_path.name,
            'status': 'PASS', # Optimistic default
            'details': []
        }
        
        try:
            # Open data_only=True to get formula results
            wb = openpyxl.load_workbook(file_path, data_only=True)
        except Exception as e:
            result['status'] = 'ERROR'
            result['details'].append(f"Could not open file: {e}")
            return result
        
        # Normalize sheet names
        sheet_map = {n.lower(): n for n in wb.sheetnames}
        
        # Check Sheets
        checks = [
            ('invoice', ['invoice_id', 'amount', 'quantity', 'pallets']),
            ('packing list', ['quantity', 'pallets']),
            ('contract', ['amount', 'quantity'])
        ]
        
        for sheet_key, fields_to_check in checks:
            # Find actual sheet name
            real_name = None
            for s in sheet_map:
                if sheet_key in s:
                    real_name = sheet_map[s]
                    break
            
            if not real_name:
                continue # Sheet not present, skip (Optional?) -> Maybe warn if invoice/packing missing?
                
            sheet = wb[real_name]
            extracted = self.extractor.extract_sheet_data(sheet, sheet_type=sheet_key.replace(' ', '_'))
            
            # Compare
            for field in fields_to_check:
                expected = master_record.get(field)
                actual = extracted.get(field)
                
                if field == 'invoice_id':
                     # String Comparison (Strict or Fuzzy?)
                     if expected:
                         if not actual:
                              result['status'] = 'FAIL'
                              result['details'].append(f"[{real_name}] ID Missing: Expected '{expected}', Found None (Check Extractor)")
                         elif str(expected).strip().upper() not in str(actual).strip().upper():
                              result['status'] = 'FAIL'
                              result['details'].append(f"[{real_name}] ID Mismatch: Found '{actual}', Expected '{expected}'")
                     continue

                # Numeric Comparison
                expected = expected if expected is not None else 0.0
                actual = actual if actual is not None else 0.0
                
                # Tolerance check
                try:
                    diff = abs(float(expected) - float(actual))
                    if diff > 0.01: # Strict tolerance (0.01)
                        result['status'] = 'FAIL'
                        result['details'].append(
                            f"[{real_name}] {field.capitalize()} Mismatch: Found {actual}, Expected {expected}"
                        )
                except:
                    # Conversion error
                    result['status'] = 'FAIL'
                    result['details'].append(f"[{real_name}] {field.capitalize()} Type Error: Found {actual}, Expected {expected}")
                    
        if not result['details']:
            result['details'].append("All sheets matched Master List.")
            
        return result
