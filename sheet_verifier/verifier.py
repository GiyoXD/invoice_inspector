from .extractor import SheetExtractor
import openpyxl
from pathlib import Path
from typing import Dict, Any, List, Set, Optional

# Define strict column requirements per sheet type
REQUIRED_COLS = {
    'invoice': {'col_qty_sf', 'col_amount', 'col_pallet_count'},
    'packing_list': {'col_qty_sf', 'col_pallet_count', 'col_qty_pcs', 'col_net', 'col_gross', 'col_cbm'},
    'contract': {'col_qty_sf', 'col_amount'}
}

class InvoiceVerifier:
    def __init__(self):
        self.extractor = SheetExtractor()
        
    def verify_file(self, file_path: Path, master_record: Dict[str, float]) -> Dict[str, Any]:
        """
        Verifies all sheets in the file against the master record using Partitioned Verification.
        
        Refactored to:
        1. Collect all values from all sheets (Partitioning).
        2. Scream if required columns are missing (Strict Existence).
        3. Verify consistency across sheets and against Master Record.
        """
        result = {
            'file_name': file_path.name,
            'status': 'PASS',
            'details': []
        }
        
        if not master_record:
             result['status'] = 'FAIL'
             result['details'].append("No Master Record found for this Invoice ID.")
             return result

        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
        except Exception as e:
            result['status'] = 'ERROR'
            result['details'].append(f"Could not open file: {e}")
            return result
        
        # Normalize sheet names
        sheet_map = {n.lower(): n for n in wb.sheetnames}
        
        # Data Collection Partitions
        # key: col_id -> val: List of {sheet: 'Invoice', value: 100.0}
        partitions: Dict[str, List[Dict[str, Any]]] = {}
        
        # Track observed sheets to ensure we checked everything
        observed_sheets = set()

        # 1. Extraction Phase
        target_sheets = {
            'invoice': 'Invoice',
            'packing list': 'Packing List',
            'contract': 'Contract'
        }

        for sheet_key, display_name in target_sheets.items():
            # Find actual sheet name
            real_name = None
            for s in sheet_map:
                if sheet_key in s:
                    real_name = sheet_map[s]
                    break
            
            if not real_name:
                continue

            observed_sheets.add(display_name)
            sheet = wb[real_name]
            
            # Extract
            extracted = self.extractor.extract_sheet_data(sheet, sheet_type=sheet_key.replace(' ', '_'))
            
            # Aggregation & Error Propagation
            if extracted['extraction_status'] == 'failed':
                result['status'] = 'FAIL'
                for err in extracted['extraction_errors']:
                    result['details'].append(f"[{display_name}] {err}")
            
            # Populate Partitions
            values = extracted.get('values', {})
            
            # Verify Existence Strictness here
            req_cols = REQUIRED_COLS.get(sheet_key.replace(' ', '_'), set())
            
            for col_id in req_cols:
                if col_id not in values or values[col_id] is None:
                    # SCREAM: Column Missing or Invalid
                    result['status'] = 'FAIL'
                    result['details'].append(f"[{display_name}] MISSING or INVALID: {col_id}")
                else:
                    # Add to partition
                    if col_id not in partitions:
                        partitions[col_id] = []
                    partitions[col_id].append({
                        'sheet': display_name,
                        'value': values[col_id]
                    })

        # 2. Value Verification Phase (Partition Consistency)
        # Check each field against Master Record
        
        for col_id, entries in partitions.items():
            expected_val = master_record.get(col_id)
            
            # Construct Partition Detail String
            # e.g. "col_amount: Invoice=100.0, Contract=100.0"
            details_parts = [f"{e['sheet']}={e['value']}" for e in entries]
            partition_str = f"{col_id}: {', '.join(details_parts)}"
            result['details'].append(partition_str)

            if expected_val is None:
                continue
                
            # Check all entries in this partition
            for entry in entries:
                actual = entry['value']
                sheet_name = entry['sheet']
                
                try:
                    diff = abs(float(expected_val) - float(actual))
                    if diff > 0.01:
                         result['status'] = 'FAIL'
                         result['details'].append(
                             f"[{sheet_name}] {col_id} Mismatch: Found {actual}, Expected {expected_val}"
                         )
                except (ValueError, TypeError):
                     result['status'] = 'FAIL'
                     result['details'].append(f"[{sheet_name}] {col_id} Invalid Value: {actual}")

        if not result['details']:
            result['details'].append("All strict checks passed.")
            
        return result
