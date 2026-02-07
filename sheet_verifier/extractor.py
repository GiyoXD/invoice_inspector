import re
import json
from pathlib import Path
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from typing import Dict, Any, Optional, Set
from core.regex_utils import regex_extract_number, regex_extract

# Inspectable column types (verification-critical numeric fields)
INSPECTABLE_COLS = {
    'col_qty_sf', 'col_amount', 'col_pallet_count',
    'col_qty_pcs', 'col_net', 'col_gross', 'col_cbm'
}

class SheetExtractor:
    def __init__(self):
        # Centralized Regex Configuration
        self.patterns = {
            'total_row': re.compile(r'total(\s+of)?\s*[：:]', re.IGNORECASE),
            'invoice_id': re.compile(r'(invoice\s*no|ref\s*no|inv\s*id)', re.IGNORECASE),
            'pallet_footer': re.compile(r'(\d+)\s*pallet', re.IGNORECASE)
        }
        
        # Load header mappings from config
        self.header_mappings = self._load_header_mappings()

    def extract_sheet_data(self, sheet: Worksheet, sheet_type: str = "generic") -> Dict[str, Any]:
        """
        Extracts Amount, Quantity, Pallets, and Invoice ID from a given sheet.
        Also populates target_inspect_col with detected inspectable columns.
        """
        data = {
            'amount': None,
            'quantity': None,
            'pallets': None,
            'invoice_id': None,
            'row_found': -1,
            'header_row': -1,
            'target_inspect_col': set(),  # Track detected inspectable columns
            'extraction_status': 'ok',  # 'ok', 'partial', 'failed'
            'extraction_errors': []  # List of error messages
        }

        # 1. First Pass: Scan rows and collect header matches per row
        # We need 3+ header matches on the same row to confirm it as the header row
        data_row_idx = -1
        row_header_matches = {}  # row_idx -> list of (col_idx, col_id) tuples
        
        for row in sheet.iter_rows(max_row=150): 
            for cell in row:
                if not cell.value: continue
                val_str = str(cell.value).strip()
                
                # Check 1: Total Row (for Amt/Qty/Pallets)
                if data_row_idx == -1 and self.patterns['total_row'].search(val_str):
                    data_row_idx = cell.row
                
                # Check 2: Invoice ID (Header)
                if not data['invoice_id'] and self.patterns['invoice_id'].search(val_str):
                    found_id = self._extract_id_value(sheet, cell)
                    if found_id:
                        data['invoice_id'] = found_id
                
                # Check 3: Collect header matches per row
                if val_str.lower() in self.header_mappings:
                    col_id = self.header_mappings[val_str.lower()]
                    if cell.row not in row_header_matches:
                        row_header_matches[cell.row] = []
                    row_header_matches[cell.row].append((cell.column, col_id))
        
        # 2. Find the header row (row with MOST header matches, minimum 3)
        header_row_idx = -1
        best_match_count = 0
        best_matches = []
        
        for row_idx, matches in row_header_matches.items():
            if len(matches) >= 3 and len(matches) > best_match_count:
                best_match_count = len(matches)
                header_row_idx = row_idx
                best_matches = matches
        
        # Populate target_inspect_col from the best header row
        for col_idx, col_id in best_matches:
            if col_id in INSPECTABLE_COLS:
                data['target_inspect_col'].add(col_id)
        
        data['header_row'] = header_row_idx
        
        # STRICT: If no header row found, mark as error
        if header_row_idx == -1:
            err_msg = f"ERROR: No header row found in sheet (need 3+ matching headers). Found rows with matches: {[(r, len(m)) for r, m in row_header_matches.items()]}"
            print(err_msg)
            data['extraction_status'] = 'failed'
            data['extraction_errors'].append('Header row not found (need 3+ matches)')
        else:
            detected = list(data['target_inspect_col'])
            print(f"  Detected inspectable columns: {detected}")

        if data_row_idx == -1:
            # CRITICAL: No Total row found - cannot extract numeric values
            err_msg = "CRITICAL: No 'Total' row found in sheet. Cannot extract Amount/Quantity values."
            print(err_msg)
            data['extraction_status'] = 'failed'
            data['extraction_errors'].append(err_msg)
            return data
        
        data['row_found'] = data_row_idx
        row_cells = sheet[data_row_idx]

        
        # 2. Extract Data from that row
        # We look for:
        # - Pallet info (text with 'pallet')
        # - Numbers/Formulas (for Amt/Qty)
        
        formula_cells = []
        
        for cell in row_cells:
            val = cell.value
            
            # Pallet Check (Text) - also detect via footer pattern
            if isinstance(val, str) and 'pallet' in val.lower():
                extracted_pal = self._extract_pallet_number(val)
                if extracted_pal is not None:
                    data['pallets'] = extracted_pal
                    # Auto-add col_pallet_count even without header
                    data['target_inspect_col'].add('col_pallet_count')
            
            # Numeric/Formula Check
            is_formula = isinstance(val, str) and val.strip().startswith('=')
            is_number = isinstance(val, (int, float))
            
            if is_formula or is_number:
                formula_cells.append(cell)

        # 3. Disambiguate Amount vs Quantity
        # We look up the column headers to decide
        
        candidates = {} # col_idx -> 'amount' | 'quantity' | None
        
        for cell in formula_cells:
             header_type = self._find_header_type(sheet, cell.row, cell.column)
             candidates[cell.column] = header_type

        # Heuristic assignment
        qty_col = -1
        amt_col = -1
        
        # Pass 1: explicit headers
        for col, h_type in candidates.items():
            if h_type == 'quantity':
                qty_col = col
            elif h_type == 'amount':
                amt_col = col
                
        # Pass 2: Inference based on sheet type if still missing
        if qty_col == -1 and amt_col == -1 and len(formula_cells) >= 1:
            # If we have 2 numbers, usually larger is Amount, smaller is Qty? Unsafe.
            # If Packing List -> Primary is Quantity.
            if sheet_type == 'packing_list':
                # First numeric col is likely Qty ??
                qty_col = formula_cells[0].column
        
        # Fetch Values (using data_only assumption or re-reading logic if needed)
        # Note: Openpyxl object passes here might not be data_only=True. 
        # The caller should ideally pass a data_only workbook or we can't evaluate formulas.
        # We will assume the passed sheet has values (data_only=True was used on load).
        
        if qty_col != -1:
            data['quantity'] = sheet.cell(row=data_row_idx, column=qty_col).value
            
        if amt_col != -1:
            data['amount'] = sheet.cell(row=data_row_idx, column=amt_col).value

        # Clean values
        data['quantity'] = self._clean_number(data['quantity'])
        data['amount'] = self._clean_number(data['amount'])

        return data

    def _extract_pallet_number(self, text: str) -> Optional[float]:
        if not text: return None
        result = regex_extract_number(text, default=None)
        return result

    def _clean_number(self, val: Any) -> float:
        if val is None: return 0.0
        if isinstance(val, (int, float)): return float(val)
        try:
            return float(val)
        except:
            return 0.0

    def _find_header_type(self, sheet, row_idx, col_idx) -> Optional[str]:
        # Search upwards
        for r in range(row_idx - 1, max(0, row_idx - 50), -1):
            cell_val = sheet.cell(row=r, column=col_idx).value
            if not cell_val: continue
            
            txt = str(cell_val).lower()
            
            if 'quantity' in txt or 'qty' in txt or 'pcs' in txt or 'sqft' in txt:
                return 'quantity'
            if 'amount' in txt or 'usd' in txt or 'value' in txt:
                return 'amount'
        return None

    def _extract_id_value(self, sheet, cell) -> Optional[str]:
        """
        Extracts the ID string by searching the cell and its neighbors.
        Priority:
        1. Same Cell (stripped)
        2. Right Cell (Col + 1)
        3. Below Cell (Row + 1)
        4. Right-Below (Row + 1, Col + 1)
        """
        val = str(cell.value)
        # Remove label (Invoice No:, Ref No:, etc) using regex_extract
        clean_val = regex_extract(val, r'(?:invoice\s*no|ref\s*no|inv\s*id)\s*[:.]?\s*(.+)', group=1, default='')
        if not clean_val:
            # No match - try to use the whole value if it doesn't look like a label
            clean_val = val.strip()
            if any(x in clean_val.lower() for x in ['invoice', 'ref', 'inv']):
                clean_val = ''  # It's just a label, no ID here
        
        if len(clean_val) > 1:
            return clean_val
        
        # Define search offsets: (row_offset, col_offset)
        # We search Right, then Below, then Below-Right (optional)
        neighbors = [
            (0, 1),  # Right
            (1, 0),  # Below
            (1, 1)   # Below-Right
        ]
        
        for r_off, c_off in neighbors:
            try:
                target_cell = sheet.cell(row=cell.row + r_off, column=cell.column + c_off)
                val = target_cell.value
                if val:
                    s_val = str(val).strip()
                    if s_val:
                        return s_val
            except:
                continue
            
        return None

    def _load_header_mappings(self) -> Dict[str, str]:
        """
        Loads header text -> col_id mappings from mapping_config.json.
        Combines header_text_mappings and shipping_list_header_map.
        """
        config_path = Path("mapping_config.json")
        if not config_path.exists():
            return {}
        
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            mappings = {}
            
            # Merge source 1: header_text_mappings
            if 'header_text_mappings' in config:
                for k, v in config['header_text_mappings'].get('mappings', {}).items():
                    mappings[k.strip()] = v
            
            # Merge source 2: shipping_list_header_map
            if 'shipping_list_header_map' in config:
                for k, v in config['shipping_list_header_map'].get('mappings', {}).items():
                    mappings[k.strip()] = v
            
            return mappings
        except Exception as e:
            print(f"Warning: Could not load header mappings: {e}")
            return {}
