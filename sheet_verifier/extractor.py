import re
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from typing import Dict, Any, Optional

class SheetExtractor:
    def __init__(self):
        # Centralized Regex Configuration
        self.patterns = {
            'total_row': re.compile(r'total(\s+of)?\s*:', re.IGNORECASE),
            'invoice_id': re.compile(r'(invoice\s*no|ref\s*no|inv\s*id)', re.IGNORECASE)
        }

    def extract_sheet_data(self, sheet: Worksheet, sheet_type: str = "generic") -> Dict[str, Any]:
        """
        Extracts Amount, Quantity, Pallets, and Invoice ID from a given sheet.
        """
        data = {
            'amount': None,
            'quantity': None,
            'pallets': None,
            'invoice_id': None,
            'row_found': -1
        }

        # 1. Scan Rows for Targets
        # Scan first 150 rows. Header (ID) is top, Total is bottom.
        data_row_idx = -1
        
        for row in sheet.iter_rows(max_row=150): 
            for cell in row:
                if not cell.value: continue
                val_str = str(cell.value)
                
                # Check 1: Total Row (for Amt/Qty/Pallets)
                if data_row_idx == -1 and self.patterns['total_row'].search(val_str):
                    data_row_idx = cell.row
                
                # Check 2: Invoice ID (Header)
                if not data['invoice_id'] and self.patterns['invoice_id'].search(val_str):
                    found_id = self._extract_id_value(sheet, cell)
                    if found_id:
                        data['invoice_id'] = found_id

        if data_row_idx != -1:
            data['row_found'] = data_row_idx
            row_cells = sheet[data_row_idx]

        
        # 2. Extract Data from that row
        # We look for:
        # - Pallet info (text with 'pallet')
        # - Numbers/Formulas (for Amt/Qty)
        
        formula_cells = []
        
        for cell in row_cells:
            val = cell.value
            
            # Pallet Check (Text)
            if isinstance(val, str) and 'pallet' in val.lower():
                extracted_pal = self._extract_pallet_number(val)
                if extracted_pal is not None:
                    data['pallets'] = extracted_pal
            
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
        match = re.search(r'(\d+(\.\d+)?)', text)
        if match:
             return float(match.group(1))
        return None

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
            if 'amount' in txt or 'total' in txt or 'usd' in txt or 'value' in txt:
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
        # Remove label (Invoice No:, Ref No:, etc)
        clean_val = re.sub(r'(invoice\s*no|ref\s*no|inv\s*id)\s*[:.]?\s*', '', val, flags=re.IGNORECASE).strip()
        
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
