import pandas as pd
from pathlib import Path
from typing import Dict, Optional
from core.regex_utils import regex_extract_number

def load_master_list(file_path: Path) -> Dict[str, Dict[str, float]]:
    """
    Loads the Master List and returns a dictionary of Invoice IDs to expected values.
    
    Structure:
    {
        "INV-001": {
            "amount": 1000.0,
            "quantity": 50.0,
            "pallets": 5.0
        },
        ...
    }
    """
    if not file_path.exists():
        raise FileNotFoundError(f"Master list not found: {file_path}")

    try:
        if file_path.suffix.lower() == '.csv':
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading master file: {e}")
        return {}

    # Normalize columns to lower case key map
    # We need to find: id, amount, quantity, pallets
    col_map = {}
    for c in df.columns:
        cl = str(c).lower().strip()
        if 'invoice' in cl or ('id' in cl and 'inv' in cl):
            col_map['id'] = c
        elif ('amount' in cl or 'total' in cl or 'usd' in cl) and 'quantity' not in cl:
            col_map['amount'] = c
        elif 'quantity' in cl or 'qty' in cl:
            col_map['quantity'] = c
        elif 'pallet' in cl:
            col_map['pallets'] = c

    if 'id' not in col_map:
        print(f"Error: Could not identify 'Invoice ID' column in {file_path.name}")
        print(f"Available columns: {list(df.columns)}")
        return {}
    
    # Map additional columns
    for c in df.columns:
        cl = str(c).lower().strip()
        if 'pcs' in cl or 'pieces' in cl:
            col_map['col_qty_pcs'] = c
        elif 'net' in cl and 'weight' in cl:
            col_map['col_net'] = c
        elif 'gross' in cl and 'weight' in cl:
            col_map['col_gross'] = c
        elif 'cbm' in cl:
            col_map['col_cbm'] = c

    master_data = {}
    
    for _, row in df.iterrows():
        # Get ID
        raw_id = row.get(col_map['id'])
        if pd.isna(raw_id):
            continue
        inv_id = str(raw_id).strip()

        # Helper to clean numbers
        def get_val(key):
            if key not in col_map:
                return 0.0
            val = row.get(col_map[key])
            try:
                if pd.isna(val): return 0.0
                if isinstance(val, (int, float)): return float(val)
                # Simple string cleanup
                clean_str = str(val).replace(',', '').replace('$', '').strip()
                result = regex_extract_number(clean_str, default=0.0)
                return result
            except:
                return 0.0

        master_data[inv_id] = {
            'col_amount': get_val('amount'),
            'col_qty_sf': get_val('quantity'),
            'col_pallet_count': get_val('pallets'),
            'col_qty_pcs': get_val('col_qty_pcs'),
            'col_net': get_val('col_net'),
            'col_gross': get_val('col_gross'),
            'col_cbm': get_val('col_cbm')
        }

    print(f"Loaded {len(master_data)} records from Master List.")
    return master_data
