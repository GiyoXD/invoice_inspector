import json
import os

# --- File Configuration ---
INPUT_EXCEL_FILE = "JF.xlsx" # Or specific name for this format, e.g., "JF_Data_2024.xlsx"
# Specify sheet name, or None to use the active sheet
SHEET_NAME = "Invoice"
# OUTPUT_PICKLE_FILE = "invoice_data.pkl" # Example for future use

# --- Sheet Parsing Configuration ---
# Row/Column range to search for the header
# Adjusted to a more realistic range to improve performance and avoid matching stray text.
HEADER_SEARCH_ROW_RANGE = (1, 50)
HEADER_SEARCH_COL_RANGE = (1, 30) # Increased range slightly, adjust if many columns
# A pattern (string or regex) to identify a cell within the header row
# This pattern helps find *any* header row, the mapping below specifies exact matches
HEADER_IDENTIFICATION_PATTERN = r"^(批次号|订单号|物料代码|总张数|净重|毛重|po|item|pcs|net|gross|TTX编号)$" # Broadened slightly


EXPECTED_HEADER_DATA_TYPES = {
    'col_po': ['string', 'numeric'],
    'col_item': ['string'], # Production Order is always a string that matches the pattern
    'col_desc': ['string'],
    'col_qty_pcs': ['numeric'],
    'col_net': ['numeric'],
    'col_gross': ['numeric'],
    'col_unit_price': ['numeric'],
    'col_amount': ['numeric'],
    'col_qty_sf': ['numeric'],
    'col_cbm': ['numeric', 'string'], # CBM can be a number or a string like '1*2*3'
    'col_dc': ['string'],
    'col_batch_no': ['string'],
    'col_line_no': ['string'],
    'col_direction': ['string'],
    'col_production_date': ['string'],
    'col_production_order_no': ['string'],
    'col_reference_code': ['string'],
    'col_level': ['string'],
    'col_pallet_count': ['numeric', 'string'],
    'col_manual_no': ['string'],
    'col_remarks': ['string'],
    'col_inv_no': ['string'],
    'col_inv_date': ['string', 'numeric', 'date'], # Invoice date can be a string or numeric date
    'col_inv_ref': ['string'],
}

# --- Column Mapping Configuration ---
# Canonical Name -> List containing header variations (case-insensitive match)
TARGET_HEADERS_MAP = {
    # priority first
    "col_production_order_no": ["production order number", "生产单号", "po", "入库单号", "PO", "PO NO.", "订单号", "TTX编号"], # Primary English: 'production order number', Primary Chinese: '生产单号'
    "col_unit_price": ["unit price", "单价", "price", "unit", "USD", "usd", "单价USD", "价格", "单价 USD"],          # Primary English: 'unit price', Primary Chinese: '单价'
    # --- Core Logic Canonical Names ---
    "col_po": ["PO NO.", "po", "PO", "Po", "订单号", "order number", "order no", "Po Nb", "尺数", "PO NB", "Po Nb", "客户订单号", "订单号", "P.O Nº", "P.O NO", "PO Nº"],                 # Primary English: 'po', Primary Chinese: '订单号'
    "col_item": ["物料代码","item no", "ITEM NO.",  'item', "Item No", "ITEM NO", "Item No", "客户品名", "物料编码", "产品编号", "ITEM Nº", "Item Nº"],        # Primary English: 'item no', Primary Chinese: '物料代码'
    "col_qty_pcs": ["pcs", "总张数", "张数", "PCS", "Quantity"],                # Primary English: 'pcs', Primary Chinese: '总张数'
    "col_net": ["NW", "net weight", "净重kg", "净重", "net", "净重KG", "N.W (kgs)", "N.W", "Net Weight"],          # Primary English: 'net weight', Primary Chinese: '净重'
    "col_gross": ["GW", "gross weight", "毛重", "gross", "Gross", "gross weight", "Gross Weight", "毛重量KG", "重量KG", "重量", "毛重", "毛重KG", "G.W(kgs)", "G.W", "Gross Weight"],       # Primary English: 'gross weight', Primary Chinese: '毛重'
    "col_qty_sf": ["sqft", "出货数量 (sf)", "尺数", "SF", "出货数量(sf)", "出货数量(SF)", "出货数量 SF", "尺码", "出货数量（SF）"],      # Primary English: 'sqft', Primary Chinese: '出货数量 (sf)' (Assuming this specific text)
    "col_amount": ["金额 USD","金额USD", "金额", "USD","amount", "总价", "usd", "Amount", "Total Amount", "total", "total amount"],            # Primary English: 'amount', Primary Chinese: '金额' # Ensure this is present and mapped

    # --- Less Certain Canonical Names ---
    "col_date_recipt": ["入库时间", "入库日期", "date receipt", "Date Receipt", "date receipt", "Date Receipt", "date receipt"],
    "col_cbm": ["cbm", "材积", "CBM","remarks", "备注", "Remark", 'remark', '低', "REMARKS", "REMARK"],          # Primary English: 'remarks', Primary Chinese: '备注'
    "col_desc": ["description","产品名称", "品名规格", "描述", "desc", "DESCRIPTION"],      # Primary English: 'description', Primary Chinese: '品名规格'
    "col_inv_no": ["invoice no", "发票号码", "INV NO", "INV NO", "inv no", "INV NO", "inv no", "INVOICE NO"],    # Primary English: 'invoice no', Primary Chinese: '发票号码'
    "col_inv_date": ["invoice date", "发票日期", "INV DATE", "INV DATE", "inv date", "INV DATE", "inv date", "INVOICE DATE", "invoice date"], # Primary English: 'invoice date', Primary Chinese: '发票日期'
    "col_inv_ref": ["ref", "invoice ref", "ref no", "REF NO", "REF NO", "ref no", "inv ref", "INV REF", "INVOICE REF"],

    "col_remarks": ["cbm", "材积", "CBM", "remarks", "备注", "Remark", 'remark', '低', "REMARKS", "REMARK"],          # Primary English: 'remarks', Primary Chinese: '备注'
    # --- Other Found Headers ---
    "col_dc": ["批次号", "DC", "dc"],
    "col_batch_no": ["batch number", "批次号"],  # Primary English: 'batch number', Primary Chinese: '批次号'
    "col_line_no": ["line no", "行号"],           # Primary English: 'line no', Primary Chinese: '行号'
    "col_direction": ["direction", "内向"],      # Primary English: 'direction', Primary Chinese: '内向' (Meaning still unclear)
    "col_production_date": ["production date", "生产日期"], # Primary English: 'production date', Primary Chinese: '生产日期'
    "col_reference_code": ["reference code", "ttx编号", "生产名称"], # Primary English: 'reference code', Primary Chinese: 'ttx编号' (Verify 'ttx编号')
    "col_level": ["grade", "等级"],              # Primary English: 'grade', Primary Chinese: '等级'
    "col_pallet_count": ["pallet count", "拖数", "PALLET", "件数", "PALLET COUNT", "pallet count", "托数"],# Primary English: 'pallet count', Primary Chinese: '拖数'
    "col_manual_no": ["manual number", "手册号"], # Primary English: 'manual number', Primary Chinese: '手册号'
    # 'amount' is already defined above

    # Add any other essential headers here following the variations list format
}

# --- Header Validation Configuration ---
EXPECTED_HEADER_PATTERNS = {
    'col_production_order_no': [
        r'^(25|26|27)\d{5}-\d{2}$',
    ],
    'col_cbm': [
        r'^\d+(\.\d+)?\*\d+(\.\d+)?\*\d+(\.\d+)?$'
    ],
    # This pattern is now a fallback for the specific value check below
    'col_pallet_count': [
        r'^1$'
    ],
    'col_remarks': [r'^\D+$'],  # Non-numeric characters only
}

EXPECTED_HEADER_VALUES = {
    # If a column header maps to 'col_pallet_count', the data value below it MUST be 1.
    # Otherwise, the column will be ignored for the 'col_pallet_count' mapping.
    'col_pallet_count': [1]
}

HEADERLESS_COLUMN_PATTERNS = {
    # If an empty header cell has data below it that looks like "number*number*number",
    # map it as the 'col_cbm' column.
    'col_cbm': [
        r'^\d+(\.\d+)?\*\d+(\.\d+)?\*\d+(\.\d+)?$',
    ],


    # You can add other rules here in the future, for example:
    # 'serial_no': [r'^[A-Z]{3}-\d{5}$']
}



# --- Data Extraction Configuration ---
# Choose a column likely to be empty *only* when the data rows truly end.
# 'item' is often a good candidate if item codes are always present for data rows.
STOP_EXTRACTION_ON_EMPTY_COLUMN = 'col_item'
# Safety limit for the number of data rows to read below the header within a table
MAX_DATA_ROWS_TO_SCAN = 1000

# --- Data Processing Configuration ---
# List of canonical header names for columns where values should be distributed
# CBM processing/distribution depends on the 'col_cbm' mapping above and if the column contains L*W*H strings
COLUMNS_TO_DISTRIBUTE = ["col_net", "col_gross", "col_cbm"] # Include 'col_cbm' if you want to distribute calculated CBM values

# The canonical header name of the column used for proportional distribution
DISTRIBUTION_BASIS_COLUMN = "col_qty_pcs"

# --- Aggregation Strategy Configuration ---
# List or Tuple of *workbook filename* prefixes (case-sensitive) that trigger CUSTOM aggregation.
# Custom aggregation sums 'col_qty_sf' and 'col_amount' based ONLY on 'col_po' and 'col_item'.
# Standard aggregation sums 'col_qty_sf' based on 'col_po', 'col_item', and 'col_unit_price'.
# Example: If INPUT_EXCEL_FILE is "JF_Report_Q1.xlsx", it will match "JF".
CUSTOM_AGGREGATION_WORKBOOK_PREFIXES = () # Renamed Variable

# --- Dynamic Loading from JSON Config ---

def load_and_update_mappings():
    """
    Loads header mappings from the JSON config file and updates the
    TARGET_HEADERS_MAP dictionary. This makes the configuration dynamic.
    """
    try:
        from core.system_config import sys_config
        json_path = sys_config.mapping_config_path
        
        if not json_path.exists():
            print(f"Warning: Mapping config file not found at {json_path}. Using default TARGET_HEADERS_MAP.")
            return

        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            
        # Use header_text_mappings as per new config structure
        mappings = data.get('header_text_mappings', {}).get('mappings', {})
        
        for header, canonical in mappings.items():
            # Handle legacy canonical names in JSON by mapping them to col_ names if needed
            target_canonical = canonical
            if not target_canonical.startswith('col_'):
                 # Simple heuristic mapping - user should update JSON ideally
                 if target_canonical == 'po': target_canonical = 'col_po'
                 elif target_canonical == 'item': target_canonical = 'col_item'
                 elif target_canonical == 'unit': target_canonical = 'col_unit_price'
                 elif target_canonical == 'pcs': target_canonical = 'col_qty_pcs'
                 elif target_canonical == 'sqft': target_canonical = 'col_qty_sf'
                 elif target_canonical == 'amount': target_canonical = 'col_amount'
                 elif target_canonical == 'net': target_canonical = 'col_net'
                 elif target_canonical == 'gross': target_canonical = 'col_gross'
                 elif target_canonical == 'cbm': target_canonical = 'col_cbm'
                 elif target_canonical == 'desc': target_canonical = 'col_desc'
                 elif target_canonical == 'pallet_count': target_canonical = 'col_pallet_count'
                 # Add others as needed or rely on user to update JSON
            
            if target_canonical in TARGET_HEADERS_MAP:
                if header not in TARGET_HEADERS_MAP[target_canonical]:
                    TARGET_HEADERS_MAP[target_canonical].append(header)
            else:
                # If the canonical name (e.g., 'new_header_type') doesn't exist in the map, create it
                TARGET_HEADERS_MAP[target_canonical] = [header]

    except json.JSONDecodeError:
        print("Warning: Could not decode mapping_config.json. Check for syntax errors. Using default TARGET_HEADERS_MAP.")
    except Exception as e:
        print(f"An unexpected error occurred while loading mapping_config.json: {e}")

# Immediately call the function to update the map when this module is imported
load_and_update_mappings()