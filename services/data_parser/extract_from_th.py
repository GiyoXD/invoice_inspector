import logging
import json
from decimal import Decimal, InvalidOperation

# Import the tools from your other files
from .excel_handler import ExcelHandler
from . import sheet_parser
from .config import INPUT_EXCEL_FILE, SHEET_NAME, HEADER_IDENTIFICATION_PATTERN, HEADER_SEARCH_ROW_RANGE, HEADER_SEARCH_COL_RANGE

# Use centralized logger - no basicConfig here
logger = logging.getLogger(__name__)

def run_simple_extraction_and_sum():
    """
    Extracts data from all tables, sums specified numeric columns,
    and consolidates text columns into a single JSON object.
    """
    # --- 1. EXTRACTION (using the existing modules) ---
    logging.info(f"--- Starting Extraction for {INPUT_EXCEL_FILE} ---")
    handler = ExcelHandler(INPUT_EXCEL_FILE) #
    sheet = handler.load_sheet(sheet_name=SHEET_NAME) #
    if not sheet:
        logging.error("Failed to load the sheet. Exiting.")
        return

    # Find the header row and column mapping
    header_info = sheet_parser.find_and_map_smart_headers(sheet) #
    if not header_info:
        logging.error("Could not find a valid header row. Exiting.")
        return
        
    header_row, column_mapping = header_info
    logging.info(f"Found header at row {header_row} with mapping: {column_mapping}")

    # Find all tables on the sheet
    all_header_rows = [header_row]
    additional_headers = sheet_parser.find_all_header_rows(
        sheet, HEADER_IDENTIFICATION_PATTERN, HEADER_SEARCH_ROW_RANGE, HEADER_SEARCH_COL_RANGE, start_after_row=header_row
    ) #
    all_header_rows.extend(additional_headers)

    # Extract data from all found tables
    all_tables_data = sheet_parser.extract_multiple_tables(sheet, all_header_rows, column_mapping) #
    handler.close()

    if not all_tables_data:
        logging.warning("Extraction finished, but no data was returned.")
        return

    # --- 2. AGGREGATION AND JSON CREATION (New Logic) ---
    
    # Define which fields to process
    fields_to_sum = {'net', 'gross'}
    fields_to_collect = {'po', 'item', 'description', 'cbm'}

    # Initialize containers for the final data
    final_totals = {field: Decimal('0') for field in fields_to_sum}
    final_strings = {field: set() for field in fields_to_collect}

    # Loop through each table's data that was extracted
    for table_data in all_tables_data.values():
        # Sum numeric fields
        for field in fields_to_sum:
            for value in table_data.get(field, []):
                try:
                    # Safely convert value to Decimal for accurate sum
                    if value is not None and str(value).strip():
                        final_totals[field] += Decimal(str(value))
                except (InvalidOperation, TypeError):
                    continue # Ignore values that can't be converted

        # Collect unique strings for other fields
        for field in fields_to_collect:
            for value in table_data.get(field, []):
                if value is not None and str(value).strip():
                    final_strings[field].add(str(value))

    # --- 3. FINALIZE AND OUTPUT JSON ---
    
    # Combine the summed totals and collected strings into one dictionary
    final_json_data = {
        field: float(total) for field, total in final_totals.items()
    }
    
    for field, value_set in final_strings.items():
        # Join unique values with a comma for the final "one value" string
        final_json_data[field] = ", ".join(sorted(list(value_set)))

    # Convert the dictionary to a JSON string
    output_json = json.dumps(final_json_data, indent=4)

    logging.info("--- Aggregation Complete! ---")
    print("\nFinal JSON Output:")
    print(output_json)

    # Optionally, save to a file
    with open("output.json", "w") as f:
        f.write(output_json)
    logging.info("Saved output to output.json")


if __name__ == "__main__":
    run_simple_extraction_and_sum()