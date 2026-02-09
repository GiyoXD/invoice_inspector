import logging
import json
import argparse
from decimal import Decimal, InvalidOperation
from pathlib import Path
import sys
import os

# Add the parent directory to the Python path so we can import modules from the root
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from .excel_handler import ExcelHandler
from . import sheet_parser
from .config import SHEET_NAME, STOP_EXTRACTION_ON_EMPTY_COLUMN
from .util.converters import DataConverter

# Use centralized logger - no basicConfig here
logger = logging.getLogger(__name__)

def extract_all_tables(sheet):
    """
    Extracts all data tables from a given sheet, assuming a single header
    and multiple data blocks separated by empty rows.
    """
    # First, find the single header for the entire sheet.
    sheet_parser.HEADER_SEARCH_ROW_RANGE = (1, sheet.max_row)
    header_info = sheet_parser.find_and_map_smart_headers(sheet)
    
    if not header_info:
        logging.error("Could not find a valid header row in the sheet.")
        return {}

    header_row, column_mapping = header_info
    logging.info(f"Found header at row {header_row} with mapping: {column_mapping}")

    all_tables_data = {}
    table_count = 0
    current_row = header_row + 1

    while current_row <= sheet.max_row:
        # Extract a single table starting from the current row
        extracted_data = sheet_parser.extract_multiple_tables(sheet, [current_row -1], column_mapping)

        if extracted_data and extracted_data[1].get(STOP_EXTRACTION_ON_EMPTY_COLUMN):
            table_count += 1
            logging.info(f"Found data block {table_count} starting at row {current_row}")
            
            table_content = extracted_data[1]
            all_tables_data[table_count] = table_content
            
            # Advance current_row past the block we just extracted
            num_rows_in_table = len(table_content.get(STOP_EXTRACTION_ON_EMPTY_COLUMN, []))
            current_row += num_rows_in_table
        else:
            # If no data is found or the block is empty, move to the next row to continue scanning
            current_row += 1
            
    if not all_tables_data:
        logging.warning("Extraction finished, but no data blocks were found under the header.")

    return all_tables_data

def process_raw_data(all_tables_data):
    """Performs post-extraction processing, like parsing CBM values."""
    logging.info("--- Parsing CBM values in raw data ---")
    for table_id, table_data in all_tables_data.items():
        if 'cbm' in table_data:
            calculated_cbm_list = [sheet_parser.parse_and_calculate_cbm(cbm_str) for cbm_str in table_data['cbm']]
            all_tables_data[table_id]['cbm'] = calculated_cbm_list
    return all_tables_data

def _sum_fields(all_tables_data, fields_to_sum):
    """Helper function to sum specified numerical fields from extracted data."""
    final_totals = {field: Decimal('0') for field in fields_to_sum}
    for table_data in all_tables_data.values():
        for field in fields_to_sum:
            for value in table_data.get(field, []):
                try:
                    if value is not None and str(value).strip():
                        final_totals[field] += Decimal(str(value))
                except (InvalidOperation, TypeError):
                    continue
    return final_totals

def _collect_string_fields(all_tables_data, fields_to_collect):
    """Helper function to collect unique string values from specified fields."""
    final_strings = {field: set() for field in fields_to_collect}
    for table_data in all_tables_data.values():
        for field in fields_to_collect:
            for value in table_data.get(field, []):
                if value is not None and str(value).strip():
                    final_strings[field].add(str(value))
    return final_strings

def _count_pallets(all_tables_data, field_to_count):
    """Helper function to count pallets using the converter."""
    pallet_count = 0
    for table_data in all_tables_data.values():
        pallet_entries = table_data.get(field_to_count, [])
        pallet_count += sum(DataConverter.convert_pallet_string(entry) for entry in pallet_entries)
    return pallet_count

def aggregate_extracted_data(all_tables_data):
    """Aggregates data to create a summary view."""
    fields_to_sum = {'net', 'gross', 'cbm'}
    fields_to_collect = ['po', 'item', 'desc']
    field_to_count = 'pallet_count'

    final_totals = _sum_fields(all_tables_data, fields_to_sum)
    final_strings = _collect_string_fields(all_tables_data, fields_to_collect)
    pallet_count = _count_pallets(all_tables_data, field_to_count)

    aggregated_summary = {field: float(total) for field, total in final_totals.items()}
    for field, value_set in final_strings.items():
        aggregated_summary[field] = ", ".join(sorted(list(value_set)))
    
    aggregated_summary['pallet_count'] = pallet_count
    
    return aggregated_summary

def calculate_footer_data(all_tables_data):
    """
    Calculates footer totals from all extracted tables.
    Returns both per-table totals, grand total, and add_ons (leather_summary).
    This moves footer calculation from frontend to data parser.
    """
    from . import data_processor
    
    # Calculate per-table totals
    table_totals = {}
    for table_id, table_data in all_tables_data.items():
        table_totals[table_id] = data_processor.calculate_footer_totals(table_data)
    
    # Merge all table data for grand total calculation
    merged_data = {
        "pcs": [], "sqft": [], "net": [], "gross": [], 
        "cbm": [], "amount": [], "pallet_count": [],
        "description": [], "desc": [],  # Include both description field names for leather_summary calculation
        "po": [], "item": [], "unit_price": [], "price": []  # Include fields for aggregate_per_po_with_pallets
    }
    
    for table_data in all_tables_data.values():
        for key in merged_data:
            if key in table_data:
                merged_data[key].extend(table_data[key])
    
    # Calculate grand total
    grand_total = data_processor.calculate_footer_totals(merged_data)
    
    # Calculate leather summary (BUFFALO vs COW) across all tables
    leather_summary = data_processor.calculate_leather_summary(merged_data)
    
    # Calculate normal aggregate per PO with pallets (group by PO + price)
    normal_aggregate_per_po = data_processor.aggregate_per_po_with_pallets(merged_data)
    
    return {
        "table_totals": table_totals,
        "grand_total": grand_total,
        "add_ons": {
            "leather_summary_addon": leather_summary,  # BUFFALO vs COW summary
            "normal_aggregate_per_po_with_pallets": normal_aggregate_per_po  # PO+price aggregation
        }
    }

def write_output_json(raw_data, aggregated_summary, footer_data, output_filepath):
    """Generates and saves the final JSON output."""
    final_output = {
        "raw_data": raw_data, 
        "aggregated_summary": aggregated_summary,
        "footer_data": footer_data
    }
    output_json = json.dumps(final_output, indent=4, default=str)

    logging.info("--- Aggregation Complete! ---")
    print("\nFinal JSON Output:")
    print(output_json)

    with open(output_filepath, "w") as f:
        f.write(output_json)
    logging.info(f"Saved output to {output_filepath}")

def run_final_extraction(input_filepath, output_filepath):
    """
    Orchestrates the extraction, processing, and aggregation of data from an
    Excel file to generate a final JSON output.
    """
    logging.info(f"--- Starting Extraction for {input_filepath} ---")
    handler = ExcelHandler(input_filepath)
    sheet = handler.load_sheet(sheet_name=SHEET_NAME)
    if not sheet:
        logging.error("Failed to load the sheet. Exiting.")
        return

    # 1. Extraction
    raw_data = extract_all_tables(sheet)
    handler.close()
    if not raw_data:
        error_message = "Extraction failed: No data tables could be found in the Excel sheet. Please check the file format and content."
        logging.error(error_message)
        print(error_message, file=sys.stderr) # Also print to stderr for visibility in subprocess
        sys.exit(1) # Exit with an error code

    # 2. Post-Extraction Processing
    processed_data = process_raw_data(raw_data)

    # 3. Aggregation
    aggregated_summary = aggregate_extracted_data(processed_data)
    
    # 4. Calculate Footer Data (moved from frontend)
    footer_data = calculate_footer_data(processed_data)

    # 5. Finalize and Output JSON
    write_output_json(processed_data, aggregated_summary, footer_data, output_filepath)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract and process data from an Excel invoice file.")
    parser.add_argument("input_file", help="Path to the input Excel file (e.g., JF.xlsx).")
    parser.add_argument("-o", "--output", default="output.json", help="Path for the output JSON file (default: output.json).")
    args = parser.parse_args()
    run_final_extraction(args.input_file, args.output)