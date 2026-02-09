# --- START OF FULL FILE: main.py ---
# --- Fixed datetime JSON serialization ---

import logging
import pprint
import re
import decimal
import os
import json # Added for JSON output
import datetime # <<< ADDED IMPORT for datetime handling
import argparse # <<< ADDED IMPORT for argument parsing
from pathlib import Path # <<< ADDED IMPORT for pathlib
from typing import Dict, List, Any, Optional, Tuple, Union
import time # Added for timing operations

# Import from our refactored modules
try:
    from . import config as cfg # Keep config for fallback and other settings
except ImportError:
    logging.error("Failed to import config.py. Please ensure it exists and is configured.")
    # Define dummy cfg values if needed for script to load, but it will likely fail later
    class DummyConfig:
        INPUT_EXCEL_FILE = "fallback_excel.xlsx" # Example placeholder
        SHEET_NAME = "Sheet1"
        HEADER_IDENTIFICATION_PATTERN = r"PO#" # Example
        HEADER_SEARCH_ROW_RANGE = (1, 20) # Example
        HEADER_SEARCH_COL_RANGE = (1, 30) # Example
        COLUMNS_TO_DISTRIBUTE = [] # Example
        DISTRIBUTION_BASIS_COLUMN = "SQFT" # Example
        CUSTOM_AGGREGATION_WORKBOOK_PREFIXES = ["CUST"] # eeExample
    cfg = DummyConfig()
    logging.warning("Using dummy config values due to import failure.")


from .excel_handler import ExcelHandler
from . import sheet_parser
from . import data_processor # Includes all processing functions

# Use centralized logger - no basicConfig here
# Logging is configured by core.logger_config.setup_logging() at app startup
logger = logging.getLogger(__name__)

# --- Constants for Log Truncation ---
MAX_LOG_DICT_LEN = 3000 # Max length for printing large dicts in logs (for DEBUG)

# --- Constants for DAF Compounding Formatting ---
DAF_CHUNK_SIZE = 2  # How many items per group (e.g., PO1\\PO2)
DAF_INTRA_CHUNK_SEPARATOR = "/"  # Separator within a group (e.g., DOUBLE BACKSLASH)
DAF_INTER_CHUNK_SEPARATOR = "\n"  # Separator between groups (e.g., newline)

# Type alias for the two possible initial aggregation structures
# UPDATED Type Alias to reflect new key structures
InitialAggregationResults = Union[
    Dict[Tuple[Any, Any, Optional[decimal.Decimal], Optional[str]], Dict[str, decimal.Decimal]], # Standard Result (PO, Item, Price, Desc)
    Dict[Tuple[Any, Any, Optional[str], None], Dict[str, decimal.Decimal]]                             # Custom Result (PO, Item, Desc, None) - UPDATED
]
# Type alias for the DAF compounding result structure
DAFCompoundingResult = Dict[str, Union[str, decimal.Decimal]]

# Type alias for the final DAF result (ALWAYS a list of dicts now)
FinalDAFResultType = List[DAFCompoundingResult]


# *** DAF Compounding Function with Chunking ***
def perform_DAF_compounding(
    initial_results: InitialAggregationResults, # Type hint updated
    aggregation_mode: str # 'standard' or 'custom' -> Needed to parse input keys correctly
) -> Optional[FinalDAFResultType]: # <<< Return type is always Dict[str, ...]
    """
    Performs DAF Compounding from standard or custom aggregation results.
    - If description data IS present: Performs BUFFALO split (Groups "1" & "2").
      Uses DAF_CHUNK_SIZE=2 and DAF_INTRA_CHUNK_SEPARATOR='\\'.
    - If description data IS NOT present: Performs PO Count split (Groups "1", "2", ...).
      Uses NO_DESC_SPLIT_CHUNK_SIZE=8 and DAF_INTRA_CHUNK_SEPARATOR='\\'.
      Calculates chunk-specific totals.

    Args:
        initial_results: The dictionary from EITHER standard OR custom aggregation.
        aggregation_mode: String ('standard' or 'custom') indicating key structure.
    Returns:
        - A dictionary keyed by group/chunk index ("1", "2", ...).
        - Default structure (empty groups "1", "2") if input is empty.
        - None on critical internal errors.
    """
    prefix = "[perform_DAF_compounding]"
    logging.info(f"{prefix} Starting DAF Compounding. Checking for descriptions to determine split type.")

    # Helper function for creating a default empty group result
    def default_group_result() -> DAFCompoundingResult:
        return {
            'col_po': '',
            'col_item': '',
            'col_desc': '',
            'col_qty_sf': decimal.Decimal(0),
            'col_amount': decimal.Decimal(0)
        }

    # Handle empty input consistently -> returns default BUFFALO split dict
    if not initial_results:
        logging.warning(f"{prefix} Input aggregation results map is empty. Returning default empty DAF groups.")
        return [
            default_group_result(), # Buffalo group
            default_group_result()  # Non-Buffalo group
        ]

    # --- Check if any description data exists ---
    any_description_present = False
    for key in initial_results.keys():
        desc_key_val = None
        try:
            if aggregation_mode == 'standard' and len(key) == 4:
                desc_key_val = key[3]
            elif aggregation_mode == 'custom' and len(key) == 4:
                desc_key_val = key[3]
            elif len(key) >= 4: desc_key_val = key[3]
            if desc_key_val is not None and str(desc_key_val).strip():
                any_description_present = True
                logging.debug(f"{prefix} Found description data. Will perform BUFFALO split.")
                break
        except (IndexError, TypeError): continue

    # Reusable helper function for formatting chunks
    def format_chunks(items: List[str], chunk_size: int, intra_sep: str, inter_sep: str) -> str:
        if not items:
            return ""
        processed_chunks = []
        for i in range(0, len(items), chunk_size):
            chunk = [str(item) for item in items[i:i + chunk_size]]
            joined_chunk = intra_sep.join(chunk)
            processed_chunks.append(joined_chunk)
        return inter_sep.join(processed_chunks)

    # --- Decide Execution Path --- #

    if any_description_present:
        # --- Path 1: Descriptions ARE present -> BUFFALO Split Aggregation (Chunk Size 2) ---
        logging.info(f"{prefix} Performing BUFFALO split aggregation (Chunk Size: {DAF_CHUNK_SIZE}).")
        # Initialize accumulators for BUFFALO group ("1")
        buffalo_pos = set()
        buffalo_items = set()
        buffalo_descriptions = set()
        buffalo_sqft = decimal.Decimal(0)
        buffalo_amount = decimal.Decimal(0)
        # Initialize accumulators for NON-BUFFALO group ("2")
        non_buffalo_pos = set()
        non_buffalo_items = set()
        non_buffalo_descriptions = set()
        non_buffalo_sqft = decimal.Decimal(0)
        non_buffalo_amount = decimal.Decimal(0)

        logging.debug(f"{prefix} Processing {len(initial_results)} entries for BUFFALO split.")
        for key, sums_dict in initial_results.items():
             po_key_val, item_key_val, desc_key_val = None, None, None
             try: # Extract PO, Item, Desc
                 if aggregation_mode == 'standard' and len(key) == 4:
                     po_key_val, item_key_val, _, desc_key_val = key
                 elif aggregation_mode == 'custom' and len(key) == 4:
                     po_key_val, item_key_val, _, desc_key_val = key
                 else:
                     if len(key) != 4: logging.warning(f"{prefix} Unexpected key length ({len(key)}) for key {key} in BUFFALO split mode. Trying heuristic.")
                     if len(key) >= 2: po_key_val, item_key_val = key[0], key[1]
                     if len(key) >= 4: desc_key_val = key[3]
                     if po_key_val is None or item_key_val is None:
                         logging.warning(f"{prefix} Cannot extract PO/Item/Desc reliably from key {key} in BUFFALO split mode. Skipping.")
                         continue
             except (ValueError, TypeError, IndexError) as e:
                 logging.warning(f"{prefix} Error unpacking key {key} (BUFFALO split mode): {e}. Skipping.")
                 continue

             po_str = str(po_key_val) if po_key_val is not None else "<MISSING_PO>"
             item_str = str(item_key_val) if item_key_val is not None else "<MISSING_ITEM>"
             desc_str = str(desc_key_val).strip() if desc_key_val is not None else ""
             is_buffalo = False
             if desc_str and "BUFFALO" in desc_str.upper(): is_buffalo = True
             
             # Use new col_ keys for sums
             sqft_sum = sums_dict.get('col_qty_sf', decimal.Decimal(0))
             amount_sum = sums_dict.get('col_amount', decimal.Decimal(0))
             
             # Fallback for legacy keys if not found (just in case)
             if sqft_sum == 0 and 'sqft_sum' in sums_dict: sqft_sum = sums_dict.get('sqft_sum', decimal.Decimal(0))
             if amount_sum == 0 and 'amount_sum' in sums_dict: amount_sum = sums_dict.get('amount_sum', decimal.Decimal(0))

             if not isinstance(sqft_sum, decimal.Decimal): sqft_sum = decimal.Decimal(0)
             if not isinstance(amount_sum, decimal.Decimal): amount_sum = decimal.Decimal(0)

             if is_buffalo:
                 buffalo_pos.add(po_str)
                 buffalo_items.add(item_str)
                 buffalo_descriptions.add(desc_str)
                 buffalo_sqft += sqft_sum
                 buffalo_amount += amount_sum
             else:
                 non_buffalo_pos.add(po_str)
                 non_buffalo_items.add(item_str)
                 if desc_str: non_buffalo_descriptions.add(desc_str)
                 non_buffalo_sqft += sqft_sum
                 non_buffalo_amount += amount_sum

        logging.debug(f"{prefix} Finished processing entries for BUFFALO split.")

        # Format BUFFALO Group ("1")
        sorted_buffalo_pos = sorted(list(buffalo_pos))
        sorted_buffalo_items = sorted(list(buffalo_items))
        sorted_buffalo_descriptions = sorted([d for d in buffalo_descriptions if d])
        buffalo_result: DAFCompoundingResult = {
            'col_po': format_chunks(sorted_buffalo_pos, DAF_CHUNK_SIZE, DAF_INTRA_CHUNK_SEPARATOR, DAF_INTER_CHUNK_SEPARATOR),
            'col_item': format_chunks(sorted_buffalo_items, DAF_CHUNK_SIZE, DAF_INTRA_CHUNK_SEPARATOR, DAF_INTER_CHUNK_SEPARATOR),
            'col_desc': format_chunks(sorted_buffalo_descriptions, 1, "", "\n"),
            'col_qty_sf': buffalo_sqft,
            'col_amount': buffalo_amount
        }
        # Format NON-BUFFALO Group ("2")
        sorted_non_buffalo_pos = sorted(list(non_buffalo_pos))
        sorted_non_buffalo_items = sorted(list(non_buffalo_items))
        sorted_non_buffalo_descriptions = sorted([d for d in non_buffalo_descriptions if d])
        non_buffalo_result: DAFCompoundingResult = {
            'col_po': format_chunks(sorted_non_buffalo_pos, DAF_CHUNK_SIZE, DAF_INTRA_CHUNK_SEPARATOR, DAF_INTER_CHUNK_SEPARATOR),
            'col_item': format_chunks(sorted_non_buffalo_items, DAF_CHUNK_SIZE, DAF_INTRA_CHUNK_SEPARATOR, DAF_INTER_CHUNK_SEPARATOR),
            'col_desc': format_chunks(sorted_non_buffalo_descriptions, 1, "", "\n"),
            'col_qty_sf': non_buffalo_sqft,
            'col_amount': non_buffalo_amount
        }
        # Construct Final Result LIST for BUFFALO Split Case
        final_buffalo_split_result: FinalDAFResultType = [
            buffalo_result,
            non_buffalo_result
        ]
        logging.info(f"{prefix} BUFFALO split DAF Compounding complete.")
        return final_buffalo_split_result
        # --- End Path 1 (BUFFALO Split) --- #

    else:
        # --- Path 2: Descriptions are NOT present -> PO Count Split Aggregation ---
        # Totals are calculated based on conceptual groups of 8 POs.
        # Final string formatting uses chunk size 2.
        PO_GROUPING_FOR_TOTALS = 5 # Define the size for grouping totals
        logging.info(f"{prefix} No description data found. Performing PO count split aggregation.")
        logging.info(f"{prefix}   - Totals calculated per group of {PO_GROUPING_FOR_TOTALS} POs.")
        logging.info(f"{prefix}   - String formatting uses chunk size {DAF_CHUNK_SIZE} and separator '{DAF_INTRA_CHUNK_SEPARATOR}'.")

        # Step 1: Aggregate data by PO
        po_data_aggregation: Dict[str, Dict[str, Union[set, decimal.Decimal]]] = {}
        logging.debug(f"{prefix} Pass 1: Aggregating SQFT/Amount/Items per PO.")
        for key, sums_dict in initial_results.items():
             po_key_val, item_key_val = None, None
             try: # Extract PO/Item
                 if len(key) >= 2: po_key_val, item_key_val = key[0], key[1]
                 else: continue
             except (TypeError, IndexError) as e: continue # Ignore errors in pass 1

             po_str = str(po_key_val) if po_key_val is not None else "<MISSING_PO>"
             item_str = str(item_key_val) if item_key_val is not None else "<MISSING_ITEM>"
             
             # Use new col_ keys
             sqft_sum = sums_dict.get('col_qty_sf', decimal.Decimal(0))
             amount_sum = sums_dict.get('col_amount', decimal.Decimal(0))
             
             # Fallback
             if sqft_sum == 0 and 'sqft_sum' in sums_dict: sqft_sum = sums_dict.get('sqft_sum', decimal.Decimal(0))
             if amount_sum == 0 and 'amount_sum' in sums_dict: amount_sum = sums_dict.get('amount_sum', decimal.Decimal(0))

             if not isinstance(sqft_sum, decimal.Decimal): sqft_sum = decimal.Decimal(0)
             if not isinstance(amount_sum, decimal.Decimal): amount_sum = decimal.Decimal(0)

             if po_str not in po_data_aggregation:
                 po_data_aggregation[po_str] = {'sqft_total': decimal.Decimal(0), 'amount_total': decimal.Decimal(0), 'items': set()}
             po_data_aggregation[po_str]['sqft_total'] += sqft_sum # type: ignore
             po_data_aggregation[po_str]['amount_total'] += amount_sum # type: ignore
             po_data_aggregation[po_str]['items'].add(item_str) # type: ignore

        if not po_data_aggregation:
            logging.warning(f"{prefix} No valid PO data found for PO count splitting. Returning empty dict.")
            return {}

        # Step 2: Get sorted list of unique POs
        sorted_pos = sorted(list(po_data_aggregation.keys()))

        # Step 3: Iterate through POs in conceptual groups of 8 for total calculation
        final_po_count_split_result: FinalDAFResultType = []
        # Calculate number of output chunks based on the total grouping size
        num_conceptual_chunks = (len(sorted_pos) + PO_GROUPING_FOR_TOTALS - 1) // PO_GROUPING_FOR_TOTALS

        logging.debug(f"{prefix} Pass 2: Creating {num_conceptual_chunks} output chunks based on conceptual PO groups of {PO_GROUPING_FOR_TOTALS}.")

        for i in range(num_conceptual_chunks):
            # Determine the POs belonging to this conceptual chunk (for totals)
            start_idx = i * PO_GROUPING_FOR_TOTALS
            end_idx = start_idx + PO_GROUPING_FOR_TOTALS
            conceptual_po_chunk = sorted_pos[start_idx:end_idx]

            # Calculate totals and collect items for THIS conceptual chunk
            chunk_sqft_total = decimal.Decimal(0)
            chunk_amount_total = decimal.Decimal(0)
            chunk_items = set()
            po_list_for_formatting = [] # Collect POs in this chunk for formatting

            for po_str in conceptual_po_chunk:
                po_agg_data = po_data_aggregation.get(po_str)
                if po_agg_data:
                    chunk_sqft_total += po_agg_data.get('sqft_total', decimal.Decimal(0)) # type: ignore
                    chunk_amount_total += po_agg_data.get('amount_total', decimal.Decimal(0)) # type: ignore
                    chunk_items.update(po_agg_data.get('items', set())) # type: ignore
                    po_list_for_formatting.append(po_str) # Add the PO itself to the list for formatting
                else:
                     logging.warning(f"{prefix} PO '{po_str}' not found in aggregation data during chunking.")

            # Sort items collected for this chunk
            sorted_chunk_items = sorted(list(chunk_items))

            # Step 4: Format the collected POs and Items using desired format (size 2)
            formatted_po_chunk = format_chunks(po_list_for_formatting, DAF_CHUNK_SIZE, DAF_INTRA_CHUNK_SEPARATOR, DAF_INTER_CHUNK_SEPARATOR)
            formatted_item_chunk = format_chunks(sorted_chunk_items, DAF_CHUNK_SIZE, DAF_INTRA_CHUNK_SEPARATOR, DAF_INTER_CHUNK_SEPARATOR)

            # Create the result dictionary for this chunk index
            chunk_result: DAFCompoundingResult = {
                'col_po': formatted_po_chunk,
                'col_item': formatted_item_chunk,
                'col_desc': '', # No descriptions in this path
                'col_qty_sf': chunk_sqft_total,    # Use CHUNK total (calculated based on group of 8)
                'col_amount': chunk_amount_total   # Use CHUNK total (calculated based on group of 8)
            }
            chunk_index_str = str(i + 1)
            final_po_count_split_result.append(chunk_result)
            logging.debug(f"{prefix} Created output chunk {chunk_index_str}: {len(conceptual_po_chunk)} POs contributed totals, SQFT={chunk_sqft_total}, Amount={chunk_amount_total}")

        logging.info(f"{prefix} PO count split DAF Compounding complete ({len(final_po_count_split_result)} chunks created).")
        return final_po_count_split_result
        # --- End Path 2 (PO Count Split) --- #


# --- >>> ADDED: Default JSON Serializer Function <<< ---
def json_serializer_default(obj):
    """JSON serializer for objects not serializable by default json code"""
    if isinstance(obj, (datetime.datetime, datetime.date)):
        return obj.isoformat() # Convert date/datetime to ISO string format
    elif isinstance(obj, decimal.Decimal): # Keep Decimal handling here too
        return str(obj)
    elif isinstance(obj, set): # Optional: Handle sets if needed
        return list(obj)
    # Add other custom types if needed
    # elif isinstance(obj, YourCustomClass):
    #     return obj.__dict__
    raise TypeError (f"Object of type {obj.__class__.__name__} is not JSON serializable")
# --- >>> END OF ADDED FUNCTION <<< ---


# Helper function to make data JSON serializable
# Handles tuple keys in aggregation results
def make_json_serializable(data):
    """Recursively converts tuple keys in dicts to strings and handles non-serializable types."""
    # NOTE: Using the default serializer for json.dumps handles Decimal and datetime now.
    # This function primarily focuses on converting tuple keys.
    if isinstance(data, dict):
        # Convert all keys to string, including tuple keys
        return {str(k): make_json_serializable(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [make_json_serializable(item) for item in data]
    elif data is None:
        return None # JSON null
    # Let the default handler in json.dumps deal with Decimal, datetime, etc.
    return data

# <<< MODIFIED FUNCTION SIGNATURE >>>
# Import PipelineMonitor

from core.utils.pipeline_monitor import PipelineMonitor
from core.utils.snitch import snitch

# ... (Previous code)

@snitch
def run_invoice_automation(
    input_excel_override: str = None,
    output_dir_override: str = None,
    monitor_override: PipelineMonitor = None
) -> Tuple[Path, str]:
    """
    Main entry point for the invoice automation process.
    Refactored to be callable as a library function.
    """
    # 1. Determine Output Directory (Fast Fail)
    if output_dir_override:
        output_dir = Path(output_dir_override).resolve()
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            raise RuntimeError(f"Invalid output directory specified: {output_dir}")
    else:
        from core.system_config import sys_config
        output_dir = sys_config.temp_uploads_dir

    # 2. Determine Input File (Prep for Monitor)
    input_filepath = input_excel_override or getattr(cfg, 'INPUT_EXCEL_FILE', 'unknown.xlsx')
    input_name = Path(input_filepath).name
    
    # 3. Setup Monitor
    monitor_output_path = output_dir / f"{Path(input_name).stem}_parser.json"
    args = argparse.Namespace(input=input_filepath, output_dir=str(output_dir))
    
    # WRAP EXECUTION
    with PipelineMonitor(monitor_output_path, args=args, step_name="Data Parser") as monitor:
        start_time = time.time()
        logging.info("--- Starting Invoice Automation ---")
        monitor.update_logs("input_file", input_name) # Record inputs

        # -------------------------------------------------------------
        # Re-Validate Input File inside Monitor to capture errors
        # -------------------------------------------------------------
        if not input_excel_override:
             try:
                 input_filepath = cfg.INPUT_EXCEL_FILE
                 logging.info(f"Using input Excel path from config.py: {input_filepath}")
             except Exception as e:
                 monitor.log_process_item("Configuration", status="error", error=e)
                 raise RuntimeError("Input Excel file path is missing in config.")

        if not os.path.isfile(input_filepath):
             # Try relative resolution
             script_dir = os.path.dirname(__file__)
             potential_path = os.path.join(script_dir, input_filepath)
             if os.path.isfile(potential_path):
                 input_filepath = potential_path
                 logging.info(f"Resolved relative input path: {input_filepath}")
             else:
                 err = FileNotFoundError(f"Input Excel file not found: {input_filepath}")
                 monitor.log_process_item("Input File Check", status="error", error=err)
                 raise err
        
        input_filename = os.path.basename(input_filepath)
        logging.info(f"Processing workbook: {input_filename}")
        
        # ... [Rest of logic continues largely unchanged but inside this block] ...
        
        processed_tables: Dict[int, Dict[str, Any]] = {}
        all_tables_data: Dict[int, Dict[str, List[Any]]] = {}

        # Global definitions
        global_standard_aggregation_results: Dict[Tuple[Any, Any, Optional[decimal.Decimal], Optional[str]], Dict[str, decimal.Decimal]] = {}
        global_custom_aggregation_results: Dict[Tuple[Any, Any, Optional[str], None], Dict[str, decimal.Decimal]] = {}
        global_DAF_compounded_result: Optional[FinalDAFResultType] = None
        aggregation_mode_used = "standard"

        # Determine Aggregation Strategy
        use_custom_aggregation_for_DAF = False
        try:
            custom_prefixes = getattr(cfg, 'CUSTOM_AGGREGATION_WORKBOOK_PREFIXES', [])
            if not isinstance(custom_prefixes, list): custom_prefixes = []
            
            for prefix in custom_prefixes:
                 if input_filename.startswith(prefix):
                    use_custom_aggregation_for_DAF = True
                    aggregation_mode_used = "custom"
                    logging.info(f"Using CUSTOM aggregation (Prefix: {prefix})")
                    break
        except Exception as e:
            logging.error(f"Strategy determination error: {e}")
            monitor.log_warning(f"Aggregation strategy check failed: {e}")

        
        # --- PROCESSING STEPS ---
        try:
            logging.info(f"Loading workbook from: {input_filepath}")
            handler = ExcelHandler(input_filepath)
            sheet = handler.load_sheet(sheet_name=cfg.SHEET_NAME, data_only=True)
            if sheet is None: raise RuntimeError(f"Failed to load sheet from '{input_filepath}'.")
            
            actual_sheet_name = sheet.title
            
            # Header Detection
            smart_result = sheet_parser.find_and_map_smart_headers(sheet)
            if not smart_result: 
                 err = RuntimeError("Smart header detection failed.")
                 monitor.log_process_item("Header Detection", status="error", error=err)
                 raise err
            
            header_row, column_mapping = smart_result
            
            # Find Additional Tables
            additional_header_rows = sheet_parser.find_all_header_rows(
                sheet=sheet,
                search_pattern=cfg.HEADER_IDENTIFICATION_PATTERN,
                row_range=(header_row + 1, sheet.max_row),
                col_range=(cfg.HEADER_SEARCH_COL_RANGE[0], cfg.HEADER_SEARCH_COL_RANGE[1])
            )
            all_header_rows = [header_row] + additional_header_rows
            monitor.update_logs("tables_found", len(all_header_rows))
            
            all_tables_data = sheet_parser.extract_multiple_tables(sheet, all_header_rows, column_mapping)
            



            # --- 5. Process Each Table (Instrumented) ---
            logging.info(f"--- Starting Data Processing Loop for {len(all_tables_data)} Extracted Table(s) ---")
            
            for table_index, raw_data_dict in all_tables_data.items():
                table_id_str = f"Table {table_index}"
                current_table_data = all_tables_data.get(table_index)
                
                if current_table_data is None: continue
                
                # Check for empty/invalid data
                if not isinstance(current_table_data, dict) or not current_table_data or not any(isinstance(v, list) and v for v in current_table_data.values()):
                     monitor.log_warning(f"{table_id_str} is empty or invalid. Skipping.")
                     processed_tables[table_index] = current_table_data
                     continue
                
                try:
                    # 5a. CBM
                    data_after_cbm = data_processor.process_cbm_column(current_table_data)
                    
                    # 5b. Distribute
                    try:
                        data_after_distribution = data_processor.distribute_values(data_after_cbm, cfg.COLUMNS_TO_DISTRIBUTE, cfg.DISTRIBUTION_BASIS_COLUMN)
                        processed_tables[table_index] = data_after_distribution
                        data_for_aggregation = data_after_distribution
                    except Exception as distrib_e:
                        # Log but continue with fallback
                        monitor.log_warning(f"{table_id_str}: Distribution failed ({distrib_e}). Using raw/CBM data.")
                        processed_tables[table_index] = data_after_cbm
                        data_for_aggregation = data_after_cbm
                    
                    # 5c. Initial Aggregation
                    if data_for_aggregation:
                         data_processor.aggregate_standard_by_po_item_price(data_for_aggregation, global_standard_aggregation_results)
                         data_processor.aggregate_custom_by_po_item(data_for_aggregation, global_custom_aggregation_results)
                    
                    monitor.log_process_item(table_id_str, status="success")
                except Exception as table_e:
                    # Log failure for this specific table but continue loop
                    monitor.log_process_item(table_id_str, status="error", error=table_e)

            # --- 6. DAF Compounding (Instrumented) ---
            try:
                # Determine strategy (re-using variables set earlier if accurate, or re-calculating simple version)
                # We reuse 'aggregation_mode_used' and 'use_custom_aggregation_for_DAF' calculated at start
                agg_source = global_custom_aggregation_results if use_custom_aggregation_for_DAF else global_standard_aggregation_results
                
                logging.info(f"Performing DAF Compounding (Mode: {aggregation_mode_used})")
                
                global_DAF_compounded_result = perform_DAF_compounding(agg_source, aggregation_mode_used)
                monitor.log_process_item("DAF Compounding", status="success")
            except Exception as daf_e:
                monitor.log_process_item("DAF Compounding", status="error", error=daf_e)

        except Exception as e:
            # Catch-all for the main Loading/Parsing/Extraction block
            monitor.log_process_item("Data Parsing/Extraction", status="error", error=e)
            raise e # Re-raise to trigger outer exit codes if needed, though monitor captures it



        # --- 7. Output / Further Steps ---
        logging.info(f"Final processed data structure contains {len(processed_tables)} table(s).")
        logging.info(f"Primary aggregation mode used for DAF Compounding: {aggregation_mode_used.upper()}")

        # --- Convert pallet_count to int ---
        for table_index, table_data in processed_tables.items():
            if 'pallet_count' in table_data and isinstance(table_data['pallet_count'], list):
                logging.info(f"Converting pallet_count in table {table_index}")
                for i, value in enumerate(table_data['pallet_count']):
                    if value is not None:
                        try:
                            original_value = value
                            converted_value = int(float(value))
                            table_data['pallet_count'][i] = converted_value
                            logging.info(f"Converted pallet_count[{i}] from '{original_value}' to {converted_value}")
                        except (ValueError, TypeError):
                            logging.warning(f"Could not convert pallet_count value '{value}' to int in table {table_index}, row {i}")
                            table_data['pallet_count'][i] = value  # Keep original value if conversion fails

        # Log the converted pallet_count
        for table_index, table_data in processed_tables.items():
            if 'pallet_count' in table_data:
                logging.info(f"Final pallet_count in table {table_index}: {table_data['pallet_count']} (types: {[type(v) for v in table_data['pallet_count']]})")
            # Log Standard Results
            log_str_std = pprint.pformat(global_standard_aggregation_results)
            if len(log_str_std) > MAX_LOG_DICT_LEN: log_str_std = log_str_std[:MAX_LOG_DICT_LEN] + "\n... (output truncated)"
            logging.debug(f"--- Full Global STANDARD Aggregation Results ---\n{log_str_std}")
            # Log Custom Results
            log_str_cust = pprint.pformat(global_custom_aggregation_results)
            if len(log_str_cust) > MAX_LOG_DICT_LEN: log_str_cust = log_str_cust[:MAX_LOG_DICT_LEN] + "\n... (output truncated)"
            logging.debug(f"--- Full Global CUSTOM Aggregation Results ---\n{log_str_cust}")


        # --- Log Final DAF Compounded Result (INFO Level) - Simplified to expect split result --- #
        logging.info(f"--- Final DAF Compounded Result (Workbook: '{input_filename}', Based on '{aggregation_mode_used.upper()}' Input) ---")
        if global_DAF_compounded_result is not None and isinstance(global_DAF_compounded_result, list):
            # Assume it's the BUFFALO split result or PO split result
            logging.info(f"DAF result is a list with {len(global_DAF_compounded_result)} groups.")
            for chunk_index, chunk_data in enumerate(global_DAF_compounded_result):
                logging.info(f"--- DAF Group {chunk_index + 1} --- ")
                if chunk_data and isinstance(chunk_data, dict):
                    po_string_value = chunk_data.get('col_po', '<Not Found>')
                    item_string_value = chunk_data.get('col_item', '<Not Found>')
                    desc_string_value = chunk_data.get('col_desc', '<Not Found>')
                    total_sqft_value = chunk_data.get('col_qty_sf', 'N/A')
                    total_amount_value = chunk_data.get('col_amount', 'N/A')

                    logging.info(f"  Combined POs:\n{po_string_value}")
                    logging.info(f"  Combined Items:\n{item_string_value}")
                    logging.info(f"  Combined Descriptions:\n{desc_string_value}")
                    logging.info(f"  Total SQFT: {total_sqft_value} (Type: {type(total_sqft_value)})")
                    logging.info(f"  Total Amount: {total_amount_value} (Type: {type(total_amount_value)})")
                else:
                    logging.info(f"  Group {chunk_index + 1} data not found or invalid.")
            logging.info("-" * 30)

        elif global_DAF_compounded_result is None:
            logging.error("DAF Compounding result is None or was not set.")
        else:
            # Handle unexpected type if necessary (e.g., empty dict if input was empty and aggregation failed)
             logging.warning(f"DAF Compounding result has unexpected structure/type: {type(global_DAF_compounded_result)}")

        # --- End Final Logging ---

        # --- Calculate Footer Data ---
        logging.info("--- Calculating Footer Data ---")
        
        # Calculate per-table totals
        table_footer_data = {}
        for table_id, table_data in processed_tables.items():
            table_footer_data[table_id] = data_processor.calculate_footer_totals(table_data)
            logging.info(f"Table {table_id} Footer: {table_footer_data[table_id]}")
        
        # Calculate grand total (merged across all tables)
        merged_processed_data = {
            "col_qty_pcs": [], "col_qty_sf": [], "col_net": [], "col_gross": [], "col_cbm": [], "col_amount": [], "col_pallet_count": [],
            "col_desc": [],  # Include both description field names for leather_summary calculation
            "col_po": [], "col_item": [], "col_unit_price": []  # Include fields for aggregate_per_po_with_pallets
        }
        for table_data in processed_tables.values():
            for key in merged_processed_data:
                if key in table_data:
                    merged_processed_data[key].extend(table_data[key])
        
        grand_total_footer = data_processor.calculate_footer_totals(merged_processed_data)
        logging.info(f"Grand Total Footer: {grand_total_footer}")
        
        # --- Calculate Add-on Data (Leather Summary) ---
        logging.info("--- Calculating Add-on Data ---")
        
        # Calculate leather summary (BUFFALO vs COW) across all tables
        leather_summary = data_processor.calculate_leather_summary(merged_processed_data)
        logging.info(f"Leather Summary: {leather_summary}")
        
        # Calculate normal aggregate per PO with pallets (group by PO + price)
        normal_aggregate_per_po = data_processor.aggregate_per_po_with_pallets(merged_processed_data)
        logging.info(f"Normal Aggregate Per PO: {len(normal_aggregate_per_po)} unique PO+price combinations")

        # --- 8. Generate JSON Output ---
        logging.info("--- Preparing Data for JSON Output ---")
        try:
            # Create the structure to be converted to JSON
            # Use the helper function to ensure serializability
            final_json_structure = {
                 "metadata": {
                    "workbook_filename": input_filename, # Use the actual input filename
                    "worksheet_name": actual_sheet_name,
                    "DAF_compounding_input_mode": aggregation_mode_used, # Clarify which mode fed DAF
                    "DAF_chunk_size": DAF_CHUNK_SIZE,
                     "DAF_intra_separator": DAF_INTRA_CHUNK_SEPARATOR.encode('unicode_escape').decode('utf-8'), # Encode escapes for JSON clarity
                    "DAF_inter_separator": DAF_INTER_CHUNK_SEPARATOR.encode('unicode_escape').decode('utf-8'), # Encode escapes for JSON clarity
                    "timestamp": datetime.datetime.now() # Add generation timestamp
                },
                # Include processed table data (potentially large)
                 # RENAME: processed_tables_data -> processed_tables_multi (Matches Config)
                 "processed_tables_multi": make_json_serializable(processed_tables),
                 
                 # Include Footer Data - both per-table and grand total
                 "footer_data": {
                     "table_totals": make_json_serializable(table_footer_data),  # Per-table totals
                     "grand_total": make_json_serializable(grand_total_footer),   # Overall grand total
                     "add_ons": {
                         "leather_summary_addon": make_json_serializable(leather_summary),  # BUFFALO vs COW summary
                     }
                 },

                # Include BOTH aggregation results explicitly (formatted as lists)
                # RENAME: standard_aggregation_results -> aggregation (Matches Config)
                "aggregation": data_processor.format_aggregation_as_list(global_standard_aggregation_results, mode='standard'),
                # RENAME: custom_aggregation_results -> aggregation_custom (Matches Suffix Rule)
                "aggregation_custom": data_processor.format_aggregation_as_list(global_custom_aggregation_results, mode='custom'),
                
                # Normal aggregate per PO with pallets (group by PO + price)
                # RENAME: normal_aggregate_per_po_with_pallets -> manifest_by_pallet_per_po (User Request)
                "manifest_by_pallet_per_po": make_json_serializable(normal_aggregate_per_po),

                # Include the final compounded result (derived from one of the above, based on mode)
                # RENAME: final_DAF_compounded_result -> aggregation_DAF (Matches Suffix Rule)
                "aggregation_DAF": make_json_serializable(global_DAF_compounded_result)
            }

             # Convert the structure to a JSON string (pretty-printed)
            json_output_string = json.dumps(final_json_structure,
                                            indent=4,
                                            default=json_serializer_default) # Use the default serializer

            # Log the JSON output (or a preview if too large)
            logging.info("--- Generated JSON Output ---")
            max_log_json_len = 5000
            if len(json_output_string) <= max_log_json_len:
                logging.info(json_output_string)
            else:
                logging.info(f"JSON output is large ({len(json_output_string)} chars). Logging preview:")
                logging.info(json_output_string[:max_log_json_len] + "\n... (JSON output truncated in log)")

            # --- MODIFIED: Save JSON using output_dir and simplified filename ---
            input_stem = Path(input_filename).stem # Get filename without extension
            json_output_filename = f"{input_stem}.json" # Simplified filename
            output_json_path = output_dir / json_output_filename # Combine output dir and filename

            logging.info(f"Determined output JSON path: {output_json_path}")
            try:
                with open(output_json_path, 'w', encoding='utf-8') as f_json:
                     f_json.write(json_output_string)
                logging.info(f"Successfully saved JSON output to '{output_json_path}'")
            except IOError as io_err:
                logging.error(f"Failed to write JSON output to file '{output_json_path}': {io_err}")
                raise io_err
            except Exception as write_err:
                 logging.error(f"An unexpected error occurred while writing JSON file: {write_err}", exc_info=True)
                 raise write_err

        except TypeError as json_err:
            logging.error(f"Failed to serialize data to JSON: {json_err}. Check data types and default handler.", exc_info=True)
            raise json_err
        except Exception as e:
            logging.error(f"An unexpected error occurred during JSON generation: {e}", exc_info=True)
            raise e
        # --- End JSON Generation ---

        logging.info(f"📁 Processed file: {input_filename}")
        
        return output_json_path, input_stem


if __name__ == "__main__":
    # --- Argument Parsing ---
    parser = argparse.ArgumentParser(description="Process an Excel invoice file to generate JSON data.")
    parser.add_argument(
        "--input-excel",
        type=str,
        default=None, # Default to None, indicating fallback to config.py
        help="Path to the input Excel file. Overrides the value in config.py if provided."
    )
    # --- ADDED: Output directory argument ---
    parser.add_argument(
        "--output-dir",
        type=str,
        default=None, # Default to None, indicating use CWD
        help="Directory to save the output JSON file. Defaults to the current working directory."
    )
    # --- END ADD ---
    args = parser.parse_args()
    # --- End Argument Parsing ---

    # --- Run the main logic ---
    # Pass the parsed arguments to the main function
    run_invoice_automation(
        input_excel_override=args.input_excel,
        output_dir_override=args.output_dir # Pass the output dir argument
    )
    # --- End Run Logic ---