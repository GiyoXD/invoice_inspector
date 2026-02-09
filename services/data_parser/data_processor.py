# --- START MODIFIED FILE: data_processor.py ---

import logging
from typing import Dict, List, Any, Optional, Tuple
import decimal # Use Decimal for precise calculations
import re
import pprint
# Import config values (consider passing as arguments)
from .config import DISTRIBUTION_BASIS_COLUMN # Keep this

# Set precision for Decimal calculations
decimal.getcontext().prec = 28 # Default precision, adjust if needed
# Define precision specifically for CBM results (e.g., 4 decimal places)
CBM_DECIMAL_PLACES = decimal.Decimal('0.0001')
# Define default precision for other distributions (e.g., 4 decimal places)
DEFAULT_DIST_PRECISION = decimal.Decimal('0.0001')

## TODO: make sure all the aggregation DAF mode has price support 10 floating point


class ProcessingError(Exception):
    """Custom exception for data processing errors."""
    pass

def _convert_to_decimal(value: Any, context: str = "") -> Optional[decimal.Decimal]:
    """Safely convert a value to Decimal, logging errors.
    
    For floats, we round to a reasonable precision to avoid floating-point errors
    like 0.30000000000000004 becoming an issue.
    """
    prefix = "[_convert_to_decimal]"
    if isinstance(value, decimal.Decimal):
        return value
    if value is None:
        return None
    
    # Handle floats specially to avoid floating-point precision issues
    # Round to 14 decimal places before converting to avoid issues like 0.30000000000000004
    if isinstance(value, float):
        # Use string formatting with precision to clean up float representation
        value_str = f"{value:.14f}".rstrip('0').rstrip('.')
        if not value_str or value_str == '-':
            return None
        try:
            return decimal.Decimal(value_str)
        except (decimal.InvalidOperation, TypeError, ValueError) as e:
            logging.warning(f"{prefix} Could not convert float '{value}' to Decimal {context}: {e}")
            return None
    
    value_str = str(value).strip()
    if not value_str:
        return None
    try:
        result = decimal.Decimal(value_str)
        return result
    except (decimal.InvalidOperation, TypeError, ValueError) as e:
        logging.warning(f"{prefix} Could not convert '{value}' (Str: '{value_str}') to Decimal {context}: {e}")
        return None

# _calculate_single_cbm function remains unchanged...
def _calculate_single_cbm(cbm_value: Any, row_index: int) -> Optional[decimal.Decimal]:
    """
    Parses a CBM string (e.g., "L*W*H" or "LxWxH") and calculates the volume.

    Args:
        cbm_value: The value from the CBM cell (can be string, number, None).
        row_index: The 0-based index of the row for logging purposes.

    Returns:
        The calculated CBM as a Decimal, or None if parsing fails or input is invalid.
    """
    prefix = "[_calculate_single_cbm]"
    log_context = f"for CBM at row index {row_index}" # Use 0-based index internally

    if cbm_value is None:
        logging.debug(f"{prefix} Input CBM value is None. {log_context}")
        return None

    # If it's already a number, convert to Decimal and quantize
    if isinstance(cbm_value, (int, float, decimal.Decimal)):
        logging.debug(f"{prefix} Input CBM is already numeric: {cbm_value}. {log_context}")
        calculated = _convert_to_decimal(cbm_value, log_context)
        if calculated is not None:
             result = calculated.quantize(CBM_DECIMAL_PLACES, rounding=decimal.ROUND_HALF_UP)
             logging.debug(f"{prefix} Quantized pre-numeric CBM to {result}. {log_context}")
             return result
        else:
             # Conversion should ideally not fail here, but handle it
             logging.warning(f"{prefix} Failed to convert pre-numeric CBM value {cbm_value} to Decimal. {log_context}")
             return None


    if not isinstance(cbm_value, str):
        logging.warning(f"{prefix} Unexpected type '{type(cbm_value).__name__}' for CBM value '{cbm_value}'. Cannot calculate. {log_context}")
        return None

    cbm_str = cbm_value.strip()
    if not cbm_str:
        logging.debug(f"{prefix} Input CBM string is empty after strip. {log_context}")
        return None

    logging.debug(f"{prefix} Attempting to parse CBM string: '{cbm_str}'. {log_context}")

    # Try splitting by '*' first
    parts = cbm_str.split('*')
    separator_used = "'*'"

    # If not 3 parts, try splitting by 'x' or 'X' (case-insensitive)
    if len(parts) != 3:
        if '*' not in cbm_str and ('x' in cbm_str.lower()):
             parts = re.split(r'[xX]', cbm_str) # Split by 'x' or 'X'
             separator_used = "'x' or 'X'"
             logging.debug(f"{prefix} Split by '*' failed, trying split by {separator_used}. Parts: {parts}. {log_context}")

    # Check if we have exactly 3 parts after trying separators
    if len(parts) != 3:
        logging.warning(f"{prefix} Invalid CBM format: '{cbm_str}'. Expected 3 parts separated by '*' or 'x'. Found {len(parts)} parts: {parts}. {log_context}")
        return None

    try:
        # Convert each part to Decimal
        dims = []
        valid_dims = True
        for i, part in enumerate(parts):
             dim = _convert_to_decimal(part, f"{log_context}, part {i+1} ('{part}')")
             if dim is None:
                 logging.warning(f"{prefix} Failed to convert dimension part {i+1} ('{part}') to Decimal. {log_context}")
                 valid_dims = False
             dims.append(dim)

        if not valid_dims:
            logging.warning(f"{prefix} Failed to convert one or more dimensions for CBM string '{cbm_str}'. Cannot calculate volume. {log_context}")
            return None

        dim1, dim2, dim3 = dims
        volume = (dim1 * dim2 * dim3).quantize(CBM_DECIMAL_PLACES, rounding=decimal.ROUND_HALF_UP)
        logging.debug(f"{prefix} Calculated CBM volume: {volume} from '{cbm_str}' (Dims: {dims}). {log_context}")
        return volume

    except Exception as e:
        logging.error(f"{prefix} Unexpected error calculating CBM from '{cbm_str}': {e}. {log_context}", exc_info=True)
        return None

# process_cbm_column function remains unchanged...
# --- START MODIFIED FILE: data_processor.py ---

import logging
from typing import Dict, List, Any, Optional, Tuple
import decimal # Use Decimal for precise calculations
import re
import pprint
# Import config values (consider passing as arguments)
from .config import DISTRIBUTION_BASIS_COLUMN # Keep this

# Set precision for Decimal calculations
decimal.getcontext().prec = 28 # Default precision, adjust if needed
# Define precision specifically for CBM results (e.g., 4 decimal places)
CBM_DECIMAL_PLACES = decimal.Decimal('0.0001')
# Define default precision for other distributions (e.g., 4 decimal places)
DEFAULT_DIST_PRECISION = decimal.Decimal('0.0001')


class ProcessingError(Exception):
    """Custom exception for data processing errors."""
    pass

def _convert_to_decimal(value: Any, context: str = "") -> Optional[decimal.Decimal]:
    """Safely convert a value to Decimal, logging errors.
    
    For floats, we round to a reasonable precision to avoid floating-point errors
    like 0.30000000000000004 becoming an issue.
    """
    prefix = "[_convert_to_decimal]"
    if isinstance(value, decimal.Decimal):
        return value
    if value is None:
        return None
    
    # Handle floats specially to avoid floating-point precision issues
    # Round to 14 decimal places before converting to avoid issues like 0.30000000000000004
    if isinstance(value, float):
        # Use string formatting with precision to clean up float representation
        value_str = f"{value:.14f}".rstrip('0').rstrip('.')
        if not value_str or value_str == '-':
            return None
        try:
            return decimal.Decimal(value_str)
        except (decimal.InvalidOperation, TypeError, ValueError) as e:
            logging.warning(f"{prefix} Could not convert float '{value}' to Decimal {context}: {e}")
            return None
    
    value_str = str(value).strip()
    if not value_str:
        return None
    try:
        result = decimal.Decimal(value_str)
        return result
    except (decimal.InvalidOperation, TypeError, ValueError) as e:
        logging.warning(f"{prefix} Could not convert '{value}' (Str: '{value_str}') to Decimal {context}: {e}")
        return None

# _calculate_single_cbm function remains unchanged...
def _calculate_single_cbm(cbm_value: Any, row_index: int) -> Optional[decimal.Decimal]:
    """
    Parses a CBM string (e.g., "L*W*H" or "LxWxH") and calculates the volume.

    Args:
        cbm_value: The value from the CBM cell (can be string, number, None).
        row_index: The 0-based index of the row for logging purposes.

    Returns:
        The calculated CBM as a Decimal, or None if parsing fails or input is invalid.
    """
    prefix = "[_calculate_single_cbm]"
    log_context = f"for CBM at row index {row_index}" # Use 0-based index internally

    if cbm_value is None:
        logging.debug(f"{prefix} Input CBM value is None. {log_context}")
        return None

    # If it's already a number, convert to Decimal and quantize
    if isinstance(cbm_value, (int, float, decimal.Decimal)):
        logging.debug(f"{prefix} Input CBM is already numeric: {cbm_value}. {log_context}")
        calculated = _convert_to_decimal(cbm_value, log_context)
        if calculated is not None:
             result = calculated.quantize(CBM_DECIMAL_PLACES, rounding=decimal.ROUND_HALF_UP)
             logging.debug(f"{prefix} Quantized pre-numeric CBM to {result}. {log_context}")
             return result
        else:
             # Conversion should ideally not fail here, but handle it
             logging.warning(f"{prefix} Failed to convert pre-numeric CBM value {cbm_value} to Decimal. {log_context}")
             return None


    if not isinstance(cbm_value, str):
        logging.warning(f"{prefix} Unexpected type '{type(cbm_value).__name__}' for CBM value '{cbm_value}'. Cannot calculate. {log_context}")
        return None

    cbm_str = cbm_value.strip()
    if not cbm_str:
        logging.debug(f"{prefix} Input CBM string is empty after strip. {log_context}")
        return None

    logging.debug(f"{prefix} Attempting to parse CBM string: '{cbm_str}'. {log_context}")

    # Try splitting by '*' first
    parts = cbm_str.split('*')
    separator_used = "'*'"

    # If not 3 parts, try splitting by 'x' or 'X' (case-insensitive)
    if len(parts) != 3:
        if '*' not in cbm_str and ('x' in cbm_str.lower()):
             parts = re.split(r'[xX]', cbm_str) # Split by 'x' or 'X'
             separator_used = "'x' or 'X'"
             logging.debug(f"{prefix} Split by '*' failed, trying split by {separator_used}. Parts: {parts}. {log_context}")

    # Check if we have exactly 3 parts after trying separators
    if len(parts) != 3:
        logging.warning(f"{prefix} Invalid CBM format: '{cbm_str}'. Expected 3 parts separated by '*' or 'x'. Found {len(parts)} parts: {parts}. {log_context}")
        return None

    try:
        # Convert each part to Decimal
        dims = []
        valid_dims = True
        for i, part in enumerate(parts):
             dim = _convert_to_decimal(part, f"{log_context}, part {i+1} ('{part}')")
             if dim is None:
                 logging.warning(f"{prefix} Failed to convert dimension part {i+1} ('{part}') to Decimal. {log_context}")
                 valid_dims = False
             dims.append(dim)

        if not valid_dims:
            logging.warning(f"{prefix} Failed to convert one or more dimensions for CBM string '{cbm_str}'. Cannot calculate volume. {log_context}")
            return None

        dim1, dim2, dim3 = dims
        volume = (dim1 * dim2 * dim3).quantize(CBM_DECIMAL_PLACES, rounding=decimal.ROUND_HALF_UP)
        logging.debug(f"{prefix} Calculated CBM volume: {volume} from '{cbm_str}' (Dims: {dims}). {log_context}")
        return volume

    except Exception as e:
        logging.error(f"{prefix} Unexpected error calculating CBM from '{cbm_str}': {e}. {log_context}", exc_info=True)
        return None

# process_cbm_column function remains unchanged...
def process_cbm_column(raw_data: Dict[str, List[Any]]) -> Dict[str, List[Any]]:
    """
    Iterates through the 'cbm' list in raw_data, calculates numeric CBM values
    from strings (L*W*H or LxWxH format), and updates the list in place.
    """
    prefix = "[process_cbm_column]"
    cbm_key = 'col_cbm' # Canonical name

    if cbm_key not in raw_data:
        logging.debug(f"{prefix} No '{cbm_key}' column found in this table's extracted data. Skipping CBM calculation.")
        return raw_data

    original_cbm_list = raw_data.get(cbm_key) # Get list
    # Check if it's actually a list and not None or empty
    if not isinstance(original_cbm_list, list):
         logging.warning(f"{prefix} Key '{cbm_key}' exists but is not a list (Type: {type(original_cbm_list).__name__}). Skipping CBM calculation.")
         return raw_data
    if not original_cbm_list:
        logging.debug(f"{prefix} '{cbm_key}' column is present but the list is empty. Skipping CBM calculation.")
        return raw_data

    logging.info(f"{prefix} Processing '{cbm_key}' column for volume calculations (List length: {len(original_cbm_list)})...")
    calculated_cbm_list = []
    num_rows = len(original_cbm_list)

    # Process each value in the original list
    for i in range(num_rows):
        value = original_cbm_list[i]
        calculated_value = _calculate_single_cbm(value, i) # Calculate volume using the helper
        calculated_cbm_list.append(calculated_value) # Add Decimal or None

    # Replace the original list in the dictionary with the newly calculated list
    raw_data[cbm_key] = calculated_cbm_list
    logging.info(f"{prefix} Finished processing '{cbm_key}' column. List now contains calculated values (Decimals or Nones).")
    return raw_data

# distribute_values function remains unchanged...
def distribute_values(
    raw_data: Dict[str, List[Any]],
    columns_to_distribute: List[str],
    basis_column: str
) -> Dict[str, List[Any]]:
    """
    Distributes values in specified columns based on proportions in the basis column.
    Operates on the input raw_data (which might have pre-calculated CBM).
    Handles pre-calculated CBM decimals correctly. Modifies data in place.
    """
    prefix = "[distribute_values]"
    logging.debug(f"{prefix} Starting value distribution process.")

    if not raw_data:
        logging.warning(f"{prefix} Received empty raw_data dictionary. Skipping distribution.")
        return {} # Return empty if input is empty

    processed_data = raw_data # Operate directly on the input dictionary

    # --- Input Validation ---
    if not isinstance(raw_data, dict):
        logging.error(f"{prefix} Input 'raw_data' is not a dictionary (Type: {type(raw_data).__name__}). Cannot distribute.")
        raise ProcessingError("Input data for distribution must be a dictionary.")

    if basis_column not in processed_data:
        # Try finding the basis column with 'col_' prefix or specific mapping
        candidate_basis = basis_column
        if not candidate_basis.startswith('col_'):
            if basis_column == 'pcs':
                 candidate_basis = 'col_qty_pcs'
            elif basis_column == 'sqft':
                 candidate_basis = 'col_qty_sf'
            else:
                 candidate_basis = f"col_{basis_column}"
        
        if candidate_basis in processed_data:
             basis_column = candidate_basis
        else:
            logging.error(f"{prefix} Basis column '{basis_column}' (or '{candidate_basis}') not found in data dictionary keys: {list(processed_data.keys())}. Cannot distribute.")
            raise ProcessingError(f"Basis column '{basis_column}' not found for distribution.")

    basis_values_list = processed_data.get(basis_column)
    if not isinstance(basis_values_list, list):
        logging.error(f"{prefix} Basis column '{basis_column}' key exists but value is not a list (Type: {type(basis_values_list).__name__}). Cannot distribute.")
        raise ProcessingError(f"Basis column '{basis_column}' data is not a list.")


    valid_columns_to_distribute = []
    if columns_to_distribute: # Ensure it's not None or empty
        for col in columns_to_distribute:
            # Map legacy config names to new col_ names if needed
            target_col = col
            if not target_col.startswith('col_'):
                if target_col == 'net': target_col = 'col_net'
                elif target_col == 'gross': target_col = 'col_gross'
                elif target_col == 'cbm': target_col = 'col_cbm'
                elif target_col == 'sqft': target_col = 'col_qty_sf'
                elif target_col == 'pcs': target_col = 'col_qty_pcs'
                elif target_col == 'amount': target_col = 'col_amount'
                elif target_col == 'pallet_count': target_col = 'col_pallet_count'
                else: target_col = f"col_{target_col}"

            if target_col not in processed_data:
                 logging.warning(f"{prefix} Column '{col}' (mapped to '{target_col}') specified for distribution but not found in this table's data keys: {list(processed_data.keys())}. Skipping this column.")
            elif not isinstance(processed_data.get(target_col), list):
                 logging.warning(f"{prefix} Column '{col}' (mapped to '{target_col}') specified for distribution exists but is not a list (Type: {type(processed_data.get(target_col)).__name__}). Skipping this column.")
            else:
                valid_columns_to_distribute.append(target_col)
    else:
        logging.info(f"{prefix} No columns specified in 'columns_to_distribute' list. Skipping distribution.")
        return processed_data


    if not valid_columns_to_distribute:
         logging.warning(f"{prefix} No valid columns found to perform distribution on after checking existence and type. Columns requested: {columns_to_distribute}")
         return processed_data # Return unmodified data

    num_rows = len(basis_values_list)
    if num_rows == 0:
        logging.info(f"{prefix} Basis column '{basis_column}' is empty (0 rows). Skipping distribution.")
        return processed_data

    logging.info(f"{prefix} Starting value distribution for columns: {valid_columns_to_distribute} based on '{basis_column}' ({num_rows} rows).")

    # Pre-convert basis values to Decimal
    basis_values_dec: List[Optional[decimal.Decimal]] = [
        _convert_to_decimal(val, f"{prefix} basis column '{basis_column}' row index {i}")
        for i, val in enumerate(basis_values_list)
    ]
    logging.debug(f"{prefix} Pre-converted basis values (first 10): {basis_values_dec[:10]}")

    # --- Process each column ---
    for col_name in valid_columns_to_distribute:
        logging.info(f"{prefix} Processing column for distribution: '{col_name}'")
        original_col_values = processed_data.get(col_name, []) # Should be a list based on checks above

        # Final check on length matching
        if len(original_col_values) != num_rows:
             logging.error(f"{prefix} Row count mismatch detected just before processing! Basis '{basis_column}' ({num_rows}) vs '{col_name}' ({len(original_col_values)}). This indicates a potential data integrity issue. Skipping distribution for '{col_name}'.")
             continue # Skip this column

        # Pre-convert original values for the column being distributed
        current_col_values_dec: List[Optional[decimal.Decimal]] = [
             # Keep existing Decimals (e.g., from CBM calc), attempt conversion otherwise
             val if isinstance(val, decimal.Decimal)
             else _convert_to_decimal(val, f"{prefix} column '{col_name}' row index {i}")
             for i, val in enumerate(original_col_values)
        ]
        logging.debug(f"{prefix} Pre-converted values for '{col_name}' (first 10): {current_col_values_dec[:10]}")


        # Initialize processed list for this column
        processed_col_values: List[Optional[decimal.Decimal]] = [None] * num_rows

        i = 0 # Main loop index
        while i < num_rows:
            current_val_dec = current_col_values_dec[i]
            log_row_context = f"{prefix} Col '{col_name}', Row index {i}"

            # --- Case 1: Found a non-None, non-zero value to potentially distribute ---
            if current_val_dec is not None and current_val_dec != decimal.Decimal(0):
                logging.debug(f"{log_row_context}: Found distributable value: {current_val_dec}")
                # Store the original non-zero value at its position
                processed_col_values[i] = current_val_dec

                # --- Look ahead for the distribution block ---
                j = i + 1 # Lookahead index
                distribution_rows_indices = [] # Indices of rows following i that are empty/zero in this col
                while j < num_rows:
                     next_original_val_dec = current_col_values_dec[j]
                     # Stop lookahead if the *next* original value is non-empty/non-zero
                     if next_original_val_dec is not None and next_original_val_dec != decimal.Decimal(0):
                          logging.debug(f"{log_row_context}: Lookahead stopped at index {j}. Found non-empty/zero value {next_original_val_dec} in original data.")
                          break

                     # Check basis value for this potential distribution row
                     basis_for_j = basis_values_dec[j]
                     if basis_for_j is not None:
                          # Include row j in the potential block, regardless of basis value (handle 0 basis later)
                          distribution_rows_indices.append(j)
                          logging.debug(f"{log_row_context}: Lookahead index {j} is part of block (Original val empty/zero, Basis={basis_for_j}).")
                     else:
                          # Basis is missing for row j. It's part of the block but cannot receive distribution.
                          distribution_rows_indices.append(j) # Still part of the block length calculation
                          logging.warning(f"{log_row_context}: Lookahead index {j} has MISSING basis. Will assign 0 later.")
                     j += 1
                # --- End of Look ahead ---
                logging.debug(f"{log_row_context}: Lookahead finished. Indices in distribution block (excluding start row {i}): {distribution_rows_indices}")

                # --- If a distribution block was found (rows followed the value) ---
                if distribution_rows_indices:
                    block_indices = [i] + distribution_rows_indices # All indices in the block
                    logging.debug(f"{log_row_context}: Identified distribution block indices: {block_indices}")

                    # --- Calculate total POSITIVE basis for the block ---
                    total_basis_in_block = decimal.Decimal(0)
                    indices_with_valid_basis = [] # Track rows that contribute > 0 basis

                    for k in block_indices:
                        basis_val = basis_values_dec[k]
                        if basis_val is not None and basis_val > 0:
                            total_basis_in_block += basis_val
                            indices_with_valid_basis.append(k)
                        elif basis_val is not None: # Log zero/negative basis
                             logging.debug(f"{log_row_context}: Basis value is zero or negative ({basis_val}) at index {k} in block. Excluded from total.")
                        # else: # Basis is None, already logged during lookahead

                    logging.debug(f"{log_row_context}: Block Calculation - Total POSITIVE basis: {total_basis_in_block}. Indices with positive basis: {indices_with_valid_basis}")

                    # --- Perform distribution if possible ---
                    if total_basis_in_block > 0 and indices_with_valid_basis:
                         distributed_sum_check = decimal.Decimal(0)
                         dist_precision = CBM_DECIMAL_PLACES if col_name == 'col_cbm' else DEFAULT_DIST_PRECISION

                         logging.debug(f"{log_row_context}: Distributing {current_val_dec} across {len(indices_with_valid_basis)} rows with positive basis using precision {dist_precision}.")

                         # Use remainder adjustment method to ensure exact sum
                         # Distribute to all rows EXCEPT the last one, then assign remainder to the last
                         num_valid_indices = len(indices_with_valid_basis)
                         
                         if num_valid_indices == 1:
                             # Only one row - assign the entire value
                             k = indices_with_valid_basis[0]
                             processed_col_values[k] = current_val_dec.quantize(dist_precision, rounding=decimal.ROUND_HALF_UP)
                             distributed_sum_check = processed_col_values[k]
                             logging.debug(f"{log_row_context}:   Index {k}: Single row, assigned full value={processed_col_values[k]}")
                         else:
                             # Multiple rows - use remainder adjustment
                             # Distribute to all but the last row
                             for k in indices_with_valid_basis[:-1]:
                                 basis_val = basis_values_dec[k]  # Known > 0
                                 proportion = basis_val / total_basis_in_block
                                 distributed_value = (current_val_dec * proportion).quantize(dist_precision, rounding=decimal.ROUND_HALF_UP)
                                 processed_col_values[k] = distributed_value
                                 distributed_sum_check += distributed_value
                                 logging.debug(f"{log_row_context}:   Index {k}: Basis={basis_val}, Prop={proportion:.6f}, Dist Val={distributed_value}")
                             
                             # Last row gets the exact remainder to ensure perfect sum
                             last_idx = indices_with_valid_basis[-1]
                             remainder = current_val_dec - distributed_sum_check
                             processed_col_values[last_idx] = remainder.quantize(dist_precision, rounding=decimal.ROUND_HALF_UP)
                             distributed_sum_check += processed_col_values[last_idx]
                             logging.debug(f"{log_row_context}:   Index {last_idx}: REMAINDER row, assigned={processed_col_values[last_idx]} (ensures exact sum)")

                         # Assign 0 to rows in the block that had missing/zero/negative basis
                         for k in block_indices:
                             if k not in indices_with_valid_basis:
                                 # Only assign 0 if it hasn't been assigned yet (should only be for k != i)
                                 if processed_col_values[k] is None:
                                     processed_col_values[k] = decimal.Decimal(0)
                                     log_reason = "missing basis" if basis_values_dec[k] is None else f"zero/negative basis ({basis_values_dec[k]})"
                                     logging.warning(f"{log_row_context}:   Index {k}: Assigning 0 due to {log_reason}.")


                         # --- Distribution Check (should always pass now with remainder adjustment) ---
                         tolerance = dist_precision / decimal.Decimal(2)
                         diff = abs(distributed_sum_check - current_val_dec)
                         if not diff <= tolerance:
                              logging.warning(f"{log_row_context}: Distribution Check potentially FAILED for block. Original: {current_val_dec}, Distributed Sum: {distributed_sum_check}, Difference: {diff:.10f} (Tolerance: {tolerance})")
                         else:
                              logging.debug(f"{log_row_context}: Distribution Check PASSED for block. Original: {current_val_dec}, Sum: {distributed_sum_check}")

                    else: # Cannot distribute (no positive basis found in the block)
                        logging.warning(f"{log_row_context}: Cannot distribute value {current_val_dec}. Total positive basis in block is zero or none found. Keeping original value at index {i}, setting others in block {distribution_rows_indices} to 0.")
                        # Ensure subsequent rows in the identified block are set to 0 if not already set
                        for k in distribution_rows_indices:
                            if processed_col_values[k] is None:
                                processed_col_values[k] = decimal.Decimal(0)

                    # Move main loop index past the processed block
                    i = j # Start next iteration after the block
                    logging.debug(f"{log_row_context}: End of block processing. Moving main index i to {i}")

                # --- Case 1b: Non-zero value found, but NO block followed ---
                else:
                    logging.debug(f"{log_row_context}: Value {current_val_dec} found, but no empty/zero rows followed. Keeping value as is.")
                    # The value processed_col_values[i] = current_val_dec was already set
                    i += 1 # Move to the next row normally

            # --- Case 2: Current original value is None or zero ---
            else:
                logging.debug(f"{log_row_context}: Original value is None or zero ('{current_col_values_dec[i]}').")
                # Check if this position was already filled by the distribution from a previous block
                if processed_col_values[i] is None:
                    # If not filled, set it explicitly to 0
                    logging.debug(f"{log_row_context}: Position was not filled by previous block, setting to 0.")
                    processed_col_values[i] = decimal.Decimal(0)
                else:
                     logging.debug(f"{log_row_context}: Position was already filled with {processed_col_values[i]} by a previous block's distribution.")
                i += 1 # Move to the next row

        # --- End of main loop (while i < num_rows) ---

        # Update the main data dictionary with the processed list (containing Decimals or Nones)
        processed_data[col_name] = processed_col_values
        logging.debug(f"{prefix} Finished processing column '{col_name}'. Final values (first 10): {processed_col_values[:10]}")
        logging.info(f"{prefix} Completed distribution processing for column: '{col_name}'.")


    logging.info(f"{prefix} Value distribution processing COMPLETED for all requested columns.")
    return processed_data # Return the dictionary with modified lists


# *** Standard Aggregation Function (MODIFIED to handle SQFT, AMOUNT, and DESCRIPTION key) ***
def aggregate_standard_by_po_item_price(
    processed_data: Dict[str, List[Any]],
    # UPDATED TYPE HINT: Key is now 4 elements, Value is Dict storing sums
    global_aggregation_map: Dict[Tuple[Any, Any, Optional[decimal.Decimal], Optional[str]], Dict[str, decimal.Decimal]] # Added Optional[str] for description
) -> Dict[Tuple[Any, Any, Optional[decimal.Decimal], Optional[str]], Dict[str, decimal.Decimal]]:
    """
    STANDARD Aggregation: Aggregates 'sqft' AND 'amount' values based on unique
    combinations of 'po', 'item', 'unit' price, AND 'description'.
    Updates the global_aggregation_map in place.

    Args:
        processed_data: Dictionary representing the data of the current table.
        global_aggregation_map: The dictionary holding the cumulative aggregation results.
                                  Key: (po, item, price, description)
                                  Value: Dict{'sqft_sum': Decimal, 'amount_sum': Decimal}.

    Returns:
        The updated global_aggregation_map.
    """
    aggregated_results = global_aggregation_map
    # UPDATED: Add 'description' to required columns (handle its absence later)
    required_cols = ['col_po', 'col_item', 'col_unit_price', 'col_qty_sf', 'col_amount'] # Keep description optional for now
    prefix = "[aggregate_standard]"

    logging.debug(f"{prefix} Updating global STANDARD aggregation (SQFT & Amount by PO/Item/Price/Desc) with new table data.")
    logging.debug(f"{prefix} Size of global map BEFORE processing this table: {len(aggregated_results)}")

    # --- Input Validation ---
    if not isinstance(processed_data, dict):
        logging.error(f"{prefix} Input 'processed_data' is not a dictionary. Cannot aggregate.")
        return aggregated_results

    missing_cols = [col for col in required_cols if col not in processed_data]
    if missing_cols:
        logging.warning(f"{prefix} Cannot perform STANDARD aggregation: Missing required columns {missing_cols}. Skipping aggregation for this table.")
        return aggregated_results

    # --- Check for optional 'description' column ---
    has_description_col = 'col_desc' in processed_data
    if has_description_col and not isinstance(processed_data.get('col_desc'), list):
        logging.warning(f"{prefix} 'col_desc' column exists but is not a list. Will use None for description keys.")
        has_description_col = False # Treat as missing if not a list
    elif not has_description_col:
        logging.info(f"{prefix} 'col_desc' column not found or is invalid. Will use None for description keys.")

    # Safely get lists and check types - Use get() with    # Get column data using new 'col_' keys
    po_list = processed_data.get('col_po', [])
    item_list = processed_data.get('col_item', [])
    unit_list = processed_data.get('col_unit_price', [])
    sqft_list = processed_data.get('col_qty_sf', [])
    amount_list = processed_data.get('col_amount', [])
    description_list = processed_data.get('col_desc', []) if has_description_col else [] # Get description list if valid

    # Use length of 'po' list as the reference number of rows
    num_rows = len(po_list)
    logging.debug(f"{prefix} Input data contains lists. Number of rows based on 'col_po' list: {num_rows}")

    # Check for length consistency across all required lists
    all_lists_to_check = {'col_item': item_list, 'col_unit_price': unit_list, 'col_qty_sf': sqft_list, 'col_amount': amount_list}
    # Add description list to check only if it exists and is supposed to be used
    if has_description_col:
        all_lists_to_check['col_desc'] = description_list

    if not all(len(lst) == num_rows for lst in all_lists_to_check.values()):
        lengths = {k: len(v) for k, v in all_lists_to_check.items()}
        lengths['col_po'] = num_rows # Add PO length for context
        logging.error(f"{prefix} Data length mismatch! Lengths:{lengths}. Aborting standard aggregation for this table.")
        return aggregated_results

    if num_rows == 0:
        logging.info(f"{prefix} No data rows found in this table. Global map unchanged.")
        return aggregated_results

    logging.info(f"{prefix} Processing {num_rows} rows for STANDARD aggregation (SQFT & Amount by PO/Item/Price/Desc).")

    # --- Iterate and Aggregate ---
    rows_processed_this_table = 0
    successful_conversions_sqft = 0
    successful_conversions_amount = 0

    for i in range(num_rows):
        rows_processed_this_table += 1
        log_row_context = f"{prefix} Table Row index {i}"
        logging.debug(f"{log_row_context} --- Processing ---")

        # Get raw values
        po_val, item_val = po_list[i], item_list[i]
        unit_price_raw, sqft_raw, amount_raw = unit_list[i], sqft_list[i], amount_list[i]
        # Get description if available, else None
        desc_raw = description_list[i] if has_description_col and i < len(description_list) else None

        logging.debug(f"{log_row_context}: Raw values - PO='{po_val}', Item='{item_val}', Price='{unit_price_raw}', Desc='{desc_raw}', SQFT='{sqft_raw}', Amount='{amount_raw}'")

        # Prepare key components
        po_key = str(po_val).strip() if isinstance(po_val, str) else po_val
        item_key = str(item_val).strip() if isinstance(item_val, str) else item_val
        description_key = str(desc_raw).strip() if isinstance(desc_raw, str) else desc_raw # Keep None as None
        description_key = description_key if description_key else None # Ensure empty strings become None

        po_key = po_key if po_key is not None else "<MISSING_PO>"
        item_key = item_key if item_key is not None else "<MISSING_ITEM>"
        # Description key can be None

        # Convert price to Decimal for the key
        price_dec = _convert_to_decimal(unit_price_raw, f"{log_row_context} price")

        # UPDATED Key: (PO, Item, Price, Description)
        key = (po_key, item_key, price_dec, description_key)
        logging.debug(f"{log_row_context}: Generated Key Tuple = {key}")


        # Convert SQFT and Amount to Decimal for summation
        sqft_dec = _convert_to_decimal(sqft_raw, f"{log_row_context} SQFT")
        if sqft_dec is None:
             # logging.debug(f"{log_row_context}: SQFT value '{sqft_raw}' is None or failed conversion. Using 0.") # Reduced verbosity
             sqft_dec = decimal.Decimal(0)
        else:
             successful_conversions_sqft +=1

        amount_dec = _convert_to_decimal(amount_raw, f"{log_row_context} Amount")
        if amount_dec is None:
            # logging.debug(f"{log_row_context}: Amount value '{amount_raw}' is None or failed conversion. Using 0.") # Reduced verbosity
            amount_dec = decimal.Decimal(0)
        else:
            successful_conversions_amount +=1

        # logging.debug(f"{log_row_context}: Converted values - SQFT='{sqft_dec}', Amount='{amount_dec}'") # Reduced verbosity

        # --- Add to the global aggregate sums (SQFT & Amount) ---
        current_sums = aggregated_results.get(key, {'sqft_sum': decimal.Decimal(0), 'amount_sum': decimal.Decimal(0)})

        # logging.debug(f"{log_row_context}: Sums for key {key} BEFORE add = {current_sums}") # Reduced verbosity

        # Update the sums
        current_sums['sqft_sum'] += sqft_dec
        current_sums['amount_sum'] += amount_dec

        # Store the updated dictionary back into the global map
        aggregated_results[key] = current_sums
        # logging.debug(f"{log_row_context}: Global sums for key {key} AFTER add = {aggregated_results[key]}") # Reduced verbosity


    logging.info(f"{prefix} Finished processing {rows_processed_this_table} rows.")
    logging.info(f"{prefix} SQFT values successfully converted/defaulted for {successful_conversions_sqft} rows.")
    logging.info(f"{prefix} Amount values successfully converted/defaulted for {successful_conversions_amount} rows.")
    logging.info(f"{prefix} Global standard aggregation map size: {len(aggregated_results)}")
    # logging.debug(f"{prefix} Global Standard Aggregated Results (End of Table):\n{pprint.pformat(aggregated_results)}") # Keep DEBUG for detailed tracing if needed
    return aggregated_results


# *** Custom Aggregation Function (MODIFIED to include DESCRIPTION key) ***
def aggregate_custom_by_po_item(
    processed_data: Dict[str, List[Any]],
    # UPDATED TYPE HINT: Key is now 4 elements (PO, Item, None, Description)
    global_custom_aggregation_map: Dict[Tuple[Any, Any, None, Optional[str]], Dict[str, decimal.Decimal]]
) -> Dict[Tuple[Any, Any, None, Optional[str]], Dict[str, decimal.Decimal]]:
    """
    CUSTOM Aggregation: Aggregates 'sqft' and 'amount' values based on unique
    combinations of 'po', 'item', AND 'description'. Uses a 4-element key
    (PO, Item, None, Description) for structural consistency with standard aggregation.
    Updates the global_custom_aggregation_map in place.

    Args:
        processed_data: Dictionary representing the data of the current table.
        global_custom_aggregation_map: The dictionary holding the cumulative custom
                                       aggregation results.
                                       Key: (po, item, None, description)
                                       Value: Dict{'sqft_sum': Decimal, 'amount_sum': Decimal}.

    Returns:
        The updated global_custom_aggregation_map.
    """
    aggregated_results = global_custom_aggregation_map
    # Required columns for this aggregation (Description is optional)
    required_cols = ['col_po', 'col_item', 'col_qty_sf', 'col_amount']
    prefix = "[aggregate_custom]"

    logging.debug(f"{prefix} Updating global CUSTOM aggregation (SQFT & Amount by PO/Item/Desc) with new table data.")
    logging.debug(f"{prefix} Size of global custom map BEFORE processing this table: {len(aggregated_results)}")

    # --- Input Validation ---
    if not isinstance(processed_data, dict):
        logging.error(f"{prefix} Input 'processed_data' is not a dictionary. Cannot aggregate.")
        return aggregated_results

    missing_cols = [col for col in required_cols if col not in processed_data]
    if missing_cols:
        logging.warning(f"{prefix} Cannot perform full CUSTOM aggregation: Missing required columns {missing_cols}. Proceeding cautiously.")
        # Allow proceeding, rows without needed data will be skipped or defaulted

    # --- Check for optional 'description' column ---
    has_description_col = 'col_desc' in processed_data
    if has_description_col and not isinstance(processed_data.get('col_desc'), list):
        logging.warning(f"{prefix} 'col_desc' column exists but is not a list. Will use None for description keys.")
        has_description_col = False
    elif not has_description_col:
        logging.info(f"{prefix} 'col_desc' column not found or is invalid. Will use None for description keys.")


    # Safely get lists and check types - Use get() with    # Get column data using new 'col_' keys
    po_list = processed_data.get('col_po', [])
    item_list = processed_data.get('col_item', [])
    sqft_list = processed_data.get('col_qty_sf', [])
    amount_list = processed_data.get('col_amount', [])
    description_list = processed_data.get('col_desc', []) if has_description_col else []

    # Use length of 'po' list as reference (handle potential missing 'po'?)
    num_rows = len(po_list) if po_list else 0
    if not po_list and (item_list or sqft_list or amount_list or description_list):
        # If PO is missing but others exist, try using another core list length
        core_lists = [l for l in [item_list, sqft_list, amount_list] if l]
        if core_lists:
            num_rows = len(core_lists[0])
            logging.warning(f"{prefix} 'col_po' list missing or empty, using length of another list ({num_rows}) as reference.")
        else:
            logging.warning(f"{prefix} Core lists ('col_po', 'col_item', 'col_qty_sf', 'col_amount') seem missing or empty. Cannot reliably determine row count.")
            num_rows = 0 # Cannot proceed safely

    logging.debug(f"{prefix} Determined number of rows: {num_rows}")

    # Check for length consistency across all *found* required lists + description if present
    all_lists = {'col_po': po_list, 'col_item': item_list, 'col_qty_sf': sqft_list, 'col_amount': amount_list}
    if has_description_col:
        all_lists['col_desc'] = description_list

    found_lists = {k: v for k, v in all_lists.items() if k in processed_data and isinstance(processed_data[k], list)}

    if not all(len(lst) == num_rows for lst in found_lists.values()):
        lengths = {k: len(v) for k, v in found_lists.items()}
        logging.error(f"{prefix} Data length mismatch for found columns! Ref:{num_rows}, Found:{lengths}. Aborting custom aggregation for this table.")
        return aggregated_results

    if num_rows == 0:
        logging.info(f"{prefix} No data rows found in this table. Global custom aggregation map remains unchanged.")
        return aggregated_results

    logging.info(f"{prefix} Processing {num_rows} rows from this table to update global CUSTOM aggregation (by PO/Item/Desc).")

    # --- Iterate and Aggregate ---
    rows_processed_this_table = 0
    successful_conversions_sqft = 0
    successful_conversions_amount = 0

    for i in range(num_rows):
        rows_processed_this_table += 1
        log_row_context = f"{prefix} Table Row index {i}"
        # logging.debug(f"{log_row_context} --- Processing ---") # Reduced verbosity

        # Get raw values (handle potential missing lists safely using get with default None)
        po_val = po_list[i] if i < len(po_list) else None
        item_val = item_list[i] if i < len(item_list) else None
        sqft_raw = sqft_list[i] if i < len(sqft_list) else None
        amount_raw = amount_list[i] if i < len(amount_list) else None
        desc_raw = description_list[i] if has_description_col and i < len(description_list) else None

        # logging.debug(f"{log_row_context}: Raw values - PO='{po_val}', Item='{item_val}', Desc='{desc_raw}', SQFT='{sqft_raw}', Amount='{amount_raw}'") # Reduced verbosity

        # Prepare the key components (Handle None, strip strings)
        po_key = str(po_val).strip() if isinstance(po_val, str) else po_val
        item_key = str(item_val).strip() if isinstance(item_val, str) else item_val
        description_key = str(desc_raw).strip() if isinstance(desc_raw, str) else desc_raw
        description_key = description_key if description_key else None # Ensure empty strings become None


        po_key = po_key if po_key is not None else "<MISSING_PO>"
        item_key = item_key if item_key is not None else "<MISSING_ITEM>"
        # Description key can be None

        # logging.debug(f"{log_row_context}: Key parts - PO Key='{po_key}', Item Key='{item_key}', Desc Key='{description_key}'") # Reduced verbosity

        # UPDATED Key: (PO, Item, None, Description) - Move description to index 3 to match standard
        key = (po_key, item_key, None, description_key)
        # logging.debug(f"{log_row_context}: Generated Key Tuple = {key}") # Reduced verbosity

        # Convert SQFT to Decimal for summation (default to 0 if fails/None)
        sqft_dec = _convert_to_decimal(sqft_raw, f"{log_row_context} SQFT")
        if sqft_dec is None:
             # logging.debug(f"{log_row_context}: SQFT value '{sqft_raw}' is None or failed conversion. Using 0.") # Reduced verbosity
             sqft_dec = decimal.Decimal(0)
        else:
             successful_conversions_sqft +=1

        # Convert Amount to Decimal for summation (default to 0 if fails/None)
        amount_dec = _convert_to_decimal(amount_raw, f"{log_row_context} Amount")
        if amount_dec is None:
            # logging.debug(f"{log_row_context}: Amount value '{amount_raw}' is None or failed conversion. Using 0.") # Reduced verbosity
            amount_dec = decimal.Decimal(0)
        else:
            successful_conversions_amount +=1

        # logging.debug(f"{log_row_context}: Converted values - SQFT='{sqft_dec}', Amount='{amount_dec}'") # Reduced verbosity

        # --- Add to the global aggregate sums (SQFT & Amount) ---
        current_sums = aggregated_results.get(key, {'sqft_sum': decimal.Decimal(0), 'amount_sum': decimal.Decimal(0)})

        # logging.debug(f"{log_row_context}: Sums for key {key} BEFORE add = {current_sums}") # Reduced verbosity

        # Update the sums
        current_sums['sqft_sum'] += sqft_dec
        current_sums['amount_sum'] += amount_dec

        # Store the updated dictionary back into the global map
        aggregated_results[key] = current_sums


    # --- Log summary for this table's contribution ---
    logging.info(f"{prefix} Finished processing {rows_processed_this_table} rows for this table.")
    logging.info(f"{prefix} SQFT values successfully converted/defaulted for {successful_conversions_sqft} rows.")
    logging.info(f"{prefix} Amount values successfully converted/defaulted for {successful_conversions_amount} rows.")
    logging.info(f"{prefix} Global custom aggregation map now contains {len(aggregated_results)} unique (PO, Item, None, Description) keys.")

    return aggregated_results


def calculate_leather_summary(processed_data: Dict[str, List[Any]]) -> Dict[str, Any]:
    """
    Calculates the leather summary (PCS, SQFT, Net, Gross, Pallet Count) per leather type.
    Iterates through rows to sum values for each leather type found in 'description' or 'desc'.
    BUFFALO = rows containing "BUFFALO" in description
    COW = rows NOT containing "BUFFALO" (all other leather)
    """
    # Initialize summary structure with default 0s
    summary = {
        'BUFFALO': {'pcs': 0, 'sqft': decimal.Decimal(0), 'net': decimal.Decimal(0), 'gross': decimal.Decimal(0), 'pallet_count': 0},
        'COW': {'pcs': 0, 'sqft': decimal.Decimal(0), 'net': decimal.Decimal(0), 'gross': decimal.Decimal(0), 'pallet_count': 0}
    }

    if not isinstance(processed_data, dict):
        return summary

    # Get columns - check both 'description' and 'desc' field names
    description_list = processed_data.get('col_desc', [])
    
    # robustly determine num_rows
    lists_to_check = [processed_data.get(col) for col in ['col_po', 'col_item', 'col_qty_sf', 'col_amount', 'col_net', 'col_gross', 'col_quantity', 'col_qty_pcs'] if processed_data.get(col)]
    if description_list:
        num_rows = len(description_list)
    elif lists_to_check:
        num_rows = len(lists_to_check[0])
    else:
        num_rows = 0

    # Get other columns safely
    pcs_list = processed_data.get('col_qty_pcs', [])
    sqft_list = processed_data.get('col_qty_sf', [])
    net_list = processed_data.get('col_net', [])
    gross_list = processed_data.get('col_gross', [])
    pallet_count_list = processed_data.get('col_pallet_count', [])

    for i in range(num_rows):
        desc = str(description_list[i]).upper() if i < len(description_list) and description_list[i] else ""
        
        # BUFFALO = contains "BUFFALO", COW = everything else (non-buffalo leather)
        leather_type = None
        if "BUFFALO" in desc:
            leather_type = 'BUFFALO'
        else:
            leather_type = 'COW'  # All non-buffalo is considered COW/regular leather
        
        # Sum PCS
        try:
            val = pcs_list[i] if i < len(pcs_list) else 0
            summary[leather_type]['pcs'] += int(float(val)) if val else 0
        except (ValueError, TypeError): pass

        # Sum SQFT
        try:
            val = sqft_list[i] if i < len(sqft_list) else 0
            summary[leather_type]['sqft'] += _convert_to_decimal(val) if val else 0
        except (ValueError, TypeError): pass

        # Sum Net
        try:
            val = net_list[i] if i < len(net_list) else 0
            summary[leather_type]['net'] += _convert_to_decimal(val) if val else 0
        except (ValueError, TypeError): pass

        # Sum Gross
        try:
            val = gross_list[i] if i < len(gross_list) else 0
            summary[leather_type]['gross'] += _convert_to_decimal(val) if val else 0
        except (ValueError, TypeError): pass

        # Sum Pallet Count (if available per row)
        try:
            val = pallet_count_list[i] if i < len(pallet_count_list) else 0
            summary[leather_type]['pallet_count'] += int(float(val)) if val else 0
        except (ValueError, TypeError): pass

    return summary


def aggregate_per_po_with_pallets(processed_data: Dict[str, List[Any]]) -> List[Dict[str, Any]]:
    """
    Aggregates data by PO and Price, summing sqft, amount, and pallet_count.
    Groups rows that have the same PO and unit_price together.
    
    Returns a list of aggregated records with:
    - po: The PO number
    - item: Combined unique items (comma-separated)
    - desc: Combined unique descriptions (comma-separated)
    - unit_price: The unit price
    - sqft: Total sqft for this PO+price combination
    - amount: Total amount for this PO+price combination
    - pallet_count: Total pallets for this PO+price combination
    - net: Total net weight for this PO+price combination
    - gross: Total gross weight for this PO+price combination
    - cbm: Total cbm for this PO+price combination
    """
    if not isinstance(processed_data, dict):
        return []
    
    # Get column data - check both field name variants
    po_list = processed_data.get('col_po', [])
    item_list = processed_data.get('col_item', [])
    desc_list = processed_data.get('col_desc', [])
    price_list = processed_data.get('col_unit_price', [])
    sqft_list = processed_data.get('col_qty_sf', [])
    amount_list = processed_data.get('col_amount', [])
    pallet_list = processed_data.get('col_pallet_count', [])
    net_list = processed_data.get('col_net', [])
    gross_list = processed_data.get('col_gross', [])
    cbm_list = processed_data.get('col_cbm', [])
    
    if not po_list:
        return []
    
    num_rows = len(po_list)
    
    # Aggregation map: (po, price) -> {items: set, descs: set, sqft: Decimal, amount: Decimal, pallets: int, net: Decimal, gross: Decimal, cbm: Decimal}
    aggregation_map = {}
    
    for i in range(num_rows):
        po = str(po_list[i]) if i < len(po_list) and po_list[i] else ""
        
        # Get price - try to convert to Decimal for consistent key
        try:
            price_val = price_list[i] if i < len(price_list) else 0
            price = _convert_to_decimal(price_val) if price_val else decimal.Decimal(0)
        except:
            price = decimal.Decimal(0)
        # logging.debug(f"{log_row_context} --- Processing ---") # Reduced verbosity

        # Get raw values (handle potential missing lists safely using get with default None)
        po_val = po_list[i] if i < len(po_list) else None
        item_val = item_list[i] if i < len(item_list) else None
        sqft_raw = sqft_list[i] if i < len(sqft_list) else None
        amount_raw = amount_list[i] if i < len(amount_list) else None
        desc_raw = description_list[i] if has_description_col and i < len(description_list) else None

        # logging.debug(f"{log_row_context}: Raw values - PO='{po_val}', Item='{item_val}', Desc='{desc_raw}', SQFT='{sqft_raw}', Amount='{amount_raw}'") # Reduced verbosity

        # Prepare the key components (Handle None, strip strings)
        po_key = str(po_val).strip() if isinstance(po_val, str) else po_val
        item_key = str(item_val).strip() if isinstance(item_val, str) else item_val
        description_key = str(desc_raw).strip() if isinstance(desc_raw, str) else desc_raw
        description_key = description_key if description_key else None # Ensure empty strings become None


        po_key = po_key if po_key is not None else "<MISSING_PO>"
        item_key = item_key if item_key is not None else "<MISSING_ITEM>"
        # Description key can be None

        # logging.debug(f"{log_row_context}: Key parts - PO Key='{po_key}', Item Key='{item_key}', Desc Key='{description_key}'") # Reduced verbosity

        # UPDATED Key: (PO, Item, None, Description) - Move description to index 3 to match standard
        key = (po_key, item_key, None, description_key)
        # logging.debug(f"{log_row_context}: Generated Key Tuple = {key}") # Reduced verbosity

        # Convert SQFT to Decimal for summation (default to 0 if fails/None)
        sqft_dec = _convert_to_decimal(sqft_raw, f"{log_row_context} SQFT")
        if sqft_dec is None:
             # logging.debug(f"{log_row_context}: SQFT value '{sqft_raw}' is None or failed conversion. Using 0.") # Reduced verbosity
             sqft_dec = decimal.Decimal(0)
        else:
             successful_conversions_sqft +=1

        # Convert Amount to Decimal for summation (default to 0 if fails/None)
        amount_dec = _convert_to_decimal(amount_raw, f"{log_row_context} Amount")
        if amount_dec is None:
            # logging.debug(f"{log_row_context}: Amount value '{amount_raw}' is None or failed conversion. Using 0.") # Reduced verbosity
            amount_dec = decimal.Decimal(0)
        else:
            successful_conversions_amount +=1

        # logging.debug(f"{log_row_context}: Converted values - SQFT='{sqft_dec}', Amount='{amount_dec}'") # Reduced verbosity

        # --- Add to the global aggregate sums (SQFT & Amount) ---
        current_sums = aggregated_results.get(key, {'sqft_sum': decimal.Decimal(0), 'amount_sum': decimal.Decimal(0)})

        # logging.debug(f"{log_row_context}: Sums for key {key} BEFORE add = {current_sums}") # Reduced verbosity

        # Update the sums
        current_sums['sqft_sum'] += sqft_dec
        current_sums['amount_sum'] += amount_dec

        # Store the updated dictionary back into the global map
        aggregated_results[key] = current_sums


    # --- Log summary for this table's contribution ---
    logging.info(f"{prefix} Finished processing {rows_processed_this_table} rows for this table.")
    logging.info(f"{prefix} SQFT values successfully converted/defaulted for {successful_conversions_sqft} rows.")
    logging.info(f"{prefix} Amount values successfully converted/defaulted for {successful_conversions_amount} rows.")
    logging.info(f"{prefix} Global custom aggregation map now contains {len(aggregated_results)} unique (PO, Item, None, Description) keys.")

    return aggregated_results


def calculate_leather_summary(processed_data: Dict[str, List[Any]]) -> Dict[str, Any]:
    """
    Calculates the leather summary (PCS, SQFT, Net, Gross, Pallet Count) per leather type.
    Iterates through rows to sum values for each leather type found in 'description' or 'desc'.
    BUFFALO = rows containing "BUFFALO" in description
    COW = rows NOT containing "BUFFALO" (all other leather)
    """
    # Initialize summary structure with default 0s
    summary = {
        'BUFFALO': {'col_qty_pcs': 0, 'col_qty_sf': decimal.Decimal(0), 'col_net': decimal.Decimal(0), 'col_gross': decimal.Decimal(0), 'col_pallet_count': 0},
        'COW': {'col_qty_pcs': 0, 'col_qty_sf': decimal.Decimal(0), 'col_net': decimal.Decimal(0), 'col_gross': decimal.Decimal(0), 'col_pallet_count': 0}
    }

    if not isinstance(processed_data, dict):
        return summary

    # Get columns - check both 'description' and 'desc' field names
    description_list = processed_data.get('col_desc', [])
    
    # robustly determine num_rows
    lists_to_check = [processed_data.get(col) for col in ['col_po', 'col_item', 'col_qty_sf', 'col_amount', 'col_net', 'col_gross', 'col_quantity', 'col_qty_pcs'] if processed_data.get(col)]
    if description_list:
        num_rows = len(description_list)
    elif lists_to_check:
        num_rows = len(lists_to_check[0])
    else:
        num_rows = 0

    # Get other columns safely
    pcs_list = processed_data.get('col_qty_pcs', [])
    sqft_list = processed_data.get('col_qty_sf', [])
    net_list = processed_data.get('col_net', [])
    gross_list = processed_data.get('col_gross', [])
    pallet_count_list = processed_data.get('col_pallet_count', [])

    for i in range(num_rows):
        desc = str(description_list[i]).upper() if i < len(description_list) and description_list[i] else ""
        
        # BUFFALO = contains "BUFFALO", COW = everything else (non-buffalo leather)
        leather_type = None
        if "BUFFALO" in desc:
            leather_type = 'BUFFALO'
        else:
            leather_type = 'COW'  # All non-buffalo is considered COW/regular leather
        
        # Sum PCS
        try:
            val = pcs_list[i] if i < len(pcs_list) else 0
            summary[leather_type]['col_qty_pcs'] += int(float(val)) if val else 0
        except (ValueError, TypeError): pass

        # Sum SQFT
        try:
            val = sqft_list[i] if i < len(sqft_list) else 0
            summary[leather_type]['col_qty_sf'] += _convert_to_decimal(val) if val else 0
        except (ValueError, TypeError): pass

        # Sum Net
        try:
            val = net_list[i] if i < len(net_list) else 0
            summary[leather_type]['col_net'] += _convert_to_decimal(val) if val else 0
        except (ValueError, TypeError): pass

        # Sum Gross
        try:
            val = gross_list[i] if i < len(gross_list) else 0
            summary[leather_type]['col_gross'] += _convert_to_decimal(val) if val else 0
        except (ValueError, TypeError): pass

        # Sum Pallet Count (if available per row)
        try:
            val = pallet_count_list[i] if i < len(pallet_count_list) else 0
            summary[leather_type]['col_pallet_count'] += int(float(val)) if val else 0
        except (ValueError, TypeError): pass

    return summary


def aggregate_per_po_with_pallets(processed_data: Dict[str, List[Any]]) -> List[Dict[str, Any]]:
    """
    Aggregates data by PO and Price, summing sqft, amount, and pallet_count.
    Groups rows that have the same PO and unit_price together.
    
    Returns a list of aggregated records with:
    - po: The PO number
    - item: Combined unique items (comma-separated)
    - desc: Combined unique descriptions (comma-separated)
    - unit_price: The unit price
    - sqft: Total sqft for this PO+price combination
    - amount: Total amount for this PO+price combination
    - pallet_count: Total pallets for this PO+price combination
    - net: Total net weight for this PO+price combination
    - gross: Total gross weight for this PO+price combination
    - cbm: Total cbm for this PO+price combination
    """
    if not isinstance(processed_data, dict):
        return []
    
    # Get column data - check both field name variants
    po_list = processed_data.get('col_po', [])
    item_list = processed_data.get('col_item', [])
    desc_list = processed_data.get('col_desc', [])
    price_list = processed_data.get('col_unit_price', [])
    sqft_list = processed_data.get('col_qty_sf', [])
    amount_list = processed_data.get('col_amount', [])
    pallet_list = processed_data.get('col_pallet_count', [])
    net_list = processed_data.get('col_net', [])
    gross_list = processed_data.get('col_gross', [])
    cbm_list = processed_data.get('col_cbm', [])
    
    if not po_list:
        return []
    
    num_rows = len(po_list)
    
    # Aggregation map: (po, price) -> {items: set, descs: set, sqft: Decimal, amount: Decimal, pallets: int, net: Decimal, gross: Decimal, cbm: Decimal}
    aggregation_map = {}
    
    for i in range(num_rows):
        po = str(po_list[i]) if i < len(po_list) and po_list[i] else ""
        
        # Get price - try to convert to Decimal for consistent key
        try:
            price_val = price_list[i] if i < len(price_list) else 0
            price = _convert_to_decimal(price_val) if price_val else decimal.Decimal(0)
        except:
            price = decimal.Decimal(0)
        
        key = (po, price)
        
        if key not in aggregation_map:
            aggregation_map[key] = {
                'items': set(),
                'descs': set(),
                'col_qty_sf': decimal.Decimal(0),
                'col_amount': decimal.Decimal(0),
                'col_pallet_count': 0,
                'col_net': decimal.Decimal(0),
                'col_gross': decimal.Decimal(0),
                'col_cbm': decimal.Decimal(0)
            }
        
        # Collect unique items
        if i < len(item_list) and item_list[i]:
            aggregation_map[key]['items'].add(str(item_list[i]))
        
        # Collect unique descriptions
        if i < len(desc_list) and desc_list[i]:
            aggregation_map[key]['descs'].add(str(desc_list[i]))
        
        # Sum sqft
        try:
            val = sqft_list[i] if i < len(sqft_list) else 0
            aggregation_map[key]['col_qty_sf'] += _convert_to_decimal(val) if val else decimal.Decimal(0)
        except (ValueError, TypeError):
            pass
        
        # Sum amount
        try:
            val = amount_list[i] if i < len(amount_list) else 0
            aggregation_map[key]['col_amount'] += _convert_to_decimal(val) if val else decimal.Decimal(0)
        except (ValueError, TypeError):
            pass
        
        # Sum pallet_count
        try:
            val = pallet_list[i] if i < len(pallet_list) else 0
            aggregation_map[key]['col_pallet_count'] += int(float(val)) if val else 0
        except (ValueError, TypeError):
            pass
        
        # Sum net
        try:
            val = net_list[i] if i < len(net_list) else 0
            aggregation_map[key]['col_net'] += _convert_to_decimal(val) if val else decimal.Decimal(0)
        except (ValueError, TypeError):
            pass
        
        # Sum gross
        try:
            val = gross_list[i] if i < len(gross_list) else 0
            aggregation_map[key]['col_gross'] += _convert_to_decimal(val) if val else decimal.Decimal(0)
        except (ValueError, TypeError):
            pass
        
        # Sum cbm
        try:
            val = cbm_list[i] if i < len(cbm_list) else 0
            aggregation_map[key]['col_cbm'] += _convert_to_decimal(val) if val else decimal.Decimal(0)
        except (ValueError, TypeError):
            pass
    
    # Convert to list of dicts
    result = []
    for (po, price), data in aggregation_map.items():
        result.append({
            'col_po': po,
            'col_item': ', '.join(sorted(data['items'])),
            'col_desc': ', '.join(sorted(data['descs'])),
            'col_unit_price': price,
            'col_qty_sf': data['col_qty_sf'],
            'col_amount': data['col_amount'],
            'col_pallet_count': data['col_pallet_count'],
            'col_net': data['col_net'],
            'col_gross': data['col_gross'],
            'col_cbm': data['col_cbm']
        })
    
    # Sort by PO for consistent output
    result.sort(key=lambda x: x['col_po'])
    
    return result


def calculate_weight_summary(processed_data: Dict[str, List[Any]]) -> Dict[str, decimal.Decimal]:
    """
    Calculates the weight summary (Net Weight and Gross Weight).
    
    Args:
        processed_data: Dictionary representing the data of the current table.
        
    Returns:
        Dictionary containing 'net' and 'gross' weights.
    """
    summary = {'col_net': decimal.Decimal(0), 'col_gross': decimal.Decimal(0)}
    
    if not isinstance(processed_data, dict):
        return summary
        
    # Sum Net Weight
    net_values = processed_data.get('col_net', [])
    for val in net_values:
        dec_val = _convert_to_decimal(val)
        if dec_val is not None:
            summary['col_net'] += dec_val
            
    # Sum Gross Weight
    gross_values = processed_data.get('col_gross', [])
    for val in gross_values:
        dec_val = _convert_to_decimal(val)
        if dec_val is not None:
            summary['col_gross'] += dec_val
            
    return summary

def calculate_pallet_summary(processed_data: Dict[str, List[Any]]) -> int:
    """
    Calculates the total pallet count for the table.
    
    Args:
        processed_data: Dictionary representing the data of the current table.
        
    Returns:
        Total pallet count as integer.
    """
    total_pallets = 0
    
    if not isinstance(processed_data, dict):
        return 0
        
    pallet_values = processed_data.get('col_pallet_count', [])
    for val in pallet_values:
        if val is not None:
            try:
                total_pallets += int(float(val))
            except (ValueError, TypeError):
                pass
                
    return total_pallets

def calculate_footer_totals(processed_data: Dict[str, List[Any]]) -> Dict[str, Any]:
    """
    Calculates totals for footer fields based on processed data.
    Returns a dictionary with keys matching the expected footer data structure.
    """
    totals = {
        "col_qty_pcs": 0,
        "col_qty_sf": decimal.Decimal(0),
        "col_net": decimal.Decimal(0),
        "col_gross": decimal.Decimal(0),
        "col_cbm": decimal.Decimal(0),
        "col_amount": decimal.Decimal(0),
        "col_pallet_count": 0
    }
    
    if not processed_data:
        return totals

    # Helper to safely add
    def safe_add_decimal(key, value):
        if value is not None:
            try:
                # Handle strings with commas if necessary, though data parser should have cleaned it
                val_str = str(value).replace(',', '')
                totals[key] += decimal.Decimal(val_str)
            except (decimal.InvalidOperation, ValueError, TypeError):
                pass

    def safe_add_int(key, value):
        if value is not None:
            try:
                val_str = str(value).replace(',', '')
                totals[key] += int(float(val_str))
            except (ValueError, TypeError):
                pass

    # Sum each list independently
    for val in processed_data.get('col_qty_pcs', []):
        safe_add_int('col_qty_pcs', val)
        
    for val in processed_data.get('col_qty_sf', []):
        safe_add_decimal('col_qty_sf', val)
        
    for val in processed_data.get('col_net', []):
        safe_add_decimal('col_net', val)
        
    for val in processed_data.get('col_gross', []):
        safe_add_decimal('col_gross', val)
        
    for val in processed_data.get('col_cbm', []):
        safe_add_decimal('col_cbm', val)
        
    for val in processed_data.get('col_amount', []):
        safe_add_decimal('col_amount', val)
        
    for val in processed_data.get('col_pallet_count', []):
        safe_add_int('col_pallet_count', val)

    return totals


def format_aggregation_as_list(
    aggregation_map: Dict[Tuple, Dict[str, decimal.Decimal]],
    mode: str = 'standard'
) -> List[Dict[str, Any]]:
    """
    Converts the internal tuple-keyed aggregation map into a clean list of dictionaries
    suitable for JSON output. Removes tuple keys and uses 'col_' prefixed keys for all fields.

    Args:
        aggregation_map: The dictionary holding the aggregation results.
        mode: 'standard' or 'custom' indicating the key structure.

    Returns:
        A list of dictionaries, each representing an aggregated row.
    """
    flattened_list = []
    
    for key_tuple, values in aggregation_map.items():
        row_dict = {}
        
        # Extract values from the tuple key based on mode
        if mode == 'standard':
            # Key: (PO, Item, Price, Desc)
            # Mapping: Index 0->col_po, 1->col_item, 2->col_unit_price, 3->col_desc
            if len(key_tuple) >= 4:
                row_dict['col_po'] = str(key_tuple[0]) if key_tuple[0] is not None else ""
                row_dict['col_item'] = str(key_tuple[1]) if key_tuple[1] is not None else ""
                row_dict['col_unit_price'] = str(key_tuple[2]) if key_tuple[2] is not None else ""
                row_dict['col_desc'] = str(key_tuple[3]) if key_tuple[3] is not None else ""
            else:
                # Fallback for unexpected key length
                row_dict['col_po'] = str(key_tuple[0]) if len(key_tuple) > 0 else ""
                row_dict['col_item'] = str(key_tuple[1]) if len(key_tuple) > 1 else ""
                
        elif mode == 'custom':
            # Key: (PO, Item, None, Desc)
            # Mapping: Index 0->col_po, 1->col_item, 3->col_desc (Index 2 is None/Ignored)
            if len(key_tuple) >= 4:
                row_dict['col_po'] = str(key_tuple[0]) if key_tuple[0] is not None else ""
                row_dict['col_item'] = str(key_tuple[1]) if key_tuple[1] is not None else ""
                row_dict['col_desc'] = str(key_tuple[3]) if key_tuple[3] is not None else ""
            else:
                 # Fallback
                row_dict['col_po'] = str(key_tuple[0]) if len(key_tuple) > 0 else ""
                row_dict['col_item'] = str(key_tuple[1]) if len(key_tuple) > 1 else ""

        # Extract summed values (already using col_ keys internally, but safe get)
        # Handle both new col_ keys and potential legacy keys if any remain
        row_dict['col_qty_sf'] = values.get('col_qty_sf', values.get('sqft_sum', decimal.Decimal(0)))
        row_dict['col_amount'] = values.get('col_amount', values.get('amount_sum', decimal.Decimal(0)))
        
        flattened_list.append(row_dict)
        
    return flattened_list