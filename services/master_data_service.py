
import pandas as pd
from pathlib import Path
from typing import Set, Dict, List, Optional, Tuple
from core.models import VerificationStatus
import re

class MasterDataService:
    def __init__(self, master_path: Path):
        self.master_path = master_path
        self.df = None
        self.col_map = {}
        
    def load(self) -> bool:
        """Loads the Master List into memory."""
        if not self.master_path.exists():
            print(f"Master file not found: {self.master_path}")
            return False
            
        try:
            if self.master_path.suffix.lower() == '.csv':
                self.df = pd.read_csv(self.master_path)
            else:
                self.df = pd.read_excel(self.master_path)
            self._map_columns()
            return True
        except Exception as e:
            print(f"Error loading master file: {e}")
            return False
            
    def _map_columns(self):
        """Identifies standard columns in the loaded DF using mapping_config.json."""
        if self.df is None: return
        self.col_map = {}
        
        from core.config import load_mapping_config
        mapping_dict = load_mapping_config()
        
        # Known col_ids for direct column name matching
        known_col_ids = ['invoice_id', 'col_qty_sf', 'col_amount', 'col_pallet_count',
                         'col_qty_pcs', 'col_net', 'col_gross', 'col_cbm']
        
        for c in self.df.columns:
            cl = c.lower().strip()
            
            # 1. Config Match (from mapping_config.json)
            if cl in mapping_dict:
                col_id = mapping_dict[cl]
                
                # Special Case: ID
                if col_id == 'col_inv_no':
                     self.col_map['invoice_id'] = c
                else:
                     self.col_map[col_id] = c
            
            # 2. Direct col_id match (column is already named col_cbm, col_net, etc.)
            elif cl in known_col_ids:
                self.col_map[cl] = c
            
            # 3. Fallback for ID (Legacy support if config is missing 'Invoice No')
            elif ('invoice' in cl or 'id' in cl) and 'diff' not in cl and 'verify' not in cl:
                 if 'invoice_id' not in self.col_map:
                     self.col_map['invoice_id'] = c

    def get_known_ids(self) -> Tuple[Set[str], Set[str]]:
        """Returns (all_ids, verified_ids)."""
        if self.df is None: 
            if not self.load(): return set(), set()
            
        ids = set()
        verified_ids = set()
        
        id_col = self.col_map.get('invoice_id')
        if not id_col: return set(), set()
        
        ids = set(self.df[id_col].dropna().astype(str).str.strip().unique())
        
        # Check Verify State
        verify_col = None
        for c in self.df.columns:
            if 'verify state' in c.lower() or 'verified' in c.lower():
                verify_col = c
                break
                
        if verify_col:
            v_mask = self.df[verify_col].astype(str).str.lower() == 'true'
            verified_ids = set(self.df[v_mask][id_col].dropna().astype(str).str.strip().unique())
            
        return ids, verified_ids

    def verify_and_update(self, extracted_data: list) -> None:
        """Verifies extracted data against master list and updates it."""
        if self.df is None: self.load()
        if self.df is None: return
        
        # Build master lookup by ID for fast access
        id_col = self.col_map.get('invoice_id')
        if not id_col: return
        
        # Index master rows by ID (keep all rows, not just first match)
        master_by_id = {}
        for index, row in self.df.iterrows():
            m_id = str(row.get(id_col)).strip()
            if m_id not in master_by_id:
                master_by_id[m_id] = []
            master_by_id[m_id].append((index, row))
            
        # Diffs Columns: target IS the key now (col_xxx)
        diff_cols_map = {
            'DIFF_PALLET': 'col_pallet_count', 
            'DIFF_SQFT': 'col_qty_sf', 
            'DIFF_AMOUNT': 'col_amount',
            'DIFF_PCS': 'col_qty_pcs', 
            'DIFF_NET': 'col_net', 
            'DIFF_GROSS': 'col_gross', 
            'DIFF_CBM': 'col_cbm'
        }

        for col in ['VERIFY STATE'] + list(diff_cols_map.keys()):
            if col not in self.df.columns: self.df[col] = None
        
        # Iterate over EACH extracted item (not master rows)
        for item in extracted_data:
            inv_id = item.get('invoice_id')
            if not inv_id: continue
            
            # Find matching master row
            master_rows = master_by_id.get(str(inv_id).strip(), [])
            if not master_rows:
                item['status'] = 'Missing from Master'
                continue
            
            # Use first matching master row for comparison
            master_index, master_row = master_rows[0]
            
            all_match = True
            diffs = {}
            
            # Helper - returns None for empty/NaN values, float otherwise
            def get_num(v):
                try:
                    if v is None or (isinstance(v, str) and v.strip() == ''):
                        return None
                    if isinstance(v, float) and (v != v):  # NaN check: NaN != NaN
                        return None
                    if isinstance(v, (int, float)): 
                        return float(v)
                    # Use regex_utils for string extraction
                    from core.regex_utils import regex_extract_number
                    result = regex_extract_number(str(v), default=None)
                    return result
                except: 
                    return None

            
            # Check List (Standardized Keys)
            checks = [
                'col_qty_sf', 'col_amount', 'col_pallet_count',
                'col_qty_pcs', 'col_net', 'col_gross', 'col_cbm'
            ]
            
            failures_by_sheet = {}
            
            # Get sheets data for per-sheet verification
            sheets_data = item.get('sheets', {})
            sheet_names = ['Invoice', 'PackingList', 'Contract']
            all_partition_details = []
            
            for key in checks:
                master_col = self.col_map.get(key)
                if not master_col: continue
                
                master_value = get_num(master_row.get(master_col))
                master_has_no_data = master_value is None
                
                # Check EACH sheet's value for this key against Master
                for sheet_name in sheet_names:
                    sheet_values = sheets_data.get(sheet_name, {})
                    
                    # STRICT: Always check if the sheet has this value (removed target_inspect_col gate)
                    if key not in sheet_values:
                        continue  # This sheet doesn't have this field value
                    
                    sheet_value = get_num(sheet_values.get(key))
                    if sheet_value is None or sheet_value == 0.0:
                        continue  # No meaningful value from sheet
                    
                    readable_key = key.replace('col_', '').replace('_', ' ').upper()
                    
                    # Get source file info for this item
                    # Use simple sheet name for cleaner display
                    sheet_source = sheet_name
                    
                    if master_has_no_data:
                        # Master has no value - if sheet has significant value, flag mismatch
                        if sheet_value > 1.0:
                            all_match = False
                            msg = f"{readable_key} Fail ({sheet_value} vs Master N/A)"
                            if sheet_source not in failures_by_sheet:
                                failures_by_sheet[sheet_source] = []
                            failures_by_sheet[sheet_source].append(msg)
                    else:
                        diff = sheet_value - master_value
                        
                        if abs(diff) > 0.01:
                            all_match = False
                            msg = f"{readable_key} Fail ({sheet_value} vs Master {master_value})"
                            if sheet_source not in failures_by_sheet:
                                failures_by_sheet[sheet_source] = []
                            failures_by_sheet[sheet_source].append(msg)
                
                # Report Partition Details (New Feature)
                partition_entries = []
                for s_name in sheet_names:
                    s_vals = sheets_data.get(s_name, {})
                    if key in s_vals:
                         val = get_num(s_vals[key])
                         if val is not None:
                             partition_entries.append(f"{s_name}={val}")
                
                if partition_entries:
                    partition_str = f"{key}: {', '.join(partition_entries)}"
                    # Append later to detection_summary
                    all_partition_details.append(partition_str)

                # Calculate diff using first sheet's value for master DF update
                # (Per-sheet check already validates, this is just for DIFF_* column display)
                extracted_value = get_num(item.get(key))
                if master_has_no_data:
                    diff = extracted_value if extracted_value is not None else 0.0
                else:
                    diff = (extracted_value if extracted_value is not None else 0.0) - master_value
                
                # Round diff to avoid floating point precision issues
                diff = round(diff, 7)
                    
                # Save diff to item
                for d_col, d_target_key in diff_cols_map.items():
                    if d_target_key == key:

                        diffs[d_col] = diff
            
            # --- NEW REPORT BUILDER (Sheet-Centric) ---
            report_lines = []
            source_file = item.get('file_name', 'Unknown')
            # report_lines.append(f"Source: {source_file}") # UI has file name, maybe redundant? keep for copy-paste.
            
            # 1. Iterate Sheets
            for i, sheet_name in enumerate(sheet_names, 1):
                sheet_data = sheets_data.get(sheet_name, {})
                
                # Header: "1. INVOICE (Row 20)"
                header_info = ""
                detection = sheet_data.get('detection_info', {})
                if detection and detection.get('status') == 'ok':
                    header_info = f"(Row {detection.get('header_row')})"
                elif detection:
                    header_info = f"(WARNING: {detection.get('warning')})"
                elif not sheet_data:
                    header_info = "(Not Found)"
                    
                report_lines.append(f"\n{i}. {sheet_name.upper()} {header_info}")
                
                # Table Header
                # Use fixed width formatting
                # Field (10) | Current (10) | Master (10) | Variance (10)
                report_lines.append(f"{'Field':<10} {'Current':<10} {'Master':<10} {'Variance':<10}")
                
                # Table Rows
                any_rows = False
                for key in checks:
                    # Check if this sheet has this key AND it has a value
                    # STRICT: We only report if the sheet actually extracted something for this column
                    val = get_num(sheet_data.get(key))
                    if val is None: continue
                    
                    any_rows = True
                    
                    # Get Master Val
                    master_col = self.col_map.get(key)
                    m_val = get_num(master_row.get(master_col)) if master_col else None
                    
                    # Calc Diff & formatting
                    diff_str = "N/A"
                    if m_val is not None:
                        diff = val - m_val
                        diff_str = f"{diff:+.2f}"
                        if diff == 0: diff_str = "-"
                    
                    m_str = str(m_val) if m_val is not None else "-"
                    v_str = str(val)
                    
                    # Name formatting: col_qty_sf -> Qty SF
                    name = key.replace('col_', '').replace('_', ' ').title()
                    name = name.replace('Qty Sf', 'Qty SF').replace('Qty Pcs', 'Qty PCS').replace('Cbm', 'CBM')
                    # Shorten for table
                    name = name.replace('Weight', 'Wgt').replace('Pallet Count', 'Pallets')
                    
                    report_lines.append(f"{name:<10} {v_str:<10} {m_str:<10} {diff_str:<10}")

                if not any_rows:
                    report_lines.append("  (No inspectable data detected)")

            # 2. Logic Checks
            report_lines.append("")
            # Net vs Gross (Packing List)
            pl_data = sheets_data.get('PackingList', {})
            # We need to be careful to extract them specifically from PL, 
            # extracted_value in item is aggregated (likely from PL but not guaranteed if we had multi-source)
            # Safe to take from item or PL sheet. PL sheet is explicit.
            net = get_num(pl_data.get('col_net'))
            gross = get_num(pl_data.get('col_gross'))

            if net is not None and gross is not None:
                if net > gross:
                    report_lines.append(f"[!] Critical Logic Error: Net Weight ({net}) > Gross Weight ({gross})")
            
            final_msg = "\n".join(report_lines)
            
            curr_details = item.get('verification_details', '')
            if curr_details:
                item['verification_details'] = curr_details + "\n\n" + final_msg
            else:
                item['verification_details'] = final_msg
            
            # Set status on item
            item['status'] = 'Verified' if all_match else 'Mismatch'
            
            # Update master DF (optional - updates first matching row)
            self.df.at[master_index, 'VERIFY STATE'] = all_match
            for k, v in diffs.items():
                self.df.at[master_index, k] = v
                
        # Save
        if self.master_path.suffix.lower() == '.csv':
            self.df.to_csv(self.master_path, index=False)
        else:
            self.df.to_excel(self.master_path, index=False)

    def parse_paste_data(self, clipboard_text: str) -> Tuple[List[List[str]], Dict[int, str], bool]:
        """
        Parses clipboard text using mapping_config.json.
        Returns:
        1. List of Rows
        2. Mapping {col_index: master_column_name}
        3. is_header_mode
        """
        rows = [line.split('\t') for line in clipboard_text.strip().split('\n')]
        if not rows: return [], {}, False
        
        # Load Config
        from core.config import load_mapping_config
        mapping_dict = load_mapping_config()
        
        # Check against Master Columns
        if self.df is None: self.load()
        master_cols = list(self.df.columns) if self.df is not None else []
        master_cols_lower = {c.lower(): c for c in master_cols}
        
        header_candidates = [str(cell).strip() for cell in rows[0]]
        match_count = 0
        mapping = {}
        
        for idx, header_candidate in enumerate(header_candidates):
            header_candidate_lower = header_candidate.lower()
            target_column = None
            
            # 1. Config Match (Prioritize col_id)
            if header_candidate_lower in mapping_dict:
                col_id = mapping_dict[header_candidate_lower]
                target_column = col_id 
                match_count += 1
            
            # 2. Existing Master Match (If not in config)
            elif header_candidate_lower in master_cols_lower:
                target_column = master_cols_lower[header_candidate_lower]
                match_count += 1
            
            if target_column:
                mapping[idx] = target_column
                
        is_header_mode = match_count > 0
        return rows, mapping, is_header_mode

    def apply_paste(self, rows: List[List[str]], mapping: Dict[int, str], is_header: bool):
        """Applies parsed paste data to the dataframe (Merge/Append)."""
        if self.df is None: return
        
        # Resolve Mapping to DF Columns
        # mapping contains col_ids (from parse_paste) or raw headers. 
        # Resolve to DF headers using col_map.
        final_mapping = {}
        for idx, col_ref in mapping.items():
            # Special case for Invoice ID which is stored as 'invoice_id' in col_map
            if col_ref == 'col_inv_no':
                 final_mapping[idx] = self.col_map.get('invoice_id', col_ref)
            else:
                 final_mapping[idx] = self.col_map.get(col_ref, col_ref)
        
        # Define Key Columns for Merge
        # Find the column that represents the Invoice ID
        merge_col_name = None
        
        # 1. Check if 'invoice_id' is mapped and present in paste
        if 'invoice_id' in self.col_map:
             real_id_col = self.col_map['invoice_id']
             if real_id_col in final_mapping.values():
                 merge_col_name = real_id_col
        
        # 2. Fallback
        if not merge_col_name:
             for c in final_mapping.values():
                 c_str = str(c).lower()
                 if 'id' in c_str or 'invoice' in c_str:
                     merge_col_name = c
                     break
                     
        start_idx = 1 if is_header else 0
        
        for i in range(start_idx, len(rows)):
            row_data = rows[i]
            
            # Construct row dict based on resolved mapping
            new_data = {}
            for col_idx, target_col in final_mapping.items():
                if col_idx < len(row_data):
                    new_data[target_col] = row_data[col_idx]
            
            if not new_data: continue
            
            # Merge or Append
            matched = False
            if merge_col_name and merge_col_name in new_data:
                # Try to find match
                key_val = str(new_data[merge_col_name]).strip()
                # Find index
                # This is slow for large DF, but acceptable for GUI paste
                matches = self.df[self.df[merge_col_name].astype(str).str.strip() == key_val].index
                if not matches.empty:
                    idx = matches[0]
                    # Update
                    for k, v in new_data.items():
                        self.df.at[idx, k] = v
                    matched = True
            
            if not matched:
                # Append
                # We need to ensure we use loc with new index
                new_idx = self.df.index.max() + 1 if not self.df.empty else 0
                # Assign per column to avoid NaN issues
                # Assign per column to avoid NaN issues
                for k, v in new_data.items():
                    # Attempt to convert to numeric if column is numeric
                    if k in self.df.columns and pd.api.types.is_numeric_dtype(self.df[k]):
                        try:
                            v = float(v)
                        except:
                            pass
                    self.df.at[new_idx, k] = v
        
        # Auto-save? Or let user click save?
        # Service should probably separate state from persistence, but for now we update memory.
        # GUI sets explicit "Save" button to flush to disk.
