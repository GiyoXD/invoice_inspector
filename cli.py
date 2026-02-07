
import argparse
import json
import sys
import contextlib
import pandas as pd
from pathlib import Path
from services.pipeline_service import PipelineService
from services.master_data_service import MasterDataService

def main():
    parser = argparse.ArgumentParser(description="Invoice Inspector CLI Adapter")
    
    # Modes: 'run_pipeline', 'parse_paste', 'verify'
    subparsers = parser.add_subparsers(dest='command', required=True)
    
    # Command: Run Inspection
    cmd_run = subparsers.add_parser('inspect', help='Run inspection pipeline')
    cmd_run.add_argument('--folder', required=True, help='Invoice Folder Path')
    cmd_run.add_argument('--master', required=True, help='Master List Path')
    
    # Command: Parse Paste
    cmd_paste = subparsers.add_parser('parse_paste', help='Parse clipboard data')
    cmd_paste.add_argument('--text', required=False, help='Clipboard text content')
    cmd_paste.add_argument('--file', required=False, help='Path to file containing clipboard content')
    cmd_paste.add_argument('--master', required=True, help='Master List Path')

    # Command: Load Master
    cmd_load = subparsers.add_parser('load_master', help='Load Master List Data')
    cmd_load.add_argument('--master', required=True, help='Master List Path')

    # Command: Save Master
    cmd_save = subparsers.add_parser('save_master', help='Save Master List Data')
    cmd_save.add_argument('--master', required=True, help='Master List Path')
    cmd_save.add_argument('--file', required=True, help='Path to JSON file containing new grid data')

    # Command: Merge Paste
    cmd_merge = subparsers.add_parser('merge_paste', help='Merge Paste Data into Master')
    cmd_merge.add_argument('--master', required=True, help='Master List Path')
    cmd_merge.add_argument('--file', required=True, help='Path to file containing clipboard content')
    cmd_merge.add_argument('--mapping', required=False, help='JSON file with user-edited column mappings')
    
    args = parser.parse_args()
    
    # Output structure: { "status": "ok", "data": ... } or { "status": "error", "message": ... }
    
    try:
        if args.command == 'inspect':
            with contextlib.redirect_stdout(sys.stderr):
                pipeline = PipelineService(args.folder, args.master)
                output = pipeline.run()
            
            # Output structure: { "status": "ok", "data": [...], "missing": [...] }
            # Pipeline now returns { "results": [...], "missing": [...] }
            print(json.dumps({
                "status": "ok", 
                "data": output['results'],
                "missing": output['missing']
            }, default=str))
            
        elif args.command == 'parse_paste':
            content = ""
            if args.file:
                with open(args.file, 'r', encoding='utf-8-sig') as f:
                    content = f.read()
            elif args.text:
                content = args.text
            
            svc = MasterDataService(Path(args.master))
            rows, mapping, is_header = svc.parse_paste_data(content)
            
            # Get available columns for dropdown
            master_cols = list(svc.df.columns) if svc.df is not None else []
            # Also add standard col_* options
            std_cols = ['invoice_id', 'col_qty_sf', 'col_amount', 'col_pallet_count', 
                        'col_qty_pcs', 'col_net', 'col_gross', 'col_cbm', '(skip)']
            available = list(set(master_cols + std_cols))
            
            # Serialize for generic UI
            print(json.dumps({
                "status": "ok",
                "data": {
                    "rows": rows,
                    "mapping": {str(k): v for k,v in mapping.items()}, # JS needs string keys
                    "is_header": is_header,
                    "available_columns": sorted(available)
                }
            }))
            
        elif args.command == 'load_master':
            svc = MasterDataService(Path(args.master))
            if svc.load():
                # Convert DF to list of lists (including header)
                # handle NaN
                df_to_send = svc.df.fillna('')
                
                # Apply Column Mapping (Rename Headers to col_id if mapped)
                # col_map is { 'col_qty_sf': 'Original Name' }
                # We need reverse map: { 'Original Name': 'col_qty_sf' }
                reverse_map = {v: k for k, v in svc.col_map.items()}
                
                new_columns = []
                for col in df_to_send.columns:
                    if col in reverse_map:
                        new_columns.append(reverse_map[col])
                    else:
                        new_columns.append(col)
                        
                columns = new_columns
                rows = df_to_send.values.tolist()
                
                print(json.dumps({
                    "status": "ok",
                    "data": {
                        "columns": columns,
                        "rows": rows
                    }
                }, default=str))
            else:
                 print(json.dumps({"status": "error", "message": "Could not load master file."}))

        elif args.command == 'save_master':
            try:
                if not args.file:
                     print(json.dumps({"status": "error", "message": "No data file provided for save."}))
                     return
                
                # Check extension
                is_csv = str(args.master).lower().endswith('.csv')

                # Read JSON input
                with open(args.file, 'r', encoding='utf-8-sig') as f:
                    input_data = json.load(f)
                
                # Reconstruct DF
                cols = input_data.get('columns', [])
                rows = input_data.get('rows', [])
                
                if not cols:
                     print(json.dumps({"status": "error", "message": "No columns provided."}))
                else:
                    new_df = pd.DataFrame(rows, columns=cols)
                    
                    if is_csv:
                        new_df.to_csv(args.master, index=False)
                    else:
                        new_df.to_excel(args.master, index=False)
                        
                    print(json.dumps({"status": "ok", "message": "Saved successfully."}))
                
            except Exception as e:
                 print(json.dumps({"status": "error", "message": str(e)}))

        elif args.command == 'merge_paste':
            try:
                if not args.file:
                     print(json.dumps({"status": "error", "message": "No data file provided for merge."}))
                     return
                
                # 1. Load Master
                svc = MasterDataService(Path(args.master))
                if not svc.load():
                     print(json.dumps({"status": "error", "message": "Could not load master file."}))
                     return
                
                # 2. Read Paste Content
                content = ""
                with open(args.file, 'r', encoding='utf-8-sig') as f:
                    content = f.read()

                # 3. Parse Paste OR use provided mapping
                rows, mapping, is_header = svc.parse_paste_data(content)
                
                # Override with user-edited mapping if provided
                if args.mapping:
                    with open(args.mapping, 'r', encoding='utf-8') as f:
                        user_mapping_data = json.load(f)
                        # user_mapping_data: {"mapping": {"0": "col_id", "1": "col_amount", ...}, "is_header": true}
                        mapping = {int(k): v for k, v in user_mapping_data.get('mapping', {}).items()}
                        is_header = user_mapping_data.get('is_header', is_header)
                
                # 4. Apply Paste (Merge)
                svc.apply_paste(rows, mapping, is_header)
                
                # 5. Save
                is_csv = str(args.master).lower().endswith('.csv')
                if is_csv:
                    svc.df.to_csv(args.master, index=False)
                else:
                    svc.df.to_excel(args.master, index=False)
                    
                print(json.dumps({"status": "ok", "message": "Merged and Saved successfully."}))
                
            except Exception as e:
                 import traceback
                 print(json.dumps({"status": "error", "message": str(e) + "\n" + traceback.format_exc()}))
            
    except Exception as e:
        print(json.dumps({"status": "error", "message": str(e)}))
        sys.exit(1)

if __name__ == "__main__":
    main()
