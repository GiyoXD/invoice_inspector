import csv
from pathlib import Path
from typing import List, Dict, Any

def generate_report(results: List[Dict[str, Any]], output_folder: Path):
    """
    Generates a CSV report from verification results.
    """
    if not results:
        print("No results to report.")
        return

    output_path = output_folder / "verification_report.csv"
    
    fieldnames = ['Filename', 'Status', 'Details']
    
    try:
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            
            for res in results:
                # Flattern details for CSV readability
                details_text = " | ".join(res.get('details', []))
                
                writer.writerow({
                    'Filename': res.get('file_name'),
                    'Status': res.get('status'),
                    'Details': details_text
                })
                
        print(f"Report generated successfully: {output_path}")
        
    except Exception as e:
        print(f"Error generating report: {e}")
