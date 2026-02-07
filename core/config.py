
import json
from pathlib import Path

# Constants
MAPPING_CONFIG_PATH = Path("mapping_config.json")
REPORTS_DIR = Path("reports")

def load_mapping_config() -> dict:
    """Loads and normalizes the mapping configuration (Alias -> Canonical)."""
    if not MAPPING_CONFIG_PATH.exists():
        print(f"Warning: {MAPPING_CONFIG_PATH} not found. Using empty mapping.")
        return {}
    
    try:
        with open(MAPPING_CONFIG_PATH, 'r', encoding='utf-8') as f:
            config = json.load(f)
            
        # Normalize mappings: Lowercase key -> Col ID
        normalized = {}
        
        # Merge source 1: header_text_mappings
        if 'header_text_mappings' in config:
            for k, v in config['header_text_mappings'].get('mappings', {}).items():
                normalized[k.lower().strip()] = v
                
        # Merge source 2: shipping_list_header_map
        if 'shipping_list_header_map' in config:
            for k, v in config['shipping_list_header_map'].get('mappings', {}).items():
                normalized[k.lower().strip()] = v
                
        return normalized
    except Exception as e:
        print(f"Error loading mapping config: {e}")
        return {}
