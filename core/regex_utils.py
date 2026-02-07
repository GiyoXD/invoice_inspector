"""
Regex utilities for searching and extracting data from Excel sheets.
"""

import re
from typing import Pattern


def regex_search_sheet(sheet, pattern: str | Pattern, max_row: int = 200, max_col: int = 30, 
                       case_insensitive: bool = True) -> list:
    """
    Searches an Excel sheet for cells matching a regex pattern.
    
    Args:
        sheet: openpyxl worksheet object
        pattern: Regex pattern (string or compiled Pattern)
        max_row: Maximum row to search (default 200)
        max_col: Maximum column to search (default 30)
        case_insensitive: If True, search is case-insensitive (default True)
        
    Returns:
        List of dicts: [{'row': int, 'col': int, 'value': str, 'match': str, 'groups': tuple}, ...]
    """
    results = []
    
    # Compile pattern if string
    if isinstance(pattern, str):
        flags = re.IGNORECASE if case_insensitive else 0
        pattern = re.compile(pattern, flags)
    
    for row_idx in range(1, min(max_row + 1, (sheet.max_row or max_row) + 1)):
        for col_idx in range(1, min(max_col + 1, (sheet.max_column or max_col) + 1)):
            cell = sheet.cell(row=row_idx, column=col_idx)
            if cell.value is None:
                continue
            
            cell_str = str(cell.value)
            match = pattern.search(cell_str)
            
            if match:
                results.append({
                    'row': row_idx,
                    'col': col_idx,
                    'value': cell_str,
                    'match': match.group(0),
                    'groups': match.groups() if match.groups() else None
                })
    
    return results


def regex_search_first(sheet, pattern: str | Pattern, max_row: int = 200, max_col: int = 30,
                       case_insensitive: bool = True) -> dict | None:
    """
    Searches an Excel sheet and returns the FIRST cell matching a regex pattern.
    
    Args:
        sheet: openpyxl worksheet object
        pattern: Regex pattern (string or compiled Pattern)
        max_row: Maximum row to search
        max_col: Maximum column to search
        case_insensitive: If True, search is case-insensitive
        
    Returns:
        Dict with cell info, or None if not found
    """
    results = regex_search_sheet(sheet, pattern, max_row, max_col, case_insensitive)
    return results[0] if results else None


def regex_extract(text: str, pattern: str | Pattern, group: int = 1, 
                  default=None, case_insensitive: bool = True):
    """
    Extracts a value from text using regex.
    
    Args:
        text: Text to search
        pattern: Regex pattern with capture group(s)
        group: Which capture group to return (default 1)
        default: Default value if no match
        case_insensitive: If True, search is case-insensitive
        
    Returns:
        Matched group value, or default if not found
    """
    if text is None:
        return default
    
    if isinstance(pattern, str):
        flags = re.IGNORECASE if case_insensitive else 0
        pattern = re.compile(pattern, flags)
    
    match = pattern.search(str(text))
    if match:
        try:
            return match.group(group)
        except IndexError:
            return match.group(0)
    return default


def regex_extract_number(text: str, default: float = 0.0) -> float:
    """
    Extracts the first number from text.
    
    Args:
        text: Text to search
        default: Default value if no number found
        
    Returns:
        Extracted number as float
    """
    result = regex_extract(text, r'(\d+(?:\.\d+)?)', group=1, default=None)
    if result is not None:
        try:
            return float(result.replace(',', ''))
        except:
            pass
    return default
