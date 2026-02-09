"""
Invoice Inspector Services - Public API

This module exposes the main functions for external codebases to import and use directly,
avoiding subprocess overhead.

Usage:
    from invoice_inspector.services import extract_invoice, ParsingError
    
    try:
        result = extract_invoice("path/to/file.xlsx")
        print(result.to_dict())
    except ParsingError as e:
        error_json = e.to_dict()
        # Handle error in UI
"""

from pathlib import Path
from typing import Union

# Core extraction function
from services.extraction_service import excel_data_extractor, scan_invoice_files

# Pipeline service for batch processing
from services.pipeline_service import PipelineService

# Master data service
from services.master_data_service import MasterDataService

# Models
from core.models import ExtractedInvoice, VerificationStatus

# Exceptions - for catching errors
from core.exceptions import (
    ParsingError,
    FileError,
    SheetError,
    DataExtractionError,
    ErrorCode,
    # Helper functions for creating errors
    create_file_not_found_error,
    create_sheet_not_found_error,
    create_header_not_found_error,
    create_total_row_not_found_error,
    create_value_parse_error,
    create_invalid_shipping_list_error,
    create_unknown_error,
)


def extract_invoice(file_path: Union[str, Path]) -> ExtractedInvoice:
    """
    Main entry point for extracting invoice data from an Excel file.
    
    This is the function external codebases should call.
    
    Args:
        file_path: Path to the Excel file (.xlsx or .xls)
        
    Returns:
        ExtractedInvoice object with all extracted data
        
    Raises:
        FileError: If file not found or cannot be opened
        SheetError: If required sheets/headers not found (INVALID_SHIPPING_LIST)
        DataExtractionError: If data cannot be parsed
    """
    file_path = Path(file_path)
    
    # Validate file exists
    if not file_path.exists():
        raise create_file_not_found_error(file_path.name)
    
    # Call the internal extractor
    return excel_data_extractor(file_path)


# Define public API
__all__ = [
    # Main functions
    "extract_invoice",
    "scan_invoice_files",
    
    # Services
    "PipelineService",
    "MasterDataService",
    
    # Models
    "ExtractedInvoice",
    "VerificationStatus",
    
    # Exceptions
    "ParsingError",
    "FileError", 
    "SheetError",
    "DataExtractionError",
    "ErrorCode",
    
    # Error helpers
    "create_file_not_found_error",
    "create_sheet_not_found_error",
    "create_header_not_found_error",
    "create_invalid_shipping_list_error",
]
