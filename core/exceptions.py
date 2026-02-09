"""
Custom exception classes for the Invoice Inspector data parsing module.

These exceptions carry user-friendly error messages that can be serialized
to JSON for display in the UI.
"""

from enum import Enum
from typing import Any, Dict, Optional


class ErrorCode(Enum):
    """Error codes for categorizing parsing errors."""
    
    # File-level errors
    FILE_NOT_FOUND = "FILE_NOT_FOUND"
    FILE_PERMISSION_DENIED = "FILE_PERMISSION_DENIED"
    FILE_CORRUPTED = "FILE_CORRUPTED"
    FILE_LOCKED = "FILE_LOCKED"
    
    # Sheet-level errors
    SHEET_NOT_FOUND = "SHEET_NOT_FOUND"
    HEADER_NOT_FOUND = "HEADER_NOT_FOUND"
    TOTAL_ROW_NOT_FOUND = "TOTAL_ROW_NOT_FOUND"
    INVALID_SHIPPING_LIST = "INVALID_SHIPPING_LIST"  # Wrong sheet input - no recognizable headers
    
    # Data extraction errors
    COLUMN_NOT_FOUND = "COLUMN_NOT_FOUND"
    VALUE_PARSE_ERROR = "VALUE_PARSE_ERROR"
    DATA_VALIDATION_ERROR = "DATA_VALIDATION_ERROR"
    
    # General errors
    UNKNOWN_ERROR = "UNKNOWN_ERROR"


class ParsingError(Exception):
    """
    Base exception for all parsing-related errors.
    
    Attributes:
        error_code: An ErrorCode enum value for categorization.
        message: A user-friendly error message (can be in Thai or English).
        file_name: The name of the file being processed.
        sheet_name: The name of the sheet being processed (optional).
        context: Additional context data for debugging (optional).
    """
    
    def __init__(
        self,
        error_code: ErrorCode,
        message: str,
        file_name: str = "",
        sheet_name: str = "",
        context: Optional[Dict[str, Any]] = None
    ):
        self.error_code = error_code
        self.message = message
        self.file_name = file_name
        self.sheet_name = sheet_name
        self.context = context or {}
        
        # Build the full exception message for logging
        full_message = f"[{error_code.value}] {message}"
        if file_name:
            full_message += f" (File: {file_name})"
        if sheet_name:
            full_message += f" (Sheet: {sheet_name})"
        
        super().__init__(full_message)
    
    def to_dict(self) -> Dict[str, Any]:
        """
        Serializes the exception to a dictionary for JSON output.
        
        Returns:
            A dictionary containing all error details.
        """
        return {
            "error_code": self.error_code.value,
            "message": self.message,
            "file_name": self.file_name,
            "sheet_name": self.sheet_name,
            "context": self.context
        }


# ============================================================================
# Specific Exception Classes
# ============================================================================

class FileError(ParsingError):
    """Exception for file-level errors (not found, permission, corrupted)."""
    
    def __init__(
        self,
        error_code: ErrorCode,
        message: str,
        file_name: str,
        context: Optional[Dict[str, Any]] = None
    ):
        super().__init__(error_code, message, file_name, "", context)


class SheetError(ParsingError):
    """Exception for sheet-level errors (sheet not found, header missing)."""
    
    def __init__(
        self,
        error_code: ErrorCode,
        message: str,
        file_name: str,
        sheet_name: str = "",
        context: Optional[Dict[str, Any]] = None
    ):
        super().__init__(error_code, message, file_name, sheet_name, context)


class DataExtractionError(ParsingError):
    """Exception for data extraction errors (column missing, parse errors)."""
    
    def __init__(
        self,
        error_code: ErrorCode,
        message: str,
        file_name: str,
        sheet_name: str = "",
        context: Optional[Dict[str, Any]] = None
    ):
        super().__init__(error_code, message, file_name, sheet_name, context)


# ============================================================================
# Helper Functions
# ============================================================================

def create_file_not_found_error(file_name: str) -> FileError:
    """Creates a standardized FILE_NOT_FOUND error."""
    return FileError(
        error_code=ErrorCode.FILE_NOT_FOUND,
        message=f"File not found: '{file_name}'",
        file_name=file_name
    )


def create_sheet_not_found_error(
    file_name: str,
    sheet_type: str,
    searched_patterns: list = None
) -> SheetError:
    """Creates a standardized SHEET_NOT_FOUND error."""
    return SheetError(
        error_code=ErrorCode.SHEET_NOT_FOUND,
        message=f"Sheet '{sheet_type}' not found in file",
        file_name=file_name,
        context={"sheet_type": sheet_type, "searched_patterns": searched_patterns or []}
    )


def create_header_not_found_error(
    file_name: str,
    sheet_name: str,
    expected_headers: list = None
) -> SheetError:
    """Creates a standardized HEADER_NOT_FOUND error."""
    return SheetError(
        error_code=ErrorCode.HEADER_NOT_FOUND,
        message=f"Header row not found in sheet '{sheet_name}'",
        file_name=file_name,
        sheet_name=sheet_name,
        context={"expected_headers": expected_headers or []}
    )


def create_total_row_not_found_error(
    file_name: str,
    sheet_name: str
) -> SheetError:
    """Creates a standardized TOTAL_ROW_NOT_FOUND error."""
    return SheetError(
        error_code=ErrorCode.TOTAL_ROW_NOT_FOUND,
        message=f"Total Row Sheet '{sheet_name}'",
        file_name=file_name,
        sheet_name=sheet_name
    )


def create_value_parse_error(
    file_name: str,
    sheet_name: str,
    column_name: str,
    raw_value: Any
) -> DataExtractionError:
    """Creates a standardized VALUE_PARSE_ERROR error."""
    return DataExtractionError(
        error_code=ErrorCode.VALUE_PARSE_ERROR,
        message=f"Failed to parse value from column '{column_name}'",
        file_name=file_name,
        sheet_name=sheet_name,
        context={"column_name": column_name, "raw_value": str(raw_value)}
    )


def create_invalid_shipping_list_error(
    file_name: str,
    sheet_name: str = "",
    reason: str = ""
) -> SheetError:
    """
    Creates an INVALID_SHIPPING_LIST error when user provides wrong sheet.
    This indicates no recognizable headers were found - likely wrong input file.
    """
    message = "Invalid Shipping List - file does not contain recognizable headers"
    if reason:
        message += f" ({reason})"
    
    return SheetError(
        error_code=ErrorCode.INVALID_SHIPPING_LIST,
        message=message,
        file_name=file_name,
        sheet_name=sheet_name,
        context={"reason": reason or "No recognizable headers found"}
    )


def create_unknown_error(
    file_name: str,
    original_exception: Exception,
    sheet_name: str = "",
    operation: str = ""
) -> ParsingError:
    """
    Creates an UNKNOWN_ERROR for unexpected exceptions.
    Use this to wrap any exception that isn't explicitly handled.
    This ensures NO errors are silently swallowed.
    
    Args:
        file_name: The file being processed
        original_exception: The original exception that was caught
        sheet_name: Optional sheet name for context
        operation: Optional description of what was being attempted
    """
    error_msg = str(original_exception)
    if operation:
        error_msg = f"{operation}: {error_msg}"
    
    return ParsingError(
        error_code=ErrorCode.UNKNOWN_ERROR,
        message=error_msg,
        file_name=file_name,
        sheet_name=sheet_name,
        context={
            "exception_type": type(original_exception).__name__,
            "exception_message": str(original_exception),
            "operation": operation
        }
    )
