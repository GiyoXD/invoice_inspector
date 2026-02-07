from dataclasses import dataclass, field
from typing import Dict, List, Optional
from enum import Enum

class VerificationStatus(Enum):
    PASSED = "Passed"
    FAILED = "Failed"
    WARNING = "Warning"
    EXTRACTED = "Extracted"
    UNKNOWN = "Unknown"

@dataclass
class InvoiceSheetData:
    """Data extracted from a single sheet (Invoice, Packing List, Contract)."""
    col_qty_sf: Optional[float] = None
    col_amount: Optional[float] = None
    col_pallet_count: Optional[int] = None
    col_qty_pcs: Optional[int] = None
    col_net: Optional[float] = None
    col_gross: Optional[float] = None
    col_cbm: Optional[float] = None
    
    # Source workbook file name (for multi-file verification)
    source_file: str = ""
    
    # Set of column IDs detected as inspectable on this sheet
    # Populated by header matching and footer pattern detection
    target_inspect_col: set = field(default_factory=set)
    
    def to_dict(self):
        result = {k: v for k, v in self.__dict__.items() if v is not None}
        # Convert set to list for JSON serialization
        if 'target_inspect_col' in result:
            result['target_inspect_col'] = list(result['target_inspect_col'])
        return result

@dataclass
class ExtractedInvoice:
    """Top-level object representing a processed invoice file."""
    file_path: str
    file_name: str
    invoice_id: str = "Unknown"
    
    # Aggregated / Best-Guess Data
    col_amount: str = "N/A"
    col_qty_sf: str = "N/A"
    col_pallet_count: str = "N/A"
    col_qty_pcs: str = "N/A"
    col_net: str = "N/A"
    col_gross: str = "N/A"
    col_cbm: str = "N/A"
    
    # Verification details
    verification_details: str = ""
    status: VerificationStatus = VerificationStatus.EXTRACTED
    
    # Raw Sheet Data
    sheets: Dict[str, Dict] = field(default_factory=lambda: {
        'Invoice': {},
        'PackingList': {},
        'Contract': {}
    })
    
    # Candidate Sheets (for Packing List mostly)
    packing_candidates: List[Dict] = field(default_factory=list)
    sheet_status: Dict[str, bool] = field(default_factory=lambda: {
        'Invoice': False, 
        'PackingList': False, 
        'Contract': False
    })
    
    # Source Map: {'col_amount': 'InvoiceSheetName', ...}
    sources: Dict[str, str] = field(default_factory=dict)
    
    def to_dict(self):
        """Serialization helper."""
        d = self.__dict__.copy()
        # Convert Enum if present
        if isinstance(d['status'], VerificationStatus):
            d['status'] = d['status'].value
        return d
