# Project Specification: Multi-Sheet Invoice Verification Tool

**Goal**: Build a standalone Python tool that validates invoice Excel files against a **Master List** source of truth. The tool verifies that data across multiple sheets ("Invoice", "Packing list", "Contract") within a file matches the official records in the Master List.

## 1. Overview
The tool treats the **Master List (CSV/Excel)** as the absolute source of truth.
For every Excel file in a folder:
1.  Identify the Invoice ID (from filename or content).
2.  Look up the expected values (Amount, Quantity, Pallets) in the Master List.
3.  Check each sheet (`Invoice`, `Packing list`, `Contract`) to see if it matches the Master List.

## 2. Input Data
*   **Source Files**: Excel files (`.xlsx`, `.xls`) with multiple sheets.
*   **Master List**: A CSV or Excel file containing columns: `Invoice ID`, `Total Amount`, `Total Quantity`, `Total Pallets`.

## 3. Extraction Requirements (Regex-Based)
Use flexible regex (similar to `extract_invoice_data.py`) to find "Total" rows in the breakdown sheets.

### Target Sheets & Fields
1.  **Invoice Sheet**:
    *   Find "Total" row. Extract: `Amount`, `Quantity`, `Pallets`.
2.  **Packing List Sheet**:
    *   Find "Total" row. Extract: `Quantity`, `Pallets`.
3.  **Contract Sheet** (if present):
    *   Find "Total" row. Extract: `Amount`, `Quantity`.

## 4. Verification Logic (Vs Master List)
For a file `File_123.xlsx` (ID: 123):
*   **Get Truth**: `MasterList[123]` -> {Amt: 1000, Qty: 50, Pallets: 5}

*   **Verify Invoice Sheet**:
    *   `Invoice[Amount]` == 1000?
    *   `Invoice[Quantity]` == 50?
*   **Verify Packing List Sheet**:
    *   `Packing[Quantity]` == 50?
    *   `Packing[Pallets]` == 5?
*   **Verify Contract Sheet**:
    *   `Contract[Amount]` == 1000?

## 5. Output Reports
Generate `verification_report.csv`:
*   `Filename`
*   `Invoice ID`
*   `Master Match Status` (PASS / FAIL)
*   `Detail`:
    *   "Invoice Sheet: Amount Mismatch (Found 900, Expected 1000)"
    *   "Packing List: Pallet Mismatch (Found 4, Expected 5)"

## 6. Technical Stack
*   **Language**: Python 3.9+
*   **Libraries**: `openpyxl`, `pandas`, `re`.
*   **Folder**: `sheet_verifier`
