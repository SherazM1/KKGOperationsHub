"""Parse Truck Inventory Excel inputs."""

from __future__ import annotations

from typing import NamedTuple

import pandas as pd


class FileParseResult(NamedTuple):
    """Result of parsing a file."""

    success: bool
    rows: list[dict] = []
    error_message: str = ""
    file_type: str = ""


COMBINED_REQUIRED_COLUMNS = {
    "KKG Load #": ["KK Load", "KKG Load #", "KKG Load", "KK Load #", "Load #"],
    "Retailer PO #": ["WM PO #", "Retailer PO #", "Retailer PO", "PO #", "PO Number"],
    "Item #": ["ITEM #", "Item #", "Item Number", "ITEM", "Item"],
    "Qty": ["QTY", "Qty", "Quantity"],
}

COMBINED_SHEET_PRIORITY = ["FINAL LS", "REVISED LS", "LOAD SHEET"]


def detect_file_type(df: pd.DataFrame, filename: str) -> str:
    """Detect if file is PURE or CDW based on content and filename."""
    filename_lower = filename.lower()
    
    # Check filename hints
    if "pure" in filename_lower:
        return "pure"
    if "cdw" in filename_lower:
        return "cdw"
    
    # Check column names for hints
    columns_lower = [str(c).lower() for c in df.columns]
    
    if any("po" in col for col in columns_lower) and any("event" in col for col in columns_lower):
        return "pure"
    
    if any("purchase" in col for col in columns_lower) or any("order" in col for col in columns_lower):
        return "cdw"
    
    # Default fallback
    return "unknown"


def parse_excel_file(uploaded_file) -> FileParseResult:
    """
    Parse an uploaded Excel file, detecting PURE or CDW format.
    
    Args:
        uploaded_file: Streamlit uploaded file object
        
    Returns:
        FileParseResult with rows and detected file type
    """
    if uploaded_file is None:
        return FileParseResult(success=False, error_message="No file provided")
    
    try:
        # Read Excel file
        if uploaded_file.name.endswith((".xls", ".xlsx", ".xlsm")):
            df = pd.read_excel(uploaded_file, sheet_name=0)
        else:
            return FileParseResult(
                success=False,
                error_message=f"Unsupported file type: {uploaded_file.name}",
            )
        
        if df.empty:
            return FileParseResult(
                success=False,
                error_message="Excel file is empty",
            )
        
        # Detect file type
        file_type = detect_file_type(df, uploaded_file.name)
        
        # Convert to list of dicts
        rows = df.fillna("").to_dict(orient="records")
        
        return FileParseResult(
            success=True,
            rows=rows,
            file_type=file_type,
        )
        
    except Exception as e:
        return FileParseResult(
            success=False,
            error_message=f"Error parsing file: {str(e)}",
        )


def parse_combined_load_sheet(uploaded_file) -> FileParseResult:
    """
    Parse the combined load sheet operational workbook.

    The parser scans likely operational sheets and the first few header rows,
    then chooses the sheet/header combination with all required planning columns.
    
    Args:
        uploaded_file: Streamlit uploaded file object
        
    Returns:
        FileParseResult with rows from combined sheet
    """
    if uploaded_file is None:
        return FileParseResult(success=False, error_message="No file provided")
    
    try:
        if not uploaded_file.name.endswith((".xls", ".xlsx", ".xlsm")):
            return FileParseResult(
                success=False,
                error_message=f"Unsupported file type: {uploaded_file.name}",
            )

        workbook = pd.ExcelFile(uploaded_file)
        selected = _select_combined_load_sheet(workbook)
        if selected is None:
            return FileParseResult(
                success=False,
                error_message=_combined_sheet_error(workbook),
            )

        sheet_name, header_row, df = selected
        if df.empty:
            return FileParseResult(
                success=False,
                error_message="Combined load sheet is empty",
            )

        rows = df.fillna("").to_dict(orient="records")

        return FileParseResult(
            success=True,
            rows=rows,
            file_type=f"combined_load_sheet:{sheet_name}:header_row_{header_row + 1}",
        )

    except Exception as e:
        return FileParseResult(
            success=False,
            error_message=f"Error parsing combined load sheet: {str(e)}",
        )


def _select_combined_load_sheet(workbook: pd.ExcelFile) -> tuple[str, int, pd.DataFrame] | None:
    candidates = []
    for sheet_name in workbook.sheet_names:
        for header_row in range(0, 6):
            try:
                df = pd.read_excel(workbook, sheet_name=sheet_name, header=header_row)
            except ValueError:
                continue
            if df.empty:
                continue
            matched = _matched_required_columns(df.columns)
            priority = _sheet_priority(sheet_name)
            candidates.append((len(matched), priority, sheet_name, header_row, df))

    usable = [candidate for candidate in candidates if candidate[0] == len(COMBINED_REQUIRED_COLUMNS)]
    if not usable:
        return None

    usable.sort(key=lambda item: (-item[0], item[1], item[3]))
    _, _, sheet_name, header_row, df = usable[0]
    return sheet_name, header_row, df


def _matched_required_columns(columns) -> set[str]:
    normalized_columns = {_normalize_column(column) for column in columns}
    matched = set()
    for required_name, aliases in COMBINED_REQUIRED_COLUMNS.items():
        if any(_normalize_column(alias) in normalized_columns for alias in aliases):
            matched.add(required_name)
    return matched


def _combined_sheet_error(workbook: pd.ExcelFile) -> str:
    best_sheet = ""
    best_header = 0
    best_matches: set[str] = set()
    for sheet_name in workbook.sheet_names:
        for header_row in range(0, 6):
            try:
                df = pd.read_excel(workbook, sheet_name=sheet_name, header=header_row)
            except ValueError:
                continue
            matches = _matched_required_columns(df.columns)
            if len(matches) > len(best_matches):
                best_sheet = sheet_name
                best_header = header_row
                best_matches = matches

    missing = [name for name in COMBINED_REQUIRED_COLUMNS if name not in best_matches]
    if best_sheet:
        return (
            "No usable operational sheet found in combined load sheet. "
            f"Best candidate was '{best_sheet}' using row {best_header + 1} as headers. "
            f"Missing required columns: {', '.join(missing)}."
        )
    return (
        "No usable operational sheet found in combined load sheet. "
        f"Missing required columns: {', '.join(COMBINED_REQUIRED_COLUMNS)}."
    )


def _sheet_priority(sheet_name: str) -> int:
    normalized = _normalize_column(sheet_name)
    for index, preferred_name in enumerate(COMBINED_SHEET_PRIORITY):
        if _normalize_column(preferred_name) == normalized:
            return index
    return len(COMBINED_SHEET_PRIORITY)


def _normalize_column(value) -> str:
    return "".join(ch for ch in str(value).strip().lower() if ch.isalnum())
