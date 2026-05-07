"""Parse PURE and CDW order files for truck inventory."""

from __future__ import annotations

from io import BytesIO
from typing import NamedTuple

import openpyxl
import pandas as pd


class FileParseResult(NamedTuple):
    """Result of parsing a file."""

    success: bool
    rows: list[dict] = []
    error_message: str = ""
    file_type: str = ""


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
    Parse a combined load sheet file.
    
    Args:
        uploaded_file: Streamlit uploaded file object
        
    Returns:
        FileParseResult with rows from combined sheet
    """
    if uploaded_file is None:
        return FileParseResult(success=False, error_message="No file provided")
    
    try:
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
                error_message="Combined load sheet is empty",
            )
        
        rows = df.fillna("").to_dict(orient="records")
        
        return FileParseResult(
            success=True,
            rows=rows,
            file_type="combined_load_sheet",
        )
        
    except Exception as e:
        return FileParseResult(
            success=False,
            error_message=f"Error parsing combined load sheet: {str(e)}",
        )
