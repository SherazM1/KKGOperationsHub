"""Parser service for Standard-mode BOL Excel uploads."""

from __future__ import annotations

from typing import Any

import pandas as pd

from app.models.bol_standard_row import BolStandardRow


STANDARD_SHEET_NAME = "MAIN LOAD SHEET"

REQUIRED_COLUMN_MAP = {
    "bol_number": ["BOL #"],
    "ship_date": ["SHIP DATE"],
    "carrier": ["CARRIER"],
    "kk_load": ["KK LOAD"],
    "kk_po": ["KK PO#"],
    "wm_po": ["WM PO #"],
    "dc_number": ["DC #"],
    "dc_name": ["DC NAME"],
    "dc_street": ["DC STREET"],
    "dc_city_state_zip": ["DC CITY, STATE, ZIP"],
    "item_number": ["ITEM #"],
    "upc": ["UPC"],
    "item_description": ["ITEM DESCRIPTION", "DESCRIPTION", "ITEM DESC", "DESC"],
    "quantity": ["QTY"],
    "weight_each": ["WEIGHT EACH"],
}


def _normalize_header(header: str) -> str:
    cleaned = " ".join(str(header).strip().upper().split())
    return cleaned


def _resolve_columns(columns: list[str]) -> dict[str, str]:
    normalized_columns = {_normalize_header(col): col for col in columns}
    resolved: dict[str, str] = {}
    missing: list[str] = []

    for logical_name, aliases in REQUIRED_COLUMN_MAP.items():
        resolved_name: str | None = None
        for alias in aliases:
            if _normalize_header(alias) in normalized_columns:
                resolved_name = normalized_columns[_normalize_header(alias)]
                break

        if resolved_name is None:
            missing.append(f"{logical_name} ({', '.join(aliases)})")
        else:
            resolved[logical_name] = resolved_name

    if missing:
        raise ValueError(
            "Missing required columns in 'MAIN LOAD SHEET': "
            + "; ".join(missing)
        )

    return resolved


def _coerce_to_string(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def parse_standard_bol_excel(file: Any) -> list[BolStandardRow]:
    if file is None:
        raise ValueError("No file uploaded. Upload an Excel file to parse.")

    try:
        file.seek(0)
        df = pd.read_excel(file, sheet_name=STANDARD_SHEET_NAME, dtype=object)
    except ValueError as exc:
        # pandas raises ValueError when sheet_name is missing.
        if "Worksheet named" in str(exc):
            raise ValueError(
                "Required worksheet 'MAIN LOAD SHEET' was not found in the uploaded workbook."
            ) from exc
        raise

    if df.empty:
        raise ValueError("Worksheet 'MAIN LOAD SHEET' contains no rows.")

    column_map = _resolve_columns(df.columns.tolist())

    parsed_rows: list[BolStandardRow] = []

    for index, row in df.iterrows():
        row_number = index + 2  # Excel row with header on row 1.

        row_values = {
            key: _coerce_to_string(row[source_column])
            for key, source_column in column_map.items()
        }

        if not any(row_values.values()):
            continue

        parsed_rows.append(
            BolStandardRow(
                source_row_number=row_number,
                bol_number=row_values["bol_number"],
                ship_date=row_values["ship_date"],
                carrier=row_values["carrier"],
                kk_load=row_values["kk_load"],
                kk_po=row_values["kk_po"],
                wm_po=row_values["wm_po"],
                dc_number=row_values["dc_number"],
                dc_name=row_values["dc_name"],
                dc_street=row_values["dc_street"],
                dc_city_state_zip=row_values["dc_city_state_zip"],
                item_number=row_values["item_number"],
                upc=row_values["upc"],
                item_description=row_values["item_description"],
                quantity=row_values["quantity"],
                weight_each=row_values["weight_each"],
            )
        )

    if not parsed_rows:
        raise ValueError("No non-empty data rows found in 'MAIN LOAD SHEET'.")

    return parsed_rows

