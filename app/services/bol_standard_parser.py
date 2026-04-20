"""Parser service for Standard-mode BOL Excel uploads."""

from __future__ import annotations

import re
from typing import Any

import pandas as pd

from app.models.bol_standard_row import BolStandardRow


STANDARD_SHEET_NAME = "MAIN LOAD SHEET"

REQUIRED_COLUMN_SPECS: dict[str, dict[str, str | list[str]]] = {
    "bol_number": {"primary": "BOL #", "fallback_aliases": []},
    "ship_date": {"primary": "ship date", "fallback_aliases": ["SHIP DATE"]},
    "carrier": {"primary": "Carrier", "fallback_aliases": ["CARRIER"]},
    "kk_load": {"primary": "KK Load", "fallback_aliases": ["KK LOAD"]},
    "kk_po": {"primary": "KK PO#", "fallback_aliases": ["KK PO #", "KK PO"]},
    "wm_po": {"primary": "WM PO #", "fallback_aliases": ["WM PO#", "WM PO"]},
    "dc_number": {"primary": "DC #", "fallback_aliases": ["DC#"]},
    "dc_name": {"primary": "DC NAME", "fallback_aliases": []},
    "dc_street": {"primary": "DC STREET", "fallback_aliases": []},
    "dc_city_state_zip": {"primary": "DC CITY, STATE, ZIP", "fallback_aliases": []},
    "item_number": {"primary": "ITEM #", "fallback_aliases": ["ITEM#"]},
    "upc": {"primary": "UPC", "fallback_aliases": []},
    "item_description": {
        "primary": "PalletDescription",
        "fallback_aliases": ["Pallet Description", "PALLETDESCRIPTION"],
    },
    "quantity": {"primary": "QTY", "fallback_aliases": []},
    "weight_each": {"primary": "weight each", "fallback_aliases": ["WEIGHT EACH"]},
}


def _normalize_header(header: str) -> str:
    cleaned = str(header).strip()
    cleaned = re.sub(r"\s*#\s*", "#", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = cleaned.upper()
    return cleaned


def _resolve_columns(columns: list[str]) -> dict[str, str]:
    resolved_columns = [str(col) for col in columns]
    exact_columns = {col: col for col in resolved_columns}
    normalized_columns = {_normalize_header(col): col for col in resolved_columns}

    resolved: dict[str, str] = {}
    missing: list[str] = []

    for logical_name, spec in REQUIRED_COLUMN_SPECS.items():
        primary = str(spec["primary"])
        fallback_aliases = [str(alias) for alias in spec["fallback_aliases"]]
        resolved_name: str | None = None

        # 1) exact match for the real workbook label first.
        if primary in exact_columns:
            resolved_name = exact_columns[primary]

        # 2) normalized match for minor formatting/case variations.
        if resolved_name is None:
            normalized_primary = _normalize_header(primary)
            if normalized_primary in normalized_columns:
                resolved_name = normalized_columns[normalized_primary]

        # 3) fallback aliases only if needed.
        if resolved_name is None:
            for alias in fallback_aliases:
                if alias in exact_columns:
                    resolved_name = exact_columns[alias]
                    break
                normalized_alias = _normalize_header(alias)
                if normalized_alias in normalized_columns:
                    resolved_name = normalized_columns[normalized_alias]
                    break

        if resolved_name is None:
            alias_text = f"; fallback aliases: {', '.join(fallback_aliases)}" if fallback_aliases else ""
            missing.append(f"{logical_name} (expected '{primary}'{alias_text})")
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
