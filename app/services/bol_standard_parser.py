"""Parser service for Standard-mode BOL Excel uploads."""

from __future__ import annotations

import re
from typing import Any

import pandas as pd

from app.models.bol_standard_row import BolStandardRow


STANDARD_SHEET_NAME = "MAIN LOAD SHEET"
ACCEPTED_LOAD_SHEET_NAMES: tuple[str, ...] = (
    "main load sheet",
    "load sheet",
    "loads",
    "load",
)

LOAD_SHEET_NOT_FOUND_MESSAGE = (
    "Could not find a BOL load sheet. Expected a sheet named MAIN LOAD SHEET "
    "or LOAD SHEET, or a sheet containing BOL load headers such as KK Load, "
    "Carrier, DC #, TGT PO #, UPC, and Total Weight."
)

REQUIRED_COLUMN_SPECS: dict[str, dict[str, str | list[str]]] = {
    "bol_number": {"primary": "BOL #", "fallback_aliases": []},
    "ship_date": {"primary": "ship date", "fallback_aliases": ["SHIP DATE"]},
    "carrier": {"primary": "Carrier", "fallback_aliases": ["CARRIER"]},
    "kk_load": {
        "primary": "load#",
        "fallback_aliases": ["Load#", "LOAD#", "KK Load", "KK LOAD"],
    },
    "kk_po": {"primary": "KK PO#", "fallback_aliases": ["KK PO #", "KK PO"]},
    "wm_po": {
        "primary": "WM PO #",
        "fallback_aliases": [
            "WM PO#",
            "WM PO",
            "TGT PO #",
            "TGT PO#",
            "TGT PO",
            "PO #",
            "PO#",
        ],
    },
    "dc_number": {"primary": "DC #", "fallback_aliases": ["DC#"]},
    "dc_name": {"primary": "DC NAME", "fallback_aliases": []},
    "dc_street": {"primary": "DC STREET", "fallback_aliases": []},
    "dc_city_state_zip": {
        "primary": "DC CITY, STATE, ZIP",
        "fallback_aliases": ["DC City, State, Zip", "DC CITY STATE ZIP"],
    },
    "item_number": {"primary": "ITEM #", "fallback_aliases": ["ITEM#"]},
    "upc": {"primary": "UPC", "fallback_aliases": []},
    "item_description": {
        "primary": "PalletDescription",
        "fallback_aliases": ["Pallet Description", "PALLETDESCRIPTION"],
    },
    "unit_qty": {"primary": "Unit Qty", "fallback_aliases": ["UNIT QTY", "QTY"]},
    "plt_qty": {
        "primary": "PLT QTY",
        "fallback_aliases": [
            "Pallet Qty",
            "PALLET QTY",
            "TOTAL PALLETS",
            "Total Pallets",
            "Total PLT",
        ],
    },
    "weight_each": {"primary": "weight each", "fallback_aliases": ["WEIGHT EACH"]},
}

DC_CITY_STATE_ZIP_COMPONENT_SPECS: dict[str, tuple[str, ...]] = {
    "dc_city": ("DC CITY", "DC City"),
    "dc_state": ("DC ST", "DCST", "DC STATE", "DC State"),
    "dc_zip": ("DC ZIP", "DCZIP", "DC Zip"),
}

LOAD_SHEET_HEADER_SCAN_COLUMNS: tuple[str, ...] = (
    "KK Load",
    "Carrier",
    "KK PO#",
    "BOL #",
    "DC #",
    "DC NAME",
    "TGT PO #",
    "UPC",
    "Pallet Description",
    "QTY",
    "TOTAL PALLETS",
    "Total Weight",
)


def _normalize_header(header: str) -> str:
    cleaned = str(header).replace("\r", " ").replace("\n", " ").strip()
    cleaned = re.sub(r"\s*#\s*", "#", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = cleaned.upper()
    return cleaned


def _normalize_sheet_name(sheet_name: str) -> str:
    cleaned = str(sheet_name).replace("\r", " ").replace("\n", " ").strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.lower()


def _normalize_header_compact(header: str) -> str:
    return _normalize_header(header).replace(" ", "")


def _build_column_lookups(
    columns: list[str],
) -> tuple[dict[str, str], dict[str, str], dict[str, str]]:
    resolved_columns = [str(col) for col in columns]
    exact_columns = {col: col for col in resolved_columns}
    normalized_columns = {_normalize_header(col): col for col in resolved_columns}
    compact_columns = {_normalize_header_compact(col): col for col in resolved_columns}
    return exact_columns, normalized_columns, compact_columns


def _resolve_column_name(
    columns: list[str],
    primary: str,
    aliases: list[str] | tuple[str, ...],
) -> str | None:
    exact_columns, normalized_columns, compact_columns = _build_column_lookups(columns)

    for candidate in (primary, *aliases):
        if candidate in exact_columns:
            return exact_columns[candidate]

        normalized_candidate = _normalize_header(candidate)
        if normalized_candidate in normalized_columns:
            return normalized_columns[normalized_candidate]

        compact_candidate = _normalize_header_compact(candidate)
        if compact_candidate in compact_columns:
            return compact_columns[compact_candidate]

    return None


def _resolve_dc_city_state_zip_components(columns: list[str]) -> dict[str, str]:
    resolved: dict[str, str] = {}
    for logical_name, aliases in DC_CITY_STATE_ZIP_COMPONENT_SPECS.items():
        source_column = _resolve_column_name(columns, aliases[0], aliases[1:])
        if source_column is not None:
            resolved[logical_name] = source_column
    return resolved


def _resolve_columns_with_missing(columns: list[str]) -> tuple[dict[str, str], list[str]]:
    resolved_columns = [str(col) for col in columns]

    resolved: dict[str, str] = {}
    missing: list[str] = []

    for logical_name, spec in REQUIRED_COLUMN_SPECS.items():
        primary = str(spec["primary"])
        fallback_aliases = [str(alias) for alias in spec["fallback_aliases"]]
        resolved_name = _resolve_column_name(resolved_columns, primary, fallback_aliases)

        if resolved_name is None:
            if logical_name == "dc_city_state_zip":
                component_columns = _resolve_dc_city_state_zip_components(resolved_columns)
                if len(component_columns) == len(DC_CITY_STATE_ZIP_COMPONENT_SPECS):
                    resolved.update(component_columns)
                    continue

            alias_text = (
                f"; fallback aliases: {', '.join(fallback_aliases)}"
                if fallback_aliases
                else ""
            )
            missing.append(f"{logical_name} (expected '{primary}'{alias_text})")
        else:
            resolved[logical_name] = resolved_name

    return resolved, missing


def _resolve_columns(columns: list[str], worksheet_name: str) -> dict[str, str]:
    resolved, missing = _resolve_columns_with_missing(columns)
    if missing:
        raise ValueError(
            f"Missing required columns in '{worksheet_name}': "
            + "; ".join(missing)
        )

    return resolved


def _load_sheet_header_score(columns: list[str]) -> int:
    resolved_columns = [str(col) for col in columns]
    detected_headers = {
        _normalize_header(header)
        for header in resolved_columns
        if not str(header).startswith("Unnamed:")
    }
    detected_compact_headers = {
        _normalize_header_compact(header)
        for header in resolved_columns
        if not str(header).startswith("Unnamed:")
    }

    score = 0
    for expected_header in LOAD_SHEET_HEADER_SCAN_COLUMNS:
        if (
            _normalize_header(expected_header) in detected_headers
            or _normalize_header_compact(expected_header) in detected_compact_headers
        ):
            score += 1
    return score


def _resolve_load_sheet_name(file: Any) -> str:
    file.seek(0)
    workbook = pd.ExcelFile(file)

    normalized_lookup = {
        _normalize_sheet_name(sheet_name): str(sheet_name)
        for sheet_name in workbook.sheet_names
    }
    for accepted_name in ACCEPTED_LOAD_SHEET_NAMES:
        resolved_name = normalized_lookup.get(accepted_name)
        if resolved_name is not None:
            return resolved_name

    best_sheet_name: str | None = None
    best_score = 0
    for sheet_name in workbook.sheet_names:
        header_df = workbook.parse(sheet_name=sheet_name, dtype=object, nrows=0)
        score = _load_sheet_header_score(header_df.columns.tolist())
        if score > best_score:
            best_score = score
            best_sheet_name = str(sheet_name)

    if best_sheet_name is not None and best_score >= 6:
        return best_sheet_name

    raise ValueError(LOAD_SHEET_NOT_FOUND_MESSAGE)


def _coerce_to_string(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _combine_city_state_zip(row: pd.Series, column_map: dict[str, str]) -> str:
    if "dc_city_state_zip" in column_map:
        return _coerce_to_string(row[column_map["dc_city_state_zip"]])

    city = _coerce_to_string(row[column_map["dc_city"]])
    state = _coerce_to_string(row[column_map["dc_state"]])
    zip_code = _coerce_to_string(row[column_map["dc_zip"]])

    city_state = ", ".join(part for part in (city, state) if part)
    return " ".join(part for part in (city_state, zip_code) if part)


def _effective_bol_number(row_values: dict[str, str]) -> str:
    bol_number = row_values["bol_number"].strip()
    if bol_number:
        return bol_number

    return row_values["wm_po"].strip()


def parse_standard_bol_excel(file: Any) -> list[BolStandardRow]:
    if file is None:
        raise ValueError("No file uploaded. Upload an Excel file to parse.")

    resolved_sheet_name = _resolve_load_sheet_name(file)
    file.seek(0)
    df = pd.read_excel(file, sheet_name=resolved_sheet_name, dtype=object)

    if df.empty:
        raise ValueError(f"Worksheet '{resolved_sheet_name}' contains no rows.")

    column_map = _resolve_columns(df.columns.tolist(), worksheet_name=resolved_sheet_name)

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
                bol_number=_effective_bol_number(row_values),
                ship_date=row_values["ship_date"],
                carrier=row_values["carrier"],
                kk_load=row_values["kk_load"],
                kk_po=row_values["kk_po"],
                wm_po=row_values["wm_po"],
                dc_number=row_values["dc_number"],
                dc_name=row_values["dc_name"],
                dc_street=row_values["dc_street"],
                dc_city_state_zip=_combine_city_state_zip(row, column_map),
                item_number=row_values["item_number"],
                upc=row_values["upc"],
                item_description=row_values["item_description"],
                unit_qty=row_values["unit_qty"],
                plt_qty=row_values["plt_qty"],
                weight_each=row_values["weight_each"],
            )
        )

    if not parsed_rows:
        raise ValueError(f"No non-empty data rows found in '{resolved_sheet_name}'.")

    return parsed_rows
