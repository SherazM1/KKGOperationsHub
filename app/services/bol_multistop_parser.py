"""Parser service for Multistop-mode BOL Excel uploads."""

from __future__ import annotations

import re
from typing import Any

import pandas as pd

from app.models.bol_multistop_row import BolMultistopRow


MULTISTOP_SHEET_NAME_VARIANTS: tuple[str, ...] = (
    "Load sheet",
    "MAIN LOAD SHEET",
    "LOAD SHEET",
    "Main Load Sheet",
)

REQUIRED_COLUMN_SPECS: dict[str, str] = {
    "kk_load": "KK Load",
    "stop": "Stop",
    "trackers": "TRACKERS",
    "carrier": "Carrier",
    "load_number": "load#",
    "kk_po_number": "KK PO#",
    "bol_number": "BOL #",
    "ship_date": "ship date",
    "dc_name": "DC Name",
    "dc_address": "DC ADDRESS",
    "dc_city_state_zip": "DC City, State, Zip",
    "dc_city": "DC CITY",
    "dc_state": "DCST",
    "dc_zip": "DCZIP",
    "dc_number": "DC #",
    "country": "COUNTRY",
    "dept": "DEPT.",
    "target_po_number": "TGT PO #",
    "mabd": "MABD",
    "upc": "UPC",
    "pallet_description": "PalletDescription",
    "cases": "Cases",
    "total_pallets": "Total PLT",
    "kit_value_each": "Kit Value (EACH)",
    "shipment_value": "Shipment Value",
    "chargeback_3_percent": "3% Chargeback",
    "weight_each": "weight each",
    "weight": "Weight",
}

OPTIONAL_COLUMN_SPECS: dict[str, tuple[str, ...]] = {
    "item_number": ("ITEM #", "Item #", "ITEM#", "Item#", "Item Number", "ITEM NUMBER"),
}


def _normalize_header(header: str) -> str:
    cleaned = str(header).strip()
    cleaned = re.sub(r"\s*#\s*", "#", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.upper()


def _normalize_header_for_fallback(header: str) -> str:
    cleaned = str(header).replace("\r", " ").replace("\n", " ")
    cleaned = cleaned.strip().lower()
    cleaned = re.sub(r"\s*#\s*", "#", cleaned)
    # Tolerate minor punctuation differences (e.g., DEPT. vs DEPT, CITY, STATE, ZIP variants).
    cleaned = re.sub(r"[.,;:/\\()\-\_]+", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def _normalize_header_compact(header: str) -> str:
    cleaned = _normalize_header_for_fallback(header)
    # Second-pass comparison form: remove whitespace for compacted header variants.
    return cleaned.replace(" ", "")


def _candidate_multistop_sheet_names(available_sheet_names: list[str]) -> list[str]:
    exact_lookup = {name: name for name in available_sheet_names}
    normalized_lookup = {_normalize_header(name): name for name in available_sheet_names}
    candidates: list[str] = []

    for candidate in MULTISTOP_SHEET_NAME_VARIANTS:
        resolved_name = exact_lookup.get(candidate)
        if resolved_name and resolved_name not in candidates:
            candidates.append(resolved_name)

    for candidate in MULTISTOP_SHEET_NAME_VARIANTS:
        resolved_name = normalized_lookup.get(_normalize_header(candidate))
        if resolved_name and resolved_name not in candidates:
            candidates.append(resolved_name)

    return candidates


def _resolve_columns_with_missing(columns: list[str]) -> tuple[dict[str, str], list[str]]:
    resolved_columns = [str(col) for col in columns]
    exact_columns = {col: col for col in resolved_columns}
    lowered_exact_columns = {col.lower(): col for col in resolved_columns}
    normalized_columns = {_normalize_header_for_fallback(col): col for col in resolved_columns}
    compact_columns = {_normalize_header_compact(col): col for col in resolved_columns}

    resolved: dict[str, str] = {}
    missing: list[str] = []

    for logical_name, source_name in REQUIRED_COLUMN_SPECS.items():
        resolved_name: str | None = None

        # 1) Exact canonical workbook header match first.
        if source_name in exact_columns:
            resolved_name = exact_columns[source_name]

        # 2) Exact text match ignoring only case.
        if resolved_name is None:
            resolved_name = lowered_exact_columns.get(source_name.lower())

        # 3) Controlled normalized fallback for slight formatting/case/punctuation differences.
        if resolved_name is None:
            normalized_name = _normalize_header_for_fallback(source_name)
            if normalized_name in normalized_columns:
                resolved_name = normalized_columns[normalized_name]

        # 4) Compacted fallback for equivalent forms like "DC ST" vs "DCST".
        if resolved_name is None:
            compact_name = _normalize_header_compact(source_name)
            if compact_name in compact_columns:
                resolved_name = compact_columns[compact_name]

        if resolved_name is None:
            missing.append(f"{logical_name} (expected '{source_name}')")
        else:
            resolved[logical_name] = resolved_name

    return resolved, missing


def _resolve_multistop_sheet_name(file: Any) -> str:
    file.seek(0)
    workbook = pd.ExcelFile(file)
    available_sheet_names = [str(name) for name in workbook.sheet_names]
    candidate_sheet_names = _candidate_multistop_sheet_names(available_sheet_names)

    if not candidate_sheet_names:
        raise ValueError(
            "Required worksheet was not found for Multistop parsing. "
            f"Expected one of: {', '.join(MULTISTOP_SHEET_NAME_VARIANTS)}."
        )

    best_sheet_name = candidate_sheet_names[0]
    best_missing_count: int | None = None

    for candidate_sheet_name in candidate_sheet_names:
        header_df = workbook.parse(sheet_name=candidate_sheet_name, dtype=object, nrows=0)
        _, missing = _resolve_columns_with_missing(header_df.columns.tolist())

        if not missing:
            return candidate_sheet_name

        if best_missing_count is None or len(missing) < best_missing_count:
            best_missing_count = len(missing)
            best_sheet_name = candidate_sheet_name

    return best_sheet_name


def _resolve_columns(columns: list[str], worksheet_name: str) -> dict[str, str]:
    resolved, missing = _resolve_columns_with_missing(columns)
    if missing:
        detected_headers = ", ".join(str(col) for col in columns)
        raise ValueError(
            f"Missing required columns in '{worksheet_name}' for Multistop mode: "
            + "; ".join(missing)
            + f". [debug] selected worksheet='{worksheet_name}'; detected headers=[{detected_headers}]"
        )

    return resolved


def _resolve_optional_columns(columns: list[str]) -> dict[str, str]:
    resolved_columns = [str(col) for col in columns]
    exact_columns = {col: col for col in resolved_columns}
    lowered_exact_columns = {col.lower(): col for col in resolved_columns}
    normalized_columns = {_normalize_header_for_fallback(col): col for col in resolved_columns}
    compact_columns = {_normalize_header_compact(col): col for col in resolved_columns}

    resolved: dict[str, str] = {}
    for logical_name, candidate_names in OPTIONAL_COLUMN_SPECS.items():
        for candidate_name in candidate_names:
            resolved_name = exact_columns.get(candidate_name)
            if resolved_name is None:
                resolved_name = lowered_exact_columns.get(candidate_name.lower())
            if resolved_name is None:
                resolved_name = normalized_columns.get(
                    _normalize_header_for_fallback(candidate_name)
                )
            if resolved_name is None:
                resolved_name = compact_columns.get(_normalize_header_compact(candidate_name))

            if resolved_name is not None:
                resolved[logical_name] = resolved_name
                break

    return resolved


def _coerce_to_string(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _parse_stop_number(value: str) -> int | None:
    cleaned = (value or "").strip()
    if not cleaned:
        return None

    cleaned = cleaned.replace(",", "")
    try:
        parsed = float(cleaned)
    except ValueError:
        return None

    if not parsed.is_integer():
        return None

    return int(parsed)


def parse_multistop_bol_excel(file: Any) -> list[BolMultistopRow]:
    if file is None:
        raise ValueError("No file uploaded. Upload an Excel file to parse.")

    resolved_sheet_name = _resolve_multistop_sheet_name(file)
    file.seek(0)
    df = pd.read_excel(file, sheet_name=resolved_sheet_name, dtype=object)

    if df.empty:
        raise ValueError(f"Worksheet '{resolved_sheet_name}' contains no rows.")

    columns = df.columns.tolist()
    column_map = _resolve_columns(columns, worksheet_name=resolved_sheet_name)
    optional_column_map = _resolve_optional_columns(columns)

    parsed_rows: list[BolMultistopRow] = []
    for index, row in df.iterrows():
        row_number = index + 2
        row_values = {
            key: _coerce_to_string(row[source_column])
            for key, source_column in column_map.items()
        }
        optional_row_values = {
            key: _coerce_to_string(row[source_column])
            for key, source_column in optional_column_map.items()
        }

        if not any(row_values.values()):
            continue

        parsed_rows.append(
            BolMultistopRow(
                source_row_number=row_number,
                kk_load=row_values["kk_load"],
                stop=row_values["stop"],
                stop_number=_parse_stop_number(row_values["stop"]),
                trackers=row_values["trackers"],
                carrier=row_values["carrier"],
                load_number=row_values["load_number"],
                kk_po_number=row_values["kk_po_number"],
                bol_number=row_values["bol_number"],
                ship_date=row_values["ship_date"],
                dc_name=row_values["dc_name"],
                dc_address=row_values["dc_address"],
                dc_city_state_zip=row_values["dc_city_state_zip"],
                dc_city=row_values["dc_city"],
                dc_state=row_values["dc_state"],
                dc_zip=row_values["dc_zip"],
                dc_number=row_values["dc_number"],
                target_po_number=row_values["target_po_number"],
                item_number=optional_row_values.get("item_number", ""),
                upc=row_values["upc"],
                pallet_description=row_values["pallet_description"],
                cases=row_values["cases"],
                total_pallets=row_values["total_pallets"],
                kit_value_each=row_values["kit_value_each"],
                shipment_value=row_values["shipment_value"],
                chargeback_3_percent=row_values["chargeback_3_percent"],
                weight_each=row_values["weight_each"],
                weight=row_values["weight"],
            )
        )

    if not parsed_rows:
        raise ValueError(f"No non-empty data rows found in '{resolved_sheet_name}'.")

    return parsed_rows
