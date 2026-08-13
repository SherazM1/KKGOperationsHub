"""Parser service for Multistop-mode BOL Excel uploads."""

from __future__ import annotations

import re
from typing import Any

import pandas as pd

from app.models.bol_multistop_row import BolMultistopRow
from app.services.bol_standard_parser import is_csv_upload


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
    "weight_each": "weight each",
    "weight": "Weight",
}

OPTIONAL_COLUMN_SPECS: dict[str, tuple[str, ...]] = {
    "item_number": ("ITEM #", "Item #", "ITEM#", "Item#", "Item Number", "ITEM NUMBER"),
    "kit_value_each": ("Kit Value (EACH)", "Kit Value Each", "Kit Value"),
    "shipment_value": ("Shipment Value",),
    "chargeback_3_percent": (
        "3% Chargeback",
        "3 % Chargeback",
        "3 Percent Chargeback",
    ),
}
CSV_WORKSHEET_NAME = "CSV"

REQUIRED_COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    "trackers": ("Pick Up #", "Pickup #", "Pick Up", "Pickup"),
    "dc_city_state_zip": ("DC CITY, STATE, ZIP", "DC CITY STATE ZIP"),
    "dc_city": ("DC City", "Destination City"),
    "dc_state": ("DC ST", "DC State", "DC STATE", "State", "STATE", "ST"),
    "dc_zip": ("DC ZIP", "DC Zip", "Zip", "ZIP", "Zip Code", "ZIP Code", "Postal Code"),
    "dept": ("DEP.", "DEP", "Dept", "Department"),
    "pallet_description": ("Pallet Description", "PALLETDESCRIPTION"),
    "cases": ("CASES", "Case Qty", "Case QTY", "Case Quantity", "CASE QTY"),
    "total_pallets": (
        "Total PLT",
        "TOTAL PLT",
        "Total Pallets",
        "TOTAL PALLETS",
        "Total Pallet",
        "Total Plts",
        "Total PLTS",
    ),
    "weight_each": ("Weight Each", "WEIGHT EACH", "Weight", "WEIGHT"),
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
    cleaned = re.sub(r"\s*%\s*", " percent ", cleaned)
    # Tolerate minor punctuation differences (e.g., DEPT. vs DEPT, CITY, STATE, ZIP variants).
    cleaned = re.sub(r"[.,;:/\\()\-\_]+", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def _normalize_header_compact(header: str) -> str:
    cleaned = _normalize_header_for_fallback(header)
    # Second-pass comparison form: remove whitespace for compacted header variants.
    return cleaned.replace(" ", "")


def _build_header_lookups(
    columns: list[str],
) -> tuple[list[str], dict[str, list[str]], dict[str, list[str]], dict[str, list[str]], dict[str, list[str]]]:
    resolved_columns = [str(col) for col in columns]
    exact_columns: dict[str, list[str]] = {}
    lowered_exact_columns: dict[str, list[str]] = {}
    normalized_columns: dict[str, list[str]] = {}
    compact_columns: dict[str, list[str]] = {}

    for column in resolved_columns:
        exact_columns.setdefault(column, []).append(column)
        lowered_exact_columns.setdefault(column.lower(), []).append(column)
        normalized_columns.setdefault(_normalize_header_for_fallback(column), []).append(column)
        compact_columns.setdefault(_normalize_header_compact(column), []).append(column)

    return (
        resolved_columns,
        exact_columns,
        lowered_exact_columns,
        normalized_columns,
        compact_columns,
    )


def _first_match(candidates: list[str] | None) -> str | None:
    return candidates[0] if candidates else None


def _resolve_column_name(
    lookups: tuple[
        list[str],
        dict[str, list[str]],
        dict[str, list[str]],
        dict[str, list[str]],
        dict[str, list[str]],
    ],
    primary: str,
    aliases: tuple[str, ...],
) -> str | None:
    _, exact_columns, lowered_exact_columns, normalized_columns, compact_columns = lookups

    for candidate in (primary, *aliases):
        resolved_name = _first_match(exact_columns.get(candidate))
        if resolved_name is not None:
            return resolved_name

        resolved_name = _first_match(lowered_exact_columns.get(candidate.lower()))
        if resolved_name is not None:
            return resolved_name

        resolved_name = _first_match(
            normalized_columns.get(_normalize_header_for_fallback(candidate))
        )
        if resolved_name is not None:
            return resolved_name

        resolved_name = _first_match(compact_columns.get(_normalize_header_compact(candidate)))
        if resolved_name is not None:
            return resolved_name

    return None


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
    lookups = _build_header_lookups(columns)

    resolved: dict[str, str] = {}
    missing: list[str] = []

    for logical_name, source_name in REQUIRED_COLUMN_SPECS.items():
        aliases = REQUIRED_COLUMN_ALIASES.get(logical_name, ())
        resolved_name = _resolve_column_name(lookups, source_name, aliases)

        if resolved_name is None:
            if logical_name == "dc_city_state_zip":
                continue
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
    lookups = _build_header_lookups(columns)

    resolved: dict[str, str] = {}
    for logical_name, candidate_names in OPTIONAL_COLUMN_SPECS.items():
        resolved_name = _resolve_column_name(lookups, candidate_names[0], candidate_names[1:])
        if resolved_name is not None:
            resolved[logical_name] = resolved_name

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


def _combine_city_state_zip_from_values(row_values: dict[str, str]) -> str:
    if row_values.get("dc_city_state_zip", "").strip():
        return row_values["dc_city_state_zip"]

    city = row_values["dc_city"]
    state = row_values["dc_state"]
    zip_code = row_values["dc_zip"]
    city_state = ", ".join(part for part in (city, state) if part)
    return " ".join(part for part in (city_state, zip_code) if part)


def _parse_multistop_dataframe_rows(
    df: pd.DataFrame,
    column_map: dict[str, str],
    optional_column_map: dict[str, str],
) -> list[BolMultistopRow]:
    parsed_rows: list[BolMultistopRow] = []
    for index, row in df.iterrows():
        row_number = int(index) + 2
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
                dc_city_state_zip=_combine_city_state_zip_from_values(row_values),
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
                kit_value_each=optional_row_values.get("kit_value_each", ""),
                shipment_value=optional_row_values.get("shipment_value", ""),
                chargeback_3_percent=optional_row_values.get("chargeback_3_percent", ""),
                weight_each=row_values["weight_each"],
                weight=row_values["weight"],
            )
        )

    return parsed_rows


def _parse_multistop_bol_csv(file: Any) -> list[BolMultistopRow]:
    file.seek(0)
    try:
        df = pd.read_csv(file, dtype=object)
    except UnicodeDecodeError:
        file.seek(0)
        df = pd.read_csv(file, dtype=object, encoding="utf-8-sig")

    if df.empty:
        raise ValueError("CSV file contains no rows.")

    columns = df.columns.tolist()
    column_map = _resolve_columns(columns, worksheet_name=CSV_WORKSHEET_NAME)
    optional_column_map = _resolve_optional_columns(columns)
    parsed_rows = _parse_multistop_dataframe_rows(
        df,
        column_map=column_map,
        optional_column_map=optional_column_map,
    )

    if not parsed_rows:
        raise ValueError("No non-empty data rows found in CSV file.")

    file.seek(0)
    return parsed_rows


def parse_multistop_bol_excel(file: Any) -> list[BolMultistopRow]:
    if file is None:
        raise ValueError("No file uploaded. Upload an Excel or CSV file to parse.")

    if is_csv_upload(file):
        return _parse_multistop_bol_csv(file)

    resolved_sheet_name = _resolve_multistop_sheet_name(file)
    file.seek(0)
    df = pd.read_excel(file, sheet_name=resolved_sheet_name, dtype=object)

    if df.empty:
        raise ValueError(f"Worksheet '{resolved_sheet_name}' contains no rows.")

    columns = df.columns.tolist()
    column_map = _resolve_columns(columns, worksheet_name=resolved_sheet_name)
    optional_column_map = _resolve_optional_columns(columns)

    parsed_rows = _parse_multistop_dataframe_rows(
        df,
        column_map=column_map,
        optional_column_map=optional_column_map,
    )

    if not parsed_rows:
        raise ValueError(f"No non-empty data rows found in '{resolved_sheet_name}'.")

    return parsed_rows
