"""Parser service for Standard-mode BOL Excel uploads."""

from __future__ import annotations

import re
from time import perf_counter
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
        "primary": "KKG Load #",
        "fallback_aliases": [
            "KKG LOAD #",
            "KKG Load",
            "KKG LOAD",
            "KK Load",
            "KK LOAD",
            "KK Load #",
        ],
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
            "QTY",
            "Qty",
            "qty",
            "Quantity",
            "QUANTITY",
        ],
    },
    "weight_each": {"primary": "weight each", "fallback_aliases": ["WEIGHT EACH"]},
}

OPTIONAL_COLUMN_SPECS: dict[str, dict[str, str | list[str]]] = {
    "carrier_pro_number": {
        "primary": "load#",
        "fallback_aliases": [
            "Load #",
            "LOAD #",
            "Load#",
            "LOAD#",
            "load",
            "Load",
            "LOAD",
        ],
    },
    "total_weight": {
        "primary": "Total Weight",
        "fallback_aliases": [
            "TOTAL WEIGHT",
            "total weight",
            "TotalWeight",
            "TOTALWEIGHT",
            "Line Total Weight",
            "Line Weight",
        ],
    },
    "pickup_number": {
        "primary": "Pick Up #",
        "fallback_aliases": [
            "Pickup #",
            "PICK UP #",
            "PICKUP #",
            "Pick Up",
            "Pickup",
            "Delivery Appt #",
            "Delivery Appt#",
            "Delivery Appointment #",
            "Delivery Appointment Number",
        ],
    },
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

PLT_QTY_EXPLICIT_ALIASES: tuple[str, ...] = (
    "Pallet Qty",
    "PALLET QTY",
    "TOTAL PALLETS",
    "Total Pallets",
    "Total PLT",
)
PLT_QTY_GENERIC_ALIASES: tuple[str, ...] = (
    "QTY",
    "Qty",
    "qty",
    "Quantity",
    "QUANTITY",
)

KK_LOAD_COLUMN_PRIORITY: tuple[str, ...] = (
    "KKG Load #",
    "KKG LOAD #",
    "KKG Load",
    "KKG LOAD",
    "KK Load",
    "KK LOAD",
    "KK Load #",
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
    lookups: tuple[dict[str, str], dict[str, str], dict[str, str]],
    primary: str,
    aliases: list[str] | tuple[str, ...],
) -> str | None:
    exact_columns, normalized_columns, compact_columns = lookups

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


def _resolve_dc_city_state_zip_components(
    lookups: tuple[dict[str, str], dict[str, str], dict[str, str]],
) -> dict[str, str]:
    resolved: dict[str, str] = {}
    for logical_name, aliases in DC_CITY_STATE_ZIP_COMPONENT_SPECS.items():
        source_column = _resolve_column_name(lookups, aliases[0], aliases[1:])
        if source_column is not None:
            resolved[logical_name] = source_column
    return resolved


def _resolve_kk_load_columns(
    columns: list[str],
    lookups: tuple[dict[str, str], dict[str, str], dict[str, str]] | None = None,
) -> list[str]:
    resolved_lookups = lookups or _build_column_lookups(columns)
    resolved_columns: list[str] = []
    for candidate in KK_LOAD_COLUMN_PRIORITY:
        source_column = _resolve_column_name(resolved_lookups, candidate, ())
        if source_column is not None and source_column not in resolved_columns:
            resolved_columns.append(source_column)
    return resolved_columns


def _resolve_plt_qty_column(
    lookups: tuple[dict[str, str], dict[str, str], dict[str, str]],
) -> str | None:
    explicit_column = _resolve_column_name(lookups, "PLT QTY", PLT_QTY_EXPLICIT_ALIASES)
    if explicit_column is not None:
        return explicit_column
    return _resolve_column_name(lookups, "QTY", PLT_QTY_GENERIC_ALIASES[1:])


def _resolve_columns_with_missing(columns: list[str]) -> tuple[dict[str, str], list[str]]:
    resolved_columns = [str(col) for col in columns]
    lookups = _build_column_lookups(resolved_columns)

    resolved: dict[str, str] = {}
    missing: list[str] = []

    for logical_name, spec in REQUIRED_COLUMN_SPECS.items():
        primary = str(spec["primary"])
        fallback_aliases = [str(alias) for alias in spec["fallback_aliases"]]
        if logical_name == "kk_load":
            kk_load_columns = _resolve_kk_load_columns(resolved_columns, lookups)
            resolved_name = kk_load_columns[0] if kk_load_columns else None
        elif logical_name == "plt_qty":
            resolved_name = _resolve_plt_qty_column(lookups)
        else:
            resolved_name = _resolve_column_name(lookups, primary, fallback_aliases)

        if resolved_name is None:
            if logical_name == "dc_city_state_zip":
                component_columns = _resolve_dc_city_state_zip_components(lookups)
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

    for logical_name, spec in OPTIONAL_COLUMN_SPECS.items():
        primary = str(spec["primary"])
        fallback_aliases = [str(alias) for alias in spec["fallback_aliases"]]
        resolved_name = _resolve_column_name(lookups, primary, fallback_aliases)
        if resolved_name is not None:
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


def _resolve_load_sheet_name(workbook: pd.ExcelFile) -> str:
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


def _effective_kk_load(row: pd.Series, kk_load_columns: list[str]) -> str:
    for source_column in kk_load_columns:
        value = _coerce_to_string(row[source_column])
        if value:
            return value
    return ""


def parse_standard_bol_excel(file: Any) -> list[BolStandardRow]:
    if file is None:
        raise ValueError("No file uploaded. Upload an Excel file to parse.")

    started_at = perf_counter()
    file.seek(0)
    workbook = pd.ExcelFile(file)
    workbook_loaded_at = perf_counter()
    print(
        "BOL parse timing: workbook_load="
        f"{workbook_loaded_at - started_at:.3f}s sheets={len(workbook.sheet_names)}"
    )

    resolved_sheet_name = _resolve_load_sheet_name(workbook)
    sheet_resolved_at = perf_counter()
    print(
        "BOL parse timing: sheet_detection="
        f"{sheet_resolved_at - workbook_loaded_at:.3f}s sheet={resolved_sheet_name!r}"
    )

    df = workbook.parse(sheet_name=resolved_sheet_name, dtype=object)
    sheet_loaded_at = perf_counter()
    print(
        "BOL parse timing: sheet_load="
        f"{sheet_loaded_at - sheet_resolved_at:.3f}s rows={len(df)} columns={len(df.columns)}"
    )

    if df.empty:
        raise ValueError(f"Worksheet '{resolved_sheet_name}' contains no rows.")

    column_map = _resolve_columns(df.columns.tolist(), worksheet_name=resolved_sheet_name)
    header_resolved_at = perf_counter()
    print(
        "BOL parse timing: header_resolution="
        f"{header_resolved_at - sheet_loaded_at:.3f}s"
    )

    kk_load_columns = _resolve_kk_load_columns(df.columns.tolist())

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
                kk_load=_effective_kk_load(row, kk_load_columns),
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
                total_weight=row_values.get("total_weight", ""),
                pickup_number=row_values.get("pickup_number", ""),
                carrier_pro_number=row_values.get("carrier_pro_number", ""),
            )
        )

    if not parsed_rows:
        raise ValueError(f"No non-empty data rows found in '{resolved_sheet_name}'.")

    rows_parsed_at = perf_counter()
    print(
        "BOL parse timing: row_parse="
        f"{rows_parsed_at - header_resolved_at:.3f}s parsed_rows={len(parsed_rows)} "
        f"total={rows_parsed_at - started_at:.3f}s"
    )

    return parsed_rows
