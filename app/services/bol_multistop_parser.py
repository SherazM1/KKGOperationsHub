"""Parser service for Multistop-mode BOL Excel uploads."""

from __future__ import annotations

from typing import Any

import pandas as pd

from app.models.bol_multistop_row import BolMultistopRow
from app.services.bol_standard_parser import (
    is_csv_upload, _standard_dataframe, _resolve_load_sheet_name,
    _resolve_plt_qty_column, _resolve_wm_po_column, REQUIRED_COLUMN_SPECS as STANDARD_FIELDS,
    OPTIONAL_COLUMN_SPECS as STANDARD_OPTIONAL, _complete_derived_fields,
)
from app.services.bol_excel_headers import (
    ColumnMappingRequired, apply_column_override, build_column_lookups,
    normalize_header, compact_header, resolve_column_name, is_weight_summary,
    attach_mapping_samples,
    read_csv_source,
)
from app.services.bol_import_totals import reconcile_pallet_groups


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
    "target_po_number": ("RETAILER PO #", "Retailer PO", "PO #", "Customer PO #"),
    "pallet_description": ("Pallet Description", "PALLETDESCRIPTION"),
    "cases": ("CASES", "Case Qty", "Case QTY", "Case Quantity", "CASE QTY", "Units"),
    "total_pallets": (
        "Total PLT",
        "TOTAL PLT",
        "Total Pallets",
        "TOTAL PALLETS",
        "Total Pallet",
        "Total Plts",
        "Total PLTS",
        "PLT QTY",
        "Pallet Qty",
    ),
    "weight_each": ("Weight Each", "WEIGHT EACH", "Weight", "WEIGHT"),
}


def _normalize_header(header: str) -> str:
    return normalize_header(header)


def _build_header_lookups(columns):
    return build_column_lookups(columns)


def _resolve_column_name(lookups, primary, aliases):
    return resolve_column_name(lookups, primary, aliases)


_SHARED_FIELDS = {
    "kk_load": "kk_load", "carrier": "carrier", "load_number": "carrier_pro_number",
    "kk_po_number": "kk_po", "bol_number": "bol_number", "ship_date": "ship_date",
    "dc_name": "dc_name", "dc_address": "dc_street", "dc_number": "dc_number",
    "target_po_number": "wm_po", "upc": "upc", "pallet_description": "item_description",
    "cases": "unit_qty", "weight_each": "weight_each", "weight": "total_weight",
}


def _field_aliases(field):
    aliases = list(REQUIRED_COLUMN_ALIASES.get(field, ()))
    shared = _SHARED_FIELDS.get(field)
    if shared:
        spec = STANDARD_FIELDS.get(shared, STANDARD_OPTIONAL.get(shared))
        shared_aliases = [str(spec["primary"]), *spec["fallback_aliases"]]
        aliases = shared_aliases + aliases if field == "weight_each" else aliases + shared_aliases
    return tuple(dict.fromkeys(aliases))


def _resolve_columns_with_missing(columns, column_overrides=None):
    lookups = _build_header_lookups(columns)
    resolved, missing = {}, []
    for field, primary in REQUIRED_COLUMN_SPECS.items():
        try:
            selected = apply_column_override(field, column_overrides, columns)
            if selected is None:
                if field == "total_pallets":
                    selected = _resolve_plt_qty_column(lookups)
                elif field == "target_po_number":
                    spec = STANDARD_FIELDS["wm_po"]
                    selected = _resolve_wm_po_column(columns, lookups, spec["primary"], spec["fallback_aliases"])
                else:
                    selected = _resolve_column_name(lookups, primary, _field_aliases(field))
        except ColumnMappingRequired as exc:
            raise ColumnMappingRequired(str(exc), {field: next(iter(exc.requests.values()))}, columns) from exc
        if selected is not None:
            resolved[field] = selected
        elif field not in {"dc_city_state_zip", "stop", "country", "dept", "mabd", "dc_number", "upc"}:
            if field == "total_pallets" and _resolve_column_name(lookups, "Total PLT QTY", ("Total Load Pallets",)):
                continue
            if field == "total_pallets" and _resolve_column_name(lookups, "Units", ()) and _resolve_column_name(lookups, "PLT Weight", ("Pallet Weight",)):
                continue
            missing.append(f"{field} (expected '{primary}')")
    return resolved, missing


def _resolve_columns(columns, worksheet_name, column_overrides=None):
    resolved, missing = _resolve_columns_with_missing(columns, column_overrides)
    if missing:
        raise ColumnMappingRequired(
            f"Missing required columns in '{worksheet_name}' for Multistop mode: " + "; ".join(missing),
            {field.split(" (expected", 1)[0]: columns for field in missing}, columns,
        )
    return resolved


def _resolve_optional_columns(columns: list[str], column_overrides=None) -> dict[str, str]:
    lookups = _build_header_lookups(columns)

    resolved: dict[str, str] = {}
    specs = {**OPTIONAL_COLUMN_SPECS, **{
        field: (spec["primary"], *spec["fallback_aliases"])
        for field, spec in STANDARD_OPTIONAL.items() if field in {"group_pallet_total", "group_weight_total"}
    }}
    for logical_name, candidate_names in specs.items():
        try:
            resolved_name = apply_column_override(logical_name, column_overrides, columns) or _resolve_column_name(lookups, candidate_names[0], candidate_names[1:])
        except ColumnMappingRequired as exc:
            raise ColumnMappingRequired(str(exc), {logical_name: next(iter(exc.requests.values()))}, columns) from exc
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
    skipped_summaries = []
    inferred_stops: dict[tuple[str, str, str, str], int] = {}
    stops_per_load: dict[str, int] = {}
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
        if is_weight_summary(row_values):
            skipped_summaries.append(row_number)
            continue
        for field in ("dc_number", "dc_city", "dc_state", "dc_zip", "total_pallets"):
            row_values.setdefault(field, "")
        notes = []
        if "total_pallets" not in column_map:
            notes.append("No separate pallet/skid count supplied; review required.")
        if "dc_number" not in column_map:
            derived = {"dc_name": row_values["dc_name"]}
            notes.extend(_complete_derived_fields(derived))
            row_values["dc_number"] = derived["dc_number"]

        inferred = "stop" not in column_map
        if inferred:
            load = row_values["kk_load"]
            destination = (load, row_values["dc_number"], row_values["dc_address"],
                           _combine_city_state_zip_from_values(row_values))
            if destination not in inferred_stops:
                stops_per_load[load] = stops_per_load.get(load, 0) + 1
                inferred_stops[destination] = stops_per_load[load]
            row_values["stop"] = str(inferred_stops[destination])

        parsed_rows.append(
            BolMultistopRow(
                source_row_number=row_number,
                stop_inferred=inferred,
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
                upc=row_values.get("upc", ""),
                pallet_description=row_values["pallet_description"],
                cases=row_values["cases"],
                total_pallets=row_values["total_pallets"],
                kit_value_each=optional_row_values.get("kit_value_each", ""),
                shipment_value=optional_row_values.get("shipment_value", ""),
                chargeback_3_percent=optional_row_values.get("chargeback_3_percent", ""),
                weight_each=row_values["weight_each"],
                weight=row_values["weight"],
                source_values={str(header): _coerce_to_string(value) for header, value in row.items()},
                column_mapping={**column_map, **optional_column_map},
                parsing_notes=notes,
                group_pallet_total=optional_row_values.get("group_pallet_total", ""),
                group_weight_total=optional_row_values.get("group_weight_total", ""),
            )
        )

    if parsed_rows and skipped_summaries:
        parsed_rows[0].parsing_notes.append(f"Weight-only summary rows excluded from shipments: {', '.join(map(str, skipped_summaries))}.")
    return reconcile_pallet_groups(parsed_rows, pallet_field="total_pallets", weight_field="weight")


def _parse_multistop_bol_csv(file, column_overrides=None):
    file.seek(0)
    raw = read_csv_source(file)
    rows = _parse_source_frame(_standard_dataframe(raw), CSV_WORKSHEET_NAME, column_overrides)
    file.seek(0)
    return rows


def _parse_source_frame(df, sheet_name, column_overrides=None):
    if df.empty:
        raise ValueError(f"Worksheet '{sheet_name}' contains no rows.")
    columns = df.columns.tolist()
    try:
        mapping = _resolve_columns(columns, sheet_name, column_overrides)
        optional = _resolve_optional_columns(columns, column_overrides)
    except ColumnMappingRequired as exc:
        raise attach_mapping_samples(exc, columns, df.head(3).values.tolist())
    rows = _parse_multistop_dataframe_rows(df, mapping, optional)
    if not rows:
        raise ValueError(f"No non-empty data rows found in '{sheet_name}'.")
    return rows


def parse_multistop_bol_excel(file, worksheet_name=None, column_overrides=None):
    if file is None:
        raise ValueError("No file uploaded. Upload an Excel or CSV file to parse.")
    if is_csv_upload(file):
        return _parse_multistop_bol_csv(file, column_overrides)
    file.seek(0)
    with pd.ExcelFile(file) as workbook:
        selected = worksheet_name or _resolve_load_sheet_name(workbook)
        if selected not in workbook.sheet_names:
            raise ValueError(f"Worksheet '{selected}' was not found.")
        raw = workbook.parse(sheet_name=selected, dtype=object, header=None)
    return _parse_source_frame(_standard_dataframe(raw), selected, column_overrides)
