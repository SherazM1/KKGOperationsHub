"""Parser service for Standard-mode BOL Excel uploads."""

from __future__ import annotations

import re
from time import perf_counter
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import pandas as pd

from app.models.bol_standard_row import BolStandardRow
from app.services.bol_excel_headers import (
    HEADER_SCAN_ROWS, ColumnLookups, ColumnMappingRequired, apply_column_override,
    build_column_lookups, compact_header, detect_headers, is_weight_summary,
    normalize_header, resolve_column_name,
    attach_mapping_samples,
    read_csv_source,
)
from app.services.bol_import_totals import reconcile_pallet_groups
from app.utils.bol_addresses import is_duplicate_bol_company, is_mclane_destination


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
CSV_WORKSHEET_NAME = "CSV"

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
            "RETAILER PO #",
            "Retailer PO",
            "MCLANE PO #",
            "WM PO",
            "wm_po",
            "WM_PO",
            "wm po",
            "WMPONumber",
            "WM_PO_Number",
            "WM PO Number",
            "WM PO No",
            "WM PO No.",
            "WM Purchase Order",
            "WM Purchase Order #",
            "RETAILER",
            "Retailer",
            "retailer",
            "Retailer #",
            "Walmart PO #",
            "Walmart PO#",
            "Walmart PO",
            "Walmart PO Number",
            "Walmart PO No",
            "Walmart PO No.",
            "Walmart Purchase Order",
            "Walmart Purchase Order #",
            "TGT PO #",
            "TGT PO#",
            "TGT PO",
            "TGT PO Number",
            "TGT PO No",
            "TGT PO No.",
            "Target PO #",
            "Target PO#",
            "Target PO",
            "Target PO Number",
            "Target PO No",
            "Target PO No.",
            "Target Purchase Order",
            "Target Purchase Order #",
            "PO #",
            "PO#",
            "PO",
            "P.O.",
            "P.O. #",
            "PO Number",
            "P.O. Number",
            "PONumber",
            "PO_Number",
            "PO No",
            "PO No.",
            "P.O. No.",
            "Purchase Order",
            "Purchase Order #",
            "Purchase Order Number",
            "PurchaseOrderNumber",
            "Purchase_Order_Number",
            "Customer PO",
            "Customer PO #",
            "Customer PO Number",
            "Customer P.O. Number",
            "Customer Purchase Order",
            "Customer Purchase Order #",
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
            "Total PLT QTY",
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
    "group_pallet_total": {
        "primary": "Total PLT QTY",
        "fallback_aliases": ["Total Load Pallets", "Load Pallet Total"],
    },
    "group_weight_total": {
        "primary": "Total Load Weight",
        "fallback_aliases": ["Load Weight Total"],
    },
    "shipwell_number": {
        "primary": "Shipwell #",
        "fallback_aliases": ["Shipwell", "Shipwell Number", "Shipwell No", "Shipwell ID"],
    },
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
    "dc_city": ("DC CITY", "DC Town", "Destination City", "Ship To City", "Consignee City"),
    "dc_state": ("DC ST", "DC STATE", "Destination State", "Ship To State", "Consignee State"),
    "dc_zip": ("DC ZIP", "DC Zip Code", "DC Postal Code", "Destination Zip", "Ship To Zip", "Consignee Zip"),
}

# Keep semantic and spelling variants explicit: fuzzy matching can silently map
# a different field (for example, unit weight to total weight).
ADDITIONAL_COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    "bol_number": ("BOL", "BOL Number", "BOL No", "Bill of Lading", "Bill of Lading Number", "Bill of Landing #"),
    "ship_date": ("Shipping Date", "Shipment Date", "Date Shipped", "Shiping Date"),
    "carrier": ("Carrier Name", "Shipping Carrier", "Carier", "Carrrier"),
    "kk_load": ("KK Load Number", "KK Load No", "KKG Load Number", "KKG Load No"),
    "kk_po": ("KK PO Number", "KK PO No", "KK Purchase Order", "KKG PO", "KKG PO #", "KKG PO Number", "KKG Purchase Order"),
    "wm_po": ("WMT PO", "WMT PO #", "WMT PO Number", "WMT Purchase Order", "Wal Mart PO #"),
    "dc_number": ("DC Number", "DC No", "DC ID", "Distribution Center #", "Distribution Centre Number", "DC"),
    "dc_name": ("Distribution Center Name", "Distribution Centre Name", "Destination Name", "Ship To Name", "Consignee Name"),
    "dc_street": ("DC ADDRESS", "DC Street Address", "DC Address 1", "DC Address Line 1", "DC Addr", "DC Adress", "DC Addres", "DC Steet", "Distribution Center Address", "Distribution Centre Address", "Destination Address", "Ship To Address", "Consignee Address"),
    "dc_city_state_zip": ("DC City ST Zip", "DC City State Zip Code", "Destination City State Zip", "Ship To City State Zip", "Consignee City State Zip"),
    "item_number": ("Item", "Item Number", "Item No", "SKU", "SKU #", "Product Number", "Product #"),
    "upc": ("UPC #", "UPC Number", "UPC Code", "Universal Product Code", "Barcode", "Bar Code"),
    "item_description": ("Item Description", "Product Description", "Description", "Pallet Desc", "Item Desc", "Pallet Descripton", "Pallet Desciption", "Item Descripton"),
    "unit_qty": ("Unit Quantity", "Units", "Total Units", "Quantity", "Unit Quanity", "Unit Quanitity"),
    "plt_qty": ("Pallet Quantity", "PLT Quantity", "Pallet Count", "Number of Pallets", "Pallets", "Pallet Quanity", "Pallet Quanitity"),
    "weight_each": ("PLT Weight", "Pallet Weight", "Each Weight", "Unit Weight", "Weight Per Unit", "Weight EA", "WT Each", "Each WT", "Wieght Each", "Weigth Each"),
    "carrier_pro_number": ("Load Number", "Load No", "Carrier PRO", "Carrier PRO Number", "PRO #", "PRO Number"),
    "total_weight": ("Weight", "Total WT", "Total Weight Lbs", "Total Wieght", "Total Weigth"),
    "pickup_number": ("Pickup Number", "Pick Up Number", "Pickup No", "Delivery Appt No", "Delivery Appointment No", "Delivery Appoinment #"),
}
for _field, _aliases in ADDITIONAL_COLUMN_ALIASES.items():
    _specs = REQUIRED_COLUMN_SPECS if _field in REQUIRED_COLUMN_SPECS else OPTIONAL_COLUMN_SPECS
    _specs[_field]["fallback_aliases"].extend(_aliases)

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
    "Total PLTS",
    "Total Pallet",
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
    return normalize_header(header)


def _normalize_sheet_name(sheet_name: str) -> str:
    cleaned = str(sheet_name).replace("\r", " ").replace("\n", " ").strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.lower()


def _normalize_header_compact(header: str) -> str:
    return compact_header(header)


def _build_column_lookups(
    columns: list[str],
) -> ColumnLookups:
    return build_column_lookups(columns)


def _resolve_column_name(
    lookups: ColumnLookups,
    primary: str,
    aliases: list[str] | tuple[str, ...],
) -> str | None:
    return resolve_column_name(lookups, primary, aliases)


def _resolve_dc_city_state_zip_components(
    lookups: ColumnLookups,
    column_overrides: dict[str, str] | None = None,
    columns: list[str] = (),
) -> dict[str, str]:
    resolved: dict[str, str] = {}
    for logical_name, aliases in DC_CITY_STATE_ZIP_COMPONENT_SPECS.items():
        field_lookups = lookups
        if logical_name == "dc_city" and "DC CITY" in columns and "DCCITY" in columns:
            # Established Sam's layout: DC CITY holds the street, DCCITY the
            # city. The row-level address check below validates the shift.
            field_lookups = _build_column_lookups([column for column in columns if column != "DCCITY"])
        source_column = apply_column_override(logical_name, column_overrides, columns) or _resolve_column_name(field_lookups, aliases[0], aliases[1:])
        if source_column is not None:
            resolved[logical_name] = source_column
    return resolved


def _resolve_kk_load_columns(
    columns: list[str],
    lookups: ColumnLookups | None = None,
) -> list[str]:
    resolved_lookups = lookups or _build_column_lookups(columns)
    resolved_columns: list[str] = []
    for candidate in (*KK_LOAD_COLUMN_PRIORITY, *ADDITIONAL_COLUMN_ALIASES["kk_load"]):
        source_column = _resolve_column_name(resolved_lookups, candidate, ())
        if source_column is not None and source_column not in resolved_columns:
            resolved_columns.append(source_column)
    return resolved_columns


def _resolve_plt_qty_column(
    lookups: ColumnLookups,
) -> str | None:
    try:
        explicit_column = _resolve_column_name(
            lookups, "PLT QTY",
            (*PLT_QTY_EXPLICIT_ALIASES, *ADDITIONAL_COLUMN_ALIASES["plt_qty"]),
        )
        if explicit_column is not None:
            return explicit_column
        return _resolve_column_name(lookups, "QTY", PLT_QTY_GENERIC_ALIASES[1:])
    except ColumnMappingRequired as exc:
        raise ColumnMappingRequired(str(exc), {"plt_qty": next(iter(exc.requests.values()))}) from exc


def _resolve_wm_po_column(
    columns: list[str],
    lookups: ColumnLookups,
    primary: str,
    aliases: list[str],
) -> str | None:
    # Prefer explicit PO labels before the legacy plain RETAILER header.
    po_aliases = [alias for alias in aliases if "PO" in compact_header(alias) or "PURCHASEORDER" in compact_header(alias)]
    resolved_name = _resolve_column_name(lookups, primary, po_aliases)
    if resolved_name is not None:
        return resolved_name

    # Unknown PO-like labels need review, not substring matching: POD and
    # Purchase Order Date are real source columns, not retailer PO numbers.
    return _resolve_column_name(lookups, "RETAILER", ())


def _resolve_total_weight_column(
    lookups: ColumnLookups,
    worksheet_name: str,
) -> str | None:
    if _normalize_sheet_name(worksheet_name) == "revised ls":
        revised_ls_weight = _resolve_column_name(
            lookups,
            "Weight",
            ("WEIGHT", "Total Weight", "TOTAL WEIGHT"),
        )
        if revised_ls_weight is not None:
            return revised_ls_weight

    spec = OPTIONAL_COLUMN_SPECS["total_weight"]
    return _resolve_column_name(
        lookups,
        str(spec["primary"]),
        [str(alias) for alias in spec["fallback_aliases"]],
    )


def _resolve_columns_with_missing(
    columns: list[str],
    worksheet_name: str,
    column_overrides: dict[str, str] | None = None,
) -> tuple[dict[str, str], list[str]]:
    resolved_columns = [str(col) for col in columns]
    lookups = _build_column_lookups(resolved_columns)

    resolved: dict[str, str] = {}
    missing: list[str] = []

    for logical_name, spec in REQUIRED_COLUMN_SPECS.items():
        primary = str(spec["primary"])
        fallback_aliases = [str(alias) for alias in spec["fallback_aliases"]]
        override = apply_column_override(logical_name, column_overrides, resolved_columns)
        if override is not None:
            resolved_name = override
        elif logical_name == "kk_load":
            kk_load_columns = _resolve_kk_load_columns(resolved_columns, lookups)
            resolved_name = kk_load_columns[0] if kk_load_columns else None
        elif logical_name == "wm_po":
            resolved_name = _resolve_wm_po_column(
                resolved_columns,
                lookups,
                primary,
                fallback_aliases,
            )
        elif logical_name == "plt_qty":
            resolved_name = _resolve_plt_qty_column(lookups)
        else:
            resolved_name = _resolve_column_name(lookups, primary, fallback_aliases)

        if resolved_name is None:
            if logical_name == "wm_po":
                continue
            if logical_name == "dc_number" and _resolve_column_name(
                lookups, "DC NAME", ADDITIONAL_COLUMN_ALIASES["dc_name"]
            ):
                continue
            if logical_name == "plt_qty" and _resolve_column_name(lookups, "Units", ()) and _resolve_column_name(lookups, "PLT Weight", ("Pallet Weight",)):
                continue
            if logical_name == "plt_qty" and _resolve_column_name(lookups, "Total PLT QTY", ("Total Load Pallets", "Load Pallet Total")):
                continue
            if logical_name == "dc_city_state_zip":
                component_columns = _resolve_dc_city_state_zip_components(lookups, column_overrides, resolved_columns)
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
        override = apply_column_override(logical_name, column_overrides, resolved_columns)
        if override is not None:
            resolved_name = override
        elif logical_name == "total_weight":
            resolved_name = _resolve_total_weight_column(lookups, worksheet_name)
        else:
            resolved_name = _resolve_column_name(lookups, primary, fallback_aliases)
        if resolved_name is not None:
            resolved[logical_name] = resolved_name

    # Some Sam's sheets label the street DC CITY and the real city DCCITY.
    # Retain both independently; only shift them when row contents confirm it.
    city_candidates = [c for c in resolved_columns if _normalize_header(c) == "DCCITY"]
    if city_candidates and "dc_city" in resolved and resolved["dc_city"] != city_candidates[0]:
        resolved["dc_city_alternate"] = city_candidates[0]
    return resolved, missing


def _resolve_columns(columns: list[str], worksheet_name: str, column_overrides=None) -> dict[str, str]:
    try:
        resolved, missing = _resolve_columns_with_missing(columns, worksheet_name, column_overrides)
    except ColumnMappingRequired as exc:
        primary, choices = next(iter(exc.requests.items()))
        logical = next((field for field, spec in {**REQUIRED_COLUMN_SPECS, **OPTIONAL_COLUMN_SPECS}.items()
                        if any(compact_header(name) == compact_header(primary) for name in (spec["primary"], *spec["fallback_aliases"]))), primary)
        logical = next((field for field, names in DC_CITY_STATE_ZIP_COMPONENT_SPECS.items()
                        if compact_header(names[0]) == compact_header(primary)), logical)
        raise ColumnMappingRequired(str(exc), {logical: choices}, columns) from exc
    if missing:
        detected_headers = ", ".join(str(col) for col in columns)
        raise ColumnMappingRequired(
            f"Missing required columns in '{worksheet_name}': "
            + "; ".join(missing)
            + f". Detected headers: [{detected_headers}]",
            {field.split(" (expected", 1)[0]: columns for field in missing}, columns
        )

    return resolved


def _load_sheet_header_score(columns: list[str]) -> int:
    # Scoring is only for finding a header. Actual mapping retains duplicates
    # and requires a user choice rather than inheriting this temporary view.
    unique = list({compact_header(column): column for column in columns if column}.values())
    resolved, _ = _resolve_columns_with_missing(unique, "")
    # Count source columns once even when QTY supplies both quantity fields.
    return len(set(resolved.values()))


def _choose_load_sheet(candidates):
    usable = [(name, score) for name, score in candidates if score >= 6]
    for preferred in ACCEPTED_LOAD_SHEET_NAMES:
        for name, _ in usable:
            if _normalize_sheet_name(name) == preferred:
                return name
    if usable:
        return max(usable, key=lambda item: item[1])[0]
    raise ValueError(LOAD_SHEET_NOT_FOUND_MESSAGE)


def _resolve_load_sheet_name(workbook: pd.ExcelFile) -> str:
    candidates = []
    for name in workbook.sheet_names:
        raw = workbook.parse(sheet_name=name, dtype=object, nrows=HEADER_SCAN_ROWS, header=None)
        headers, _ = _standard_headers_from_rows(list(raw.itertuples(index=False, name=None)))
        candidates.append((str(name), _load_sheet_header_score(headers)))
    return _choose_load_sheet(candidates)


def _resolve_openpyxl_load_sheet_name(workbook: Any) -> str:
    candidates = []
    for worksheet in workbook.worksheets:
        if worksheet.sheet_state != "visible":
            continue
        headers, _ = _read_standard_headers(worksheet)
        candidates.append((str(worksheet.title), _load_sheet_header_score(headers)))
    return _choose_load_sheet(candidates)


def _missing_worksheet_error(worksheet_name: str, available_sheets: list[str]) -> ValueError:
    available_text = ", ".join(available_sheets) if available_sheets else "(none)"
    return ValueError(
        f"Worksheet '{worksheet_name}' was not found. Available worksheets: {available_text}."
    )


def is_csv_upload(file: Any) -> bool:
    """Return True when the uploaded BOL source appears to be a CSV file."""
    filename = str(getattr(file, "name", "") or "")
    content_type = str(getattr(file, "type", "") or getattr(file, "content_type", "") or "")
    return filename.lower().endswith(".csv") or content_type.lower() in {
        "text/csv",
        "application/csv",
    }


def get_excel_sheet_names(file: Any) -> list[str]:
    """Return visible worksheet names in workbook order without consuming the upload stream."""
    if file is None:
        raise ValueError("No file uploaded. Upload an Excel file to inspect worksheets.")

    try:
        file.seek(0)
        workbook = load_workbook(file, read_only=True, data_only=True)
        try:
            return [
                str(worksheet.title)
                for worksheet in workbook.worksheets
                if worksheet.sheet_state == "visible"
            ]
        finally:
            workbook.close()
    except (InvalidFileException, OSError, KeyError, ValueError):
        try:
            file.seek(0)
            workbook = pd.ExcelFile(file)
            return [str(sheet_name) for sheet_name in workbook.sheet_names]
        except Exception as exc:
            raise ValueError(
                "Could not read Excel worksheets. Upload a valid Excel workbook."
            ) from exc
    finally:
        try:
            file.seek(0)
        except Exception:
            pass


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


def _effective_kk_load_from_values(row_values: dict[str, str], kk_load_columns: list[str]) -> str:
    for source_column in kk_load_columns:
        value = row_values.get(f"kk_load::{source_column}", "")
        if value:
            return value
    return ""


def _iter_openpyxl_standard_rows(
    worksheet: Any,
    column_map: dict[str, str],
    kk_load_columns: list[str],
    header_values: list[str] | None = None,
    first_data_row: int = 2,
) -> list[BolStandardRow]:
    header_values = header_values or [
        _coerce_to_string(value)
        for value in next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True), [])
    ]
    column_indices = {header: index for index, header in enumerate(header_values)}
    needed_columns = set(column_map.values()) | set(kk_load_columns)
    needed_indices = {
        column_name: column_indices[column_name]
        for column_name in needed_columns
        if column_name in column_indices
    }
    max_needed_index = max(needed_indices.values(), default=-1)
    if max_needed_index < 0:
        return []

    parsed_rows: list[BolStandardRow] = []
    skipped_summaries = []

    for row_number, values in enumerate(
        worksheet.iter_rows(
            min_row=first_data_row,
            max_col=len(header_values),
            values_only=True,
        ),
        start=first_data_row,
    ):
        row_values: dict[str, str] = {}
        for logical_name, source_column in column_map.items():
            index = needed_indices[source_column]
            row_values[logical_name] = _coerce_to_string(
                values[index] if index < len(values) else None
            )
        for source_column in kk_load_columns:
            index = needed_indices[source_column]
            row_values[f"kk_load::{source_column}"] = _coerce_to_string(
                values[index] if index < len(values) else None
            )
        row_values.setdefault("wm_po", "")

        if not any(row_values.values()):
            continue
        if is_weight_summary(row_values):
            skipped_summaries.append(row_number)
            continue
        source_mapping = _normalize_source_address(row_values, column_map)
        notes = _complete_derived_fields(row_values)
        parsed_rows.append(
            BolStandardRow(
                source_row_number=row_number,
                bol_number=_effective_bol_number(row_values),
                ship_date=row_values["ship_date"],
                carrier=row_values["carrier"],
                kk_load=_effective_kk_load_from_values(row_values, kk_load_columns),
                kk_po=row_values["kk_po"],
                wm_po=row_values["wm_po"],
                dc_number=row_values["dc_number"],
                dc_name=row_values["dc_name"],
                dc_street=row_values["dc_street"],
                dc_city_state_zip=_combine_city_state_zip_from_values(row_values),
                item_number=row_values["item_number"],
                upc=row_values["upc"],
                item_description=row_values["item_description"],
                unit_qty=row_values["unit_qty"],
                plt_qty=row_values["plt_qty"],
                weight_each=row_values["weight_each"],
                total_weight=row_values.get("total_weight", ""),
                pickup_number=row_values.get("pickup_number", ""),
                shipwell_number=row_values.get("shipwell_number", ""),
                carrier_pro_number=row_values.get("carrier_pro_number", ""),
                source_values={header: _coerce_to_string(values[i]) for i, header in enumerate(header_values) if header},
                column_mapping=source_mapping,
                parsing_notes=notes,
                group_pallet_total=row_values.get("group_pallet_total", ""),
                group_weight_total=row_values.get("group_weight_total", ""),
            )
        )

    if parsed_rows and skipped_summaries:
        parsed_rows[0].parsing_notes.append(f"Weight-only summary rows excluded from shipments: {', '.join(map(str, skipped_summaries))}.")
    return reconcile_pallet_groups(parsed_rows)


def _source_column_mapping(column_map: dict[str, str]) -> dict[str, str]:
    mapping = dict(column_map)
    if "dc_number" not in mapping:
        mapping["dc_number"] = "Derived from explicit DC number in " + mapping.get("dc_name", "DC Name")
    if "plt_qty" not in mapping:
        mapping["plt_qty"] = "Not provided; requires review"
    return mapping


def _normalize_source_address(values: dict[str, str], column_map: dict[str, str]) -> dict[str, str]:
    mapping = _source_column_mapping(column_map)
    if "dc_city_alternate" in values and is_duplicate_bol_company(
        values.get("dc_name", ""), values.get("dc_street", "")
    ):
        values["dc_street"] = values.get("dc_city", "")
        values["dc_city"] = values["dc_city_alternate"]
        mapping["dc_street"] = column_map["dc_city"]
        mapping["dc_city"] = column_map["dc_city_alternate"]
    mapping.pop("dc_city_alternate", None)
    return mapping


def _complete_derived_fields(values: dict[str, str]) -> list[str]:
    notes = []
    if values.get("dc_street", "").strip() and not re.search(r"[A-Za-z]", values["dc_street"]):
        notes.append("DC STREET must include a street name; review required.")
    for component in ("dc_city", "dc_state", "dc_zip"):
        if component in values and not values[component].strip():
            notes.append(f"Missing {component.replace('_', ' ').upper()}; review required.")
    if "dc_number" not in values:
        matches = re.findall(r"\bDC\s*#?\s*(\d+)\b", values.get("dc_name", ""), re.IGNORECASE)
        values["dc_number"] = matches[0] if len(matches) == 1 else ""
        if values["dc_number"]:
            notes.append("DC number derived from DC Name.")
        elif not is_mclane_destination(values.get("dc_name", "")):
            notes.append("DC Name has no unambiguous DC number; review required.")
    if "plt_qty" not in values:
        values["plt_qty"] = ""
        notes.append("No separate pallet/skid count supplied; review required.")
    return notes


def _read_standard_headers(worksheet: Any) -> tuple[list[str], int]:
    """Join a recognized header continuation without consuming a shipment row."""
    top_rows = list(worksheet.iter_rows(min_row=1, max_row=HEADER_SCAN_ROWS, values_only=True))
    return _standard_headers_from_rows(top_rows)


def _standard_headers_from_rows(top_rows: list[Any]) -> tuple[list[str], int]:
    layout = detect_headers(top_rows, _load_sheet_header_score)
    return layout.headers, layout.first_data_row


def _standard_dataframe(raw: pd.DataFrame) -> pd.DataFrame:
    headers, first_data_row = _standard_headers_from_rows(list(raw.head(HEADER_SCAN_ROWS).itertuples(index=False, name=None)))
    df = raw.iloc[first_data_row - 1:].copy()
    df.columns = headers
    # The dataframe reader adds two to the index for the original Excel row.
    df.index = range(first_data_row - 2, first_data_row - 2 + len(df))
    return df


def _combine_city_state_zip_from_values(row_values: dict[str, str]) -> str:
    if "dc_city_state_zip" in row_values:
        return row_values["dc_city_state_zip"]

    city = row_values["dc_city"]
    state = row_values["dc_state"]
    zip_code = row_values["dc_zip"]
    city_state = ", ".join(part for part in (city, state) if part)
    return " ".join(part for part in (city_state, zip_code) if part)


def _parse_standard_bol_excel_openpyxl(
    file: Any,
    worksheet_name: str | None = None,
    column_overrides: dict[str, str] | None = None,
) -> list[BolStandardRow]:
    started_at = perf_counter()
    file.seek(0)
    workbook = load_workbook(file, read_only=True, data_only=True)
    workbook_loaded_at = perf_counter()
    print(
        "BOL parse timing: workbook_load_openpyxl="
        f"{workbook_loaded_at - started_at:.3f}s sheets={len(workbook.sheetnames)}"
    )

    try:
        visible_sheet_names = [
            str(worksheet.title)
            for worksheet in workbook.worksheets
            if worksheet.sheet_state == "visible"
        ]
        if worksheet_name is not None:
            if worksheet_name not in visible_sheet_names:
                raise _missing_worksheet_error(worksheet_name, visible_sheet_names)
            resolved_sheet_name = worksheet_name
        else:
            resolved_sheet_name = _resolve_openpyxl_load_sheet_name(workbook)
        sheet_resolved_at = perf_counter()
        print(
            "BOL parse timing: sheet_detection_openpyxl="
            f"{sheet_resolved_at - workbook_loaded_at:.3f}s sheet={resolved_sheet_name!r}"
        )

        worksheet = workbook[resolved_sheet_name]
        header_values, first_data_row = _read_standard_headers(worksheet)
        header_loaded_at = perf_counter()
        print(
            "BOL parse timing: header_row_openpyxl="
            f"{header_loaded_at - sheet_resolved_at:.3f}s columns={len(header_values)}"
        )

        if not header_values:
            raise ValueError(f"Worksheet '{resolved_sheet_name}' contains no rows.")

        try:
            column_map = _resolve_columns(header_values, worksheet_name=resolved_sheet_name, column_overrides=column_overrides)
        except ColumnMappingRequired as exc:
            preview = worksheet.iter_rows(min_row=first_data_row, max_row=first_data_row + 2, values_only=True)
            raise attach_mapping_samples(exc, header_values, preview)
        kk_load_columns = [column_map["kk_load"]] if (column_overrides or {}).get("kk_load") else _resolve_kk_load_columns(header_values)
        columns_resolved_at = perf_counter()
        print(
            "BOL parse timing: column_resolution_openpyxl="
            f"{columns_resolved_at - header_loaded_at:.3f}s"
        )

        parsed_rows = _iter_openpyxl_standard_rows(worksheet, column_map, kk_load_columns, header_values, first_data_row)
        rows_parsed_at = perf_counter()
        print(
            "BOL parse timing: row_stream_openpyxl="
            f"{rows_parsed_at - columns_resolved_at:.3f}s parsed_rows={len(parsed_rows)} "
            f"total={rows_parsed_at - started_at:.3f}s"
        )

        if not parsed_rows:
            raise ValueError(f"No non-empty data rows found in '{resolved_sheet_name}'.")

        return parsed_rows
    finally:
        workbook.close()


def _parse_standard_dataframe_rows(
    df: pd.DataFrame,
    column_map: dict[str, str],
    kk_load_columns: list[str],
) -> list[BolStandardRow]:
    parsed_rows: list[BolStandardRow] = []
    skipped_summaries = []

    for index, row in df.iterrows():
        row_number = int(index) + 2  # Source row with header on row 1.

        row_values = {
            key: _coerce_to_string(row[source_column])
            for key, source_column in column_map.items()
        }
        row_values.setdefault("wm_po", "")

        if not any(row_values.values()):
            continue
        if is_weight_summary(row_values):
            skipped_summaries.append(row_number)
            continue

        source_mapping = _normalize_source_address(row_values, column_map)
        notes = _complete_derived_fields(row_values)
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
                dc_city_state_zip=_combine_city_state_zip_from_values(row_values),
                item_number=row_values["item_number"],
                upc=row_values["upc"],
                item_description=row_values["item_description"],
                unit_qty=row_values["unit_qty"],
                plt_qty=row_values["plt_qty"],
                weight_each=row_values["weight_each"],
                total_weight=row_values.get("total_weight", ""),
                pickup_number=row_values.get("pickup_number", ""),
                shipwell_number=row_values.get("shipwell_number", ""),
                carrier_pro_number=row_values.get("carrier_pro_number", ""),
                source_values={str(header): _coerce_to_string(value) for header, value in row.items()},
                column_mapping=source_mapping,
                parsing_notes=notes,
                group_pallet_total=row_values.get("group_pallet_total", ""),
                group_weight_total=row_values.get("group_weight_total", ""),
            )
        )

    if parsed_rows and skipped_summaries:
        parsed_rows[0].parsing_notes.append(f"Weight-only summary rows excluded from shipments: {', '.join(map(str, skipped_summaries))}.")
    return reconcile_pallet_groups(parsed_rows)


def _parse_standard_bol_csv(file: Any, column_overrides=None) -> list[BolStandardRow]:
    started_at = perf_counter()
    file.seek(0)
    df = _standard_dataframe(read_csv_source(file))
    csv_loaded_at = perf_counter()
    print(
        "BOL parse timing: csv_load="
        f"{csv_loaded_at - started_at:.3f}s rows={len(df)} columns={len(df.columns)}"
    )

    if df.empty:
        raise ValueError("CSV file contains no rows.")

    try:
        column_map = _resolve_columns(df.columns.tolist(), worksheet_name=CSV_WORKSHEET_NAME, column_overrides=column_overrides)
    except ColumnMappingRequired as exc:
        raise attach_mapping_samples(exc, df.columns, df.head(3).values.tolist())
    header_resolved_at = perf_counter()
    print(
        "BOL parse timing: csv_header_resolution="
        f"{header_resolved_at - csv_loaded_at:.3f}s"
    )

    kk_load_columns = [column_map["kk_load"]] if (column_overrides or {}).get("kk_load") else _resolve_kk_load_columns(df.columns.tolist())
    parsed_rows = _parse_standard_dataframe_rows(
        df,
        column_map=column_map,
        kk_load_columns=kk_load_columns,
    )

    if not parsed_rows:
        raise ValueError("No non-empty data rows found in CSV file.")

    rows_parsed_at = perf_counter()
    print(
        "BOL parse timing: csv_row_parse="
        f"{rows_parsed_at - header_resolved_at:.3f}s parsed_rows={len(parsed_rows)} "
        f"total={rows_parsed_at - started_at:.3f}s"
    )

    file.seek(0)
    return parsed_rows


def parse_standard_bol_excel(
    file: Any,
    worksheet_name: str | None = None,
    column_overrides: dict[str, str] | None = None,
) -> list[BolStandardRow]:
    if file is None:
        raise ValueError("No file uploaded. Upload an Excel or CSV file to parse.")

    if is_csv_upload(file):
        return _parse_standard_bol_csv(file, column_overrides)

    try:
        return _parse_standard_bol_excel_openpyxl(file, worksheet_name=worksheet_name, column_overrides=column_overrides)
    except (InvalidFileException, OSError, KeyError):
        file.seek(0)

    started_at = perf_counter()
    file.seek(0)
    workbook = pd.ExcelFile(file)
    workbook_loaded_at = perf_counter()
    print(
        "BOL parse timing: workbook_load="
        f"{workbook_loaded_at - started_at:.3f}s sheets={len(workbook.sheet_names)}"
    )

    available_sheet_names = [str(sheet_name) for sheet_name in workbook.sheet_names]
    if worksheet_name is not None:
        if worksheet_name not in available_sheet_names:
            raise _missing_worksheet_error(worksheet_name, available_sheet_names)
        resolved_sheet_name = worksheet_name
    else:
        resolved_sheet_name = _resolve_load_sheet_name(workbook)
    sheet_resolved_at = perf_counter()
    print(
        "BOL parse timing: sheet_detection="
        f"{sheet_resolved_at - workbook_loaded_at:.3f}s sheet={resolved_sheet_name!r}"
    )

    df = _standard_dataframe(workbook.parse(sheet_name=resolved_sheet_name, dtype=object, header=None))
    sheet_loaded_at = perf_counter()
    print(
        "BOL parse timing: sheet_load="
        f"{sheet_loaded_at - sheet_resolved_at:.3f}s rows={len(df)} columns={len(df.columns)}"
    )

    if df.empty:
        raise ValueError(f"Worksheet '{resolved_sheet_name}' contains no rows.")

    try:
        column_map = _resolve_columns(df.columns.tolist(), worksheet_name=resolved_sheet_name, column_overrides=column_overrides)
    except ColumnMappingRequired as exc:
        raise attach_mapping_samples(exc, df.columns, df.head(3).values.tolist())
    header_resolved_at = perf_counter()
    print(
        "BOL parse timing: header_resolution="
        f"{header_resolved_at - sheet_loaded_at:.3f}s"
    )

    kk_load_columns = [column_map["kk_load"]] if (column_overrides or {}).get("kk_load") else _resolve_kk_load_columns(df.columns.tolist())

    parsed_rows = _parse_standard_dataframe_rows(
        df,
        column_map=column_map,
        kk_load_columns=kk_load_columns,
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
