"""Excel reader for Sam's GCI label workflow."""

from __future__ import annotations

import re
from typing import Any

import pandas as pd

from app.models.sams_gci_label import (
    SamsGciBottomRow,
    SamsGciPayload,
    SamsGciTopLabelRow,
)


MDG_FILE_DISPLAY_NAME = "SAMS MDG Label template.xlsx"
GCI_FILE_DISPLAY_NAME = "Sams PO Labels with GCI.xlsx"


MDG_REQUIRED_FIELDS: dict[str, str] = {
    "shipper_name": "shipper name",
    "shipper_address": "shipper address",
    "shipper_city": "shipper city",
    "shipper_state": "shipper state",
    "shipper_zip": "shipper zip",
    "ship_to_name": "ship to name",
    "ship_to_address": "ship to address",
    "ship_to_city": "ship to city",
    "ship_to_state": "ship to state",
    "ship_to_zip": "ship to zip",
    "po_number": "po number",
    "club_number": "club number",
    "whse": "whse",
    "item_number": "item number",
    "description": "description",
    "quantity": "qty",
}

MDG_FIELD_ALIASES: dict[str, list[str]] = {
    "shipper_name": ["shipper name", "shipper", "shipper_name"],
    "shipper_address": [
        "shipper address",
        "shipper adress",
        "shipper adderss",
        "shipper addrerss",
        "shipper addr",
        "shipper street",
        "shipper street address",
        "shipper_address",
    ],
    "shipper_city": ["shipper city", "shipper_city"],
    "shipper_state": ["shipper state", "shipper_state"],
    "shipper_zip": [
        "shipper zip",
        "shipper zip code",
        "shipper postal code",
        "shipper_zip",
    ],
    "ship_to_name": ["ship to name", "shipto name", "ship_to_name"],
    "ship_to_address": [
        "ship to address",
        "ship to adderss",
        "ship to adress",
        "ship to addrerss",
        "ship to addr",
        "ship to street",
        "ship to street address",
        "shipto address",
        "shipto adderss",
        "shipto adress",
        "shipto addrerss",
        "shipto addr",
        "ship_to_address",
    ],
    "ship_to_city": ["city", "ship to city", "shipto city", "ship_to_city"],
    "ship_to_state": ["state", "ship to state", "shipto state", "ship_to_state"],
    "ship_to_zip": [
        "zip",
        "zip code",
        "postal code",
        "ship to zip",
        "ship to zip code",
        "ship to postal code",
        "shipto zip",
        "shipto zip code",
        "shipto postal code",
        "ship_to_zip",
    ],
    "po_number": [
        "po #",
        "po#",
        "po number",
        "po no",
        "po num",
        "po",
        "p o",
        "po_number",
    ],
    "club_number": [
        "club#",
        "club #",
        "club",
        "club number",
        "club no",
        "club num",
        "club_number",
    ],
    "whse": [
        "whse",
        "warehouse",
        "whse #",
        "whse#",
        "warehouse #",
        "warehouse number",
    ],
    "item_number": [
        "item #",
        "item#",
        "item number",
        "item no",
        "item num",
        "item",
        "item_number",
    ],
    "description": ["desc", "description", "item description", "product description"],
    "quantity": ["qty", "quantity", "qty ordered", "order qty", "ordered quantity"],
}


GCI_REQUIRED_FIELDS: dict[str, str] = {
    "program_name": "program name",
    "item_number": "item number",
    "quantity": "qty",
    "barcode_value": "barcode value / upc",
    "description": "description",
}

GCI_FIELD_ALIASES: dict[str, list[str]] = {
    "program_name": ["program name", "program", "program_name"],
    "item_number": [
        "item #",
        "item#",
        "item number",
        "item no",
        "item num",
        "item",
        "item_number",
    ],
    "quantity": ["qty", "quantity", "qty ordered", "order qty", "ordered quantity"],
    "barcode_value": [
        "barcode",
        "barcode value",
        "barcode value / upc",
        "barcode upc",
        "upc barcode",
        "upc",
        "upc code",
        "upc #",
        "upc#",
        "upc number",
    ],
    "description": ["description", "desc", "item description", "product description"],
}


def _normalize_header(header: str) -> str:
    value = str(header or "").strip().lower()
    value = value.replace("_", " ")
    value = re.sub(r"[^\w\s]", " ", value)
    value = " ".join(value.split())
    return value


def _coerce_to_string(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _resolve_columns(
    columns: list[str],
    aliases: dict[str, list[str]],
    required_fields: dict[str, str],
    file_display_name: str,
) -> dict[str, str]:
    normalized_to_actual: dict[str, str] = {}
    for col in columns:
        normalized = _normalize_header(col)
        if normalized and normalized not in normalized_to_actual:
            normalized_to_actual[normalized] = col

    resolved: dict[str, str] = {}
    missing: list[str] = []

    for logical_field, alias_list in aliases.items():
        match = None
        for alias in alias_list:
            normalized_alias = _normalize_header(alias)
            if normalized_alias in normalized_to_actual:
                match = normalized_to_actual[normalized_alias]
                break
        if match is None:
            missing.append(required_fields[logical_field])
            continue
        resolved[logical_field] = match

    if missing:
        raise ValueError(
            f"Missing required columns in {file_display_name}: {', '.join(missing)}"
        )

    return resolved


def _parse_mdg_rows(mdg_df: pd.DataFrame, column_map: dict[str, str]) -> list[SamsGciTopLabelRow]:
    labels: list[SamsGciTopLabelRow] = []

    for _, row in mdg_df.iterrows():
        values = {
            field: _coerce_to_string(row[column_map[field]])
            for field in MDG_REQUIRED_FIELDS
        }

        if not any(values.values()):
            continue

        labels.append(
            SamsGciTopLabelRow(
                shipper_name=values["shipper_name"],
                shipper_address=values["shipper_address"],
                shipper_city=values["shipper_city"],
                shipper_state=values["shipper_state"],
                shipper_zip=values["shipper_zip"],
                ship_to_name=values["ship_to_name"],
                ship_to_address=values["ship_to_address"],
                ship_to_city=values["ship_to_city"],
                ship_to_state=values["ship_to_state"],
                ship_to_zip=values["ship_to_zip"],
                po_number=values["po_number"],
                club_number=values["club_number"],
                whse=values["whse"],
                item_number=values["item_number"],
                description=values["description"],
                quantity=values["quantity"],
            )
        )

    if not labels:
        raise ValueError(f"No valid MDG rows found in {MDG_FILE_DISPLAY_NAME}.")

    return labels


def _parse_gci_bottom_rows(
    gci_df: pd.DataFrame, column_map: dict[str, str]
) -> list[SamsGciBottomRow]:
    bottom_rows: list[SamsGciBottomRow] = []

    for _, row in gci_df.iterrows():
        values = {
            field: _coerce_to_string(row[column_map[field]])
            for field in GCI_REQUIRED_FIELDS
        }

        if not any(values.values()):
            continue

        bottom_rows.append(
            SamsGciBottomRow(
                program_name=values["program_name"],
                item_number=values["item_number"],
                quantity=values["quantity"],
                barcode_value=values["barcode_value"],
                description=values["description"],
            )
        )

    if not bottom_rows:
        raise ValueError(f"No valid bottom rows found in {GCI_FILE_DISPLAY_NAME}.")

    return bottom_rows


def read_excel_sams_gci(mdg_file: Any, gci_file: Any) -> SamsGciPayload:
    """Parse Sam's GCI MDG top rows and shared bottom rows from two workbooks."""
    try:
        mdg_df = pd.read_excel(mdg_file, sheet_name=0, dtype=str)
    except Exception as exc:
        raise ValueError(f"Unable to read {MDG_FILE_DISPLAY_NAME}: {exc}") from exc

    try:
        gci_df = pd.read_excel(gci_file, sheet_name=0, dtype=str)
    except Exception as exc:
        raise ValueError(f"Unable to read {GCI_FILE_DISPLAY_NAME}: {exc}") from exc

    if mdg_df.empty:
        raise ValueError(f"{MDG_FILE_DISPLAY_NAME} contains no rows.")
    if gci_df.empty:
        raise ValueError(f"{GCI_FILE_DISPLAY_NAME} contains no rows.")

    mdg_column_map = _resolve_columns(
        mdg_df.columns.tolist(),
        aliases=MDG_FIELD_ALIASES,
        required_fields=MDG_REQUIRED_FIELDS,
        file_display_name=MDG_FILE_DISPLAY_NAME,
    )
    gci_column_map = _resolve_columns(
        gci_df.columns.tolist(),
        aliases=GCI_FIELD_ALIASES,
        required_fields=GCI_REQUIRED_FIELDS,
        file_display_name=GCI_FILE_DISPLAY_NAME,
    )

    mdg_labels = _parse_mdg_rows(mdg_df, mdg_column_map)
    bottom_rows = _parse_gci_bottom_rows(gci_df, gci_column_map)

    return SamsGciPayload(
        mdg_labels=mdg_labels,
        bottom_rows=bottom_rows,
    )
