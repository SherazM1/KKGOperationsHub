"""Excel reader for Albertsons carton labels."""

from __future__ import annotations

from typing import Any

import pandas as pd

from app.models.albertsons_label import AlbertsonsLabel


COLUMN_MAP = {
    "ship_to_name": "Buying Party Name",
    "ship_to_address": "Buying Party Address 1",
    "ship_to_city": "Buying Party City",
    "ship_to_state": "Buying Party State",
    "ship_to_zip": "Buying Party Zip",
    "po_number": "Purchase Order Number",
    "item_number": "Item #",
    "description": "Description",
    "quantity": "Qty",
}

REQUIRED_LOGICAL_COLUMNS = {
    "ship_to_name",
    "ship_to_address",
    "ship_to_city",
    "ship_to_state",
    "ship_to_zip",
    "po_number",
    "description",
}


def _normalize_header(header: str) -> str:
    return header.strip().lower()


def _resolve_columns(columns: list[str]) -> dict[str, str]:
    normalized_to_actual = {_normalize_header(col): col for col in columns}
    resolved: dict[str, str] = {}
    missing: list[str] = []

    for logical_name, expected_header in COLUMN_MAP.items():
        normalized_expected = _normalize_header(expected_header)
        if normalized_expected in normalized_to_actual:
            resolved[logical_name] = normalized_to_actual[normalized_expected]
        elif logical_name in REQUIRED_LOGICAL_COLUMNS:
            missing.append(expected_header)

    if missing:
        raise ValueError("Missing required columns: " + ", ".join(missing))

    return resolved


def _coerce_to_string(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def read_excel_albertsons(file: Any) -> list[AlbertsonsLabel]:
    df = pd.read_excel(file, dtype=str)

    if df.empty:
        raise ValueError("Excel file contains no rows.")

    column_map = _resolve_columns(df.columns.tolist())
    labels: list[AlbertsonsLabel] = []

    for index, row in df.iterrows():
        row_number = index + 2

        ship_to_name = _coerce_to_string(row[column_map["ship_to_name"]])
        ship_to_address = _coerce_to_string(row[column_map["ship_to_address"]])
        ship_to_city = _coerce_to_string(row[column_map["ship_to_city"]])
        ship_to_state = _coerce_to_string(row[column_map["ship_to_state"]])
        ship_to_zip = _coerce_to_string(row[column_map["ship_to_zip"]])
        po_number = _coerce_to_string(row[column_map["po_number"]])
        item_number = (
            _coerce_to_string(row[column_map["item_number"]])
            if "item_number" in column_map
            else ""
        )
        description = _coerce_to_string(row[column_map["description"]])
        quantity = (
            _coerce_to_string(row[column_map["quantity"]])
            if "quantity" in column_map
            else ""
        )

        if not any(
            [
                ship_to_name,
                ship_to_address,
                ship_to_city,
                ship_to_state,
                ship_to_zip,
                po_number,
                item_number,
                description,
                quantity,
            ]
        ):
            continue

        if not po_number and not item_number:
            break

        if not po_number:
            raise ValueError(f"Row {row_number}: Purchase Order Number is blank.")

        labels.append(
            AlbertsonsLabel(
                ship_to_name=ship_to_name.split("SUB")[0].strip(),
                ship_to_address=ship_to_address,
                ship_to_city=ship_to_city,
                ship_to_state=ship_to_state,
                ship_to_zip=ship_to_zip,
                po_number=po_number,
                item_number=item_number,
                description=description,
                quantity=quantity,
                dc_label="DC#",
                dc_value="WNCA",
                carton_number="1",
            )
        )

    if not labels:
        raise ValueError("No valid Albertsons label rows found in Excel file.")

    return labels
