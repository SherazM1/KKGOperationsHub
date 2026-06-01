"""Excel reader for Albertsons carton labels."""

from __future__ import annotations

from decimal import Decimal, InvalidOperation
from typing import Any

import pandas as pd

from app.models.albertsons_label import AlbertsonsLabel


COLUMN_MAP = {
    "ship_to_name": ["Buying Party Name"],
    "ship_to_address": ["Buying Party Address 1"],
    "ship_to_city": ["Buying Party City"],
    "ship_to_state": ["Buying Party State"],
    "ship_to_zip": ["Buying Party Zip"],
    "po_number": ["Purchase Order Number"],
    "item_number": ["Item #"],
    "upc": ["UPC #", "UPC#", "UPC", "Upc", "upc #", "upc"],
    "description": ["Description"],
    "quantity": ["Quantity", "Qty", "QTY", "Qty."],
    "dc_value": [
        "DC CODE",
        "DC Code",
        "DC code",
        "DCCODE",
        "DC_CODE",
        "DC #",
        "DC#",
        "DC",
        "Distribution Center Code",
        "Distribution Center",
    ],
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
    return "".join(char for char in header.strip().lower() if char.isalnum())


def _resolve_columns(
    columns: list[str],
    require_quantity: bool = False,
    require_upc: bool = False,
) -> dict[str, str]:
    normalized_to_actual = {_normalize_header(col): col for col in columns}
    resolved: dict[str, str] = {}
    missing: list[str] = []

    for logical_name, expected_headers in COLUMN_MAP.items():
        for expected_header in expected_headers:
            normalized_expected = _normalize_header(expected_header)
            if normalized_expected in normalized_to_actual:
                resolved[logical_name] = normalized_to_actual[normalized_expected]
                break
        else:
            if logical_name in REQUIRED_LOGICAL_COLUMNS:
                missing.append(expected_headers[0])

    if missing:
        raise ValueError("Missing required columns: " + ", ".join(missing))

    if require_quantity and "quantity" not in resolved:
        raise ValueError("Auto Qty requires a Quantity/Qty column in the Albertsons Excel file.")

    if require_upc and "upc" not in resolved:
        raise ValueError(
            "UPC # from Excel mode requires a UPC # column in the Albertsons Excel file."
        )

    return resolved


def _coerce_to_string(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _normalize_upc(value: Any) -> str:
    text = _coerce_to_string(value)
    if not text:
        return ""

    if "e" in text.lower():
        try:
            decimal_value = Decimal(text)
        except InvalidOperation:
            return text
        normalized = format(decimal_value, "f")
        return normalized[:-2] if normalized.endswith(".0") else normalized

    if text.endswith(".0"):
        return text[:-2]

    return text


def read_excel_albertsons(
    file: Any,
    require_quantity: bool = False,
    require_upc: bool = False,
) -> list[AlbertsonsLabel]:
    df = pd.read_excel(file, dtype=str)

    if df.empty:
        raise ValueError("Excel file contains no rows.")

    column_map = _resolve_columns(
        df.columns.tolist(),
        require_quantity=require_quantity,
        require_upc=require_upc,
    )
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
        upc = _normalize_upc(row[column_map["upc"]]) if "upc" in column_map else ""
        description = _coerce_to_string(row[column_map["description"]])
        quantity = (
            _coerce_to_string(row[column_map["quantity"]])
            if "quantity" in column_map
            else ""
        )
        dc_value = (
            _coerce_to_string(row[column_map["dc_value"]])
            if "dc_value" in column_map
            else "WNCA"
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
                upc,
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
                upc=upc,
                description=description,
                quantity=quantity,
                dc_label="DC#",
                dc_value=dc_value or "WNCA",
                carton_number="1",
            )
        )

    if not labels:
        raise ValueError("No valid Albertsons label rows found in Excel file.")

    return labels
