"""Excel reader for Sam's warehouse 4x6 labels."""

from __future__ import annotations

from typing import Any

import pandas as pd

from app.models.sams_label import SamsLabel


REQUIRED_COLUMNS = [
    "SHIPPER NAME",
    "SHIPPER ADDRESS",
    "SHIPPER CITY",
    "SHIPPER STATE",
    "SHIPPER ZIP",
    "SHIP TO NAME",
    "SHIP TO ADDRESS",
    "CITY",
    "STATE",
    "ZIP",
    "PO #",
    "QTY",
    "UPC",
    "WHSE",
    "TYPE",
    "DEPT",
    "Item #",
    "Desc",
]

HEADER_ALIASES = {
    "shipper adress": "shipper address",
    "ship to adress": "ship to address",
    "ship to adderss": "ship to address",  
}

def _normalize_header(header: str) -> str:
    cleaned = " ".join(header.strip().lower().split())
    return HEADER_ALIASES.get(cleaned, cleaned)

def _resolve_columns(columns: list[str]) -> dict[str, str]:
    normalized_to_actual = {_normalize_header(col): col for col in columns}
    resolved: dict[str, str] = {}
    missing: list[str] = []

    for required in REQUIRED_COLUMNS:
        key = _normalize_header(required)
        if key in normalized_to_actual:
            resolved[required] = normalized_to_actual[key]
        else:
            missing.append(required)

    if missing:
        raise ValueError(
            "Missing required columns: " + ", ".join(missing)
        )

    return resolved


def _coerce_to_string(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _validate_shipper_zip(zip_value: str, row_number: int) -> str:
    if not zip_value:
        raise ValueError(f"Row {row_number}: SHIPPER ZIP is blank.")

    if not zip_value.isdigit() or len(zip_value) != 5:
        raise ValueError(
            f"Row {row_number}: SHIPPER ZIP must be 5 digits."
        )

    return zip_value


def _validate_ship_to_zip(zip_value: str, row_number: int) -> str:
    if not zip_value:
        raise ValueError(f"Row {row_number}: ZIP is blank.")

    if "-" not in zip_value:
        raise ValueError(
            f"Row {row_number}: ZIP must be 5+4 format (#####-####)."
        )

    digits_only = zip_value.replace("-", "")
    if len(digits_only) != 9 or not digits_only.isdigit():
        raise ValueError(
            f"Row {row_number}: ZIP must contain 9 digits plus a dash."
        )

    return zip_value


def _validate_upc(upc: str, row_number: int) -> str:
    if not upc:
        raise ValueError(f"Row {row_number}: UPC is blank.")

    if not upc.isdigit():
        raise ValueError(
            f"Row {row_number}: UPC must be numeric. Got '{upc}'."
        )

    return upc


def read_excel_sams(file: Any) -> list[SamsLabel]:
    df = pd.read_excel(file, sheet_name=0, dtype=str)

    if df.empty:
        raise ValueError("Excel file contains no rows.")

    column_map = _resolve_columns(df.columns.tolist())

    labels: list[SamsLabel] = []

    for index, row in df.iterrows():
        row_number = index + 2
        values = {
            col: _coerce_to_string(row[column_map[col]])
            for col in REQUIRED_COLUMNS
        }

        if not any(values.values()):
            continue

        po_number = values["PO #"]
        upc = values["UPC"]

        if not po_number and not upc:
            break

        shipper_zip = _validate_shipper_zip(values["SHIPPER ZIP"], row_number)
        ship_to_zip = _validate_ship_to_zip(values["ZIP"], row_number)
        validated_upc = _validate_upc(upc, row_number)

        labels.append(
            SamsLabel(
                shipper_name=values["SHIPPER NAME"],
                shipper_address=values["SHIPPER ADDRESS"],
                shipper_city=values["SHIPPER CITY"],
                shipper_state=values["SHIPPER STATE"],
                shipper_zip=shipper_zip,
                ship_to_name=values["SHIP TO NAME"],
                ship_to_address=values["SHIP TO ADDRESS"],
                ship_to_city=values["CITY"],
                ship_to_state=values["STATE"],
                ship_to_zip=ship_to_zip,
                po_number=po_number,
                quantity=values["QTY"],
                upc=validated_upc,
                whse=values["WHSE"],
                type_code=values["TYPE"],
                dept=values["DEPT"],
                item_number=values["Item #"],
                description=values["Desc"],
            )
        )

    if not labels:
        raise ValueError("No valid Sam's label rows found in Excel file.")

    return labels
