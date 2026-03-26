"""Excel reader service for converting uploaded rows into label models."""

from __future__ import annotations

from typing import Any

import pandas as pd

from app.models.label import Label


REQUIRED_COLUMN_MAP = {
    "supplier": ["supplier"],
    "store": ["store", "store #"],
    "po": ["po", "po #"],
    "description": ["description"],
    "sap": ["sap", "sap #"],
}


def _normalize_header(header: str) -> str:
    return header.strip().lower()


def _resolve_columns(columns: list[str]) -> dict[str, str]:
    normalized = {_normalize_header(col): col for col in columns}

    resolved: dict[str, str] = {}

    for logical_name, variations in REQUIRED_COLUMN_MAP.items():
        for variant in variations:
            if variant in normalized:
                resolved[logical_name] = normalized[variant]
                break

        if logical_name not in resolved:
            raise ValueError(
                f"Missing required column for '{logical_name}'. "
                f"Accepted names: {variations}"
            )

    return resolved


def _coerce_to_string(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _normalize_sap(value: str, row_number: int) -> str:
    cleaned = value.strip()

    if not cleaned:
        raise ValueError(f"Row {row_number}: SAP is blank.")

    if not cleaned.isdigit():
        raise ValueError(
            f"Row {row_number}: SAP must be numeric. Got '{value}'."
        )

    length = len(cleaned)

    if length == 10:
        return cleaned

    if length == 9:
        return cleaned.zfill(10)

    raise ValueError(
        f"Row {row_number}: SAP must be 9 or 10 digits. "
        f"Got '{value}' ({length} digits)."
    )


def read_excel(file: Any) -> list[Label]:
    df = pd.read_excel(file, dtype=str)

    if df.empty:
        raise ValueError("Excel file contains no rows.")

    column_map = _resolve_columns(df.columns.tolist())

    labels: list[Label] = []

    for index, row in df.iterrows():
        row_number = index + 2  # Excel row number (header is row 1)

        supplier = _coerce_to_string(row[column_map["supplier"]])
        store = _coerce_to_string(row[column_map["store"]])
        po = _coerce_to_string(row[column_map["po"]])
        description = _coerce_to_string(row[column_map["description"]])
        sap_raw = _coerce_to_string(row[column_map["sap"]])

        # 🔹 Skip completely empty trailing rows
        if not any([supplier, store, po, description, sap_raw]):
            continue

        # 🔹 If partially filled, enforce required fields
        if not supplier:
            raise ValueError(f"Row {row_number}: Supplier is blank.")

        if not store:
            raise ValueError(f"Row {row_number}: Store is blank.")

        if not po:
            raise ValueError(f"Row {row_number}: PO is blank.")

        sap = _normalize_sap(sap_raw, row_number)

        labels.append(
            Label(
                supplier=supplier,
                store=store,
                po=po,
                description=description,
                sap=sap,
            )
        )

    if not labels:
        raise ValueError("No valid label rows found in Excel file.")

    return labels