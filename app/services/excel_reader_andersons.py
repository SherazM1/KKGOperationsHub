"""Excel reader for Andersons labels."""

from __future__ import annotations

from decimal import Decimal, InvalidOperation
from typing import Any

from openpyxl import load_workbook
import pandas as pd

from app.models.andersons_label import AndersonsLabel


ANDERSONS_SHIP_FROM_OPTIONS = {
    "MAD": {
        "care_of": "Mid America",
        "address": "12949 Enterprise Way, Suite 100",
        "city": "Bridgeton",
        "state": "MO",
        "zip_code": "63044",
    },
    "Flower City": {
        "care_of": "Flower City Group",
        "address": "1001 Lee Rd",
        "city": "Rochester",
        "state": "NY",
        "zip_code": "14606",
    },
    "LAMB": {
        "care_of": "Lamb & Associates",
        "address": "1700 Murphy Drive",
        "city": "Maumelle",
        "state": "AR",
        "zip_code": "72113",
    },
    "Duraco": {
        "care_of": "Duraco Specialty Tapes LLC",
        "address": "7400 Industrial Drive",
        "city": "Forest Park",
        "state": "IL",
        "zip_code": "60130",
    },
    "Kinter": {
        "care_of": "Kinter",
        "address": "3333 Oak Grove Ave",
        "city": "Waukegan",
        "state": "IL",
        "zip_code": "60087",
    },
    "RAND": {
        "care_of": "Rand Graphics",
        "address": "2820 S Hoover Rd",
        "city": "Wichita",
        "state": "KS",
        "zip_code": "67215",
    },
    "Veterans": {
        "care_of": "Phenix-Veterans",
        "address": "10430 Argonne Woods Drive",
        "city": "Woodridge",
        "state": "IL",
        "zip_code": "60517",
    },
    "Steward": {
        "care_of": "Kendal King C/O Steward Printing",
        "address": "10775 Sanden Dr",
        "city": "Dallas",
        "state": "TX",
        "zip_code": "75238",
    },
    "Stribling": {
        "care_of": "Stribling Packaging",
        "address": "419 S Lincoln Street",
        "city": "Lowell",
        "state": "AR",
        "zip_code": "72745",
    },
    "Greenbay": {
        "care_of": "Green Bay",
        "address": "5600 S Moorland Road",
        "city": "New Berlin",
        "state": "WI",
        "zip_code": "53151",
    },
    "RRD": {
        "care_of": "RRD",
        "address": "5201 S International Dr",
        "city": "Cudahy",
        "state": "WI",
        "zip_code": "53110",
    },
    "Landall": {
        "care_of": "Landaal Packaging / Westcott Display",
        "address": "3256 Iron Street",
        "city": "Burton",
        "state": "MI",
        "zip_code": "48529",
    },
}


COLUMN_MAP = {
    "client": ["Client"],
    "upc": ["UPC", "UPC #", "UPC#"],
    "brand": ["Brand"],
    "description": ["Description", "Desc"],
    "unit_of_measure": ["Unit of Measure", "UOM", "Unit"],
    "ordered_quantity": ["Ordered Quantity", "Order Qty", "ORDER QTY", "Qty"],
    "po_name": ["PO Name", "PO NAME"],
    "po_number": ["PO Number", "PO NUMBER", "PO #", "PO#"],
}

REQUIRED_LOGICAL_COLUMNS = set(COLUMN_MAP)


def _normalize_header(header: str) -> str:
    return "".join(char for char in header.strip().lower() if char.isalnum())


def _resolve_columns(columns: list[str]) -> dict[str, str]:
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

    return resolved


def _coerce_to_string(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _clean_excel_value(value: Any) -> str:
    text = _coerce_to_string(value)
    if not text:
        return ""

    decimal_candidate = text.replace(",", "")
    if "e" in decimal_candidate.lower() or decimal_candidate.endswith(".0"):
        try:
            decimal_value = Decimal(decimal_candidate)
        except InvalidOperation:
            return text[:-2] if text.endswith(".0") else text

        if decimal_value == decimal_value.to_integral_value():
            return format(decimal_value.quantize(Decimal(1)), "f")
        return format(decimal_value, "f").rstrip("0").rstrip(".")

    return text[:-2] if text.endswith(".0") else text


def _reset_file_pointer(file: Any, position: int | None) -> None:
    if position is not None and hasattr(file, "seek"):
        file.seek(position)


def _current_file_position(file: Any) -> int | None:
    if not hasattr(file, "tell"):
        return None
    try:
        return file.tell()
    except Exception:
        return None


def get_excel_sheet_names(uploaded_file: Any) -> list[str]:
    """Return visible worksheet names in workbook order without consuming the upload."""
    if uploaded_file is None:
        raise ValueError("No file uploaded. Upload an Excel file to inspect worksheets.")

    position = _current_file_position(uploaded_file)
    try:
        if hasattr(uploaded_file, "seek"):
            uploaded_file.seek(0)
        try:
            workbook = load_workbook(uploaded_file, read_only=True, data_only=True)
            return [
                str(worksheet.title)
                for worksheet in workbook.worksheets
                if worksheet.sheet_state == "visible"
            ]
        except Exception:
            _reset_file_pointer(uploaded_file, position)
            try:
                if hasattr(uploaded_file, "seek"):
                    uploaded_file.seek(0)
                workbook = pd.ExcelFile(uploaded_file)
                return [str(sheet_name) for sheet_name in workbook.sheet_names]
            except Exception as exc:
                raise ValueError(
                    "Could not read Excel worksheets. Upload a valid Excel workbook."
                ) from exc
    finally:
        _reset_file_pointer(uploaded_file, position)


def _missing_worksheet_error(
    worksheet_name: str,
    available_sheets: list[str],
) -> ValueError:
    available_text = ", ".join(available_sheets) if available_sheets else "(none)"
    return ValueError(
        f"Worksheet '{worksheet_name}' was not found. Available worksheets: {available_text}."
    )


def parse_andersons_excel(
    file: Any,
    worksheet_name: str | None = None,
) -> list[AndersonsLabel]:
    position = _current_file_position(file)
    if worksheet_name is not None:
        available_sheet_names = get_excel_sheet_names(file)
        if worksheet_name not in available_sheet_names:
            raise _missing_worksheet_error(worksheet_name, available_sheet_names)

    try:
        if hasattr(file, "seek"):
            file.seek(0)
        df = pd.read_excel(
            file,
            dtype=str,
            sheet_name=worksheet_name if worksheet_name is not None else 0,
        )
    finally:
        _reset_file_pointer(file, position)

    if df.empty:
        raise ValueError("Excel file contains no rows.")

    column_map = _resolve_columns(df.columns.tolist())
    labels: list[AndersonsLabel] = []

    for index, row in df.iterrows():
        row_number = index + 2

        client = _clean_excel_value(row[column_map["client"]])
        upc = _clean_excel_value(row[column_map["upc"]])
        brand = _clean_excel_value(row[column_map["brand"]])
        description = _clean_excel_value(row[column_map["description"]])
        unit_of_measure = _clean_excel_value(row[column_map["unit_of_measure"]])
        ordered_quantity = _clean_excel_value(row[column_map["ordered_quantity"]])
        po_name = _clean_excel_value(row[column_map["po_name"]])
        po_number = _clean_excel_value(row[column_map["po_number"]])

        if not any(
            [
                client,
                upc,
                brand,
                description,
                unit_of_measure,
                ordered_quantity,
                po_name,
                po_number,
            ]
        ):
            continue

        required_values = {
            "Client": client,
            "UPC": upc,
            "Description": description,
            "Unit of Measure": unit_of_measure,
            "Ordered Quantity": ordered_quantity,
            "PO Name": po_name,
            "PO Number": po_number,
        }
        blank_required = [name for name, value in required_values.items() if not value]
        if blank_required:
            raise ValueError(
                f"Row {row_number}: Missing required value(s): "
                + ", ".join(blank_required)
            )

        labels.append(
            AndersonsLabel(
                client=client,
                upc=upc,
                brand=brand,
                description=description,
                unit_of_measure=unit_of_measure,
                ordered_quantity=ordered_quantity,
                po_name=po_name,
                po_number=po_number,
            )
        )

    if not labels:
        raise ValueError("No valid Andersons label rows found in Excel file.")

    return labels


def read_excel_andersons(
    file: Any,
    worksheet_name: str | None = None,
) -> list[AndersonsLabel]:
    return parse_andersons_excel(file, worksheet_name=worksheet_name)
