"""Tests for Andersons Excel parsing."""

from __future__ import annotations

from io import BytesIO

from openpyxl import load_workbook
import pandas as pd
import pytest

from app.services.excel_reader_andersons import (
    get_excel_sheet_names,
    parse_andersons_excel,
)


def _base_row(**overrides: str) -> dict[str, str]:
    row = {
        "Client": "Andersons",
        "UPC": "123456789012",
        "Brand": "Brand A",
        "Description": "Display",
        "Unit of Measure": "Case",
        "Ordered Quantity": "12",
        "PO Name": "Spring Set",
        "PO Number": "PO-1",
    }
    row.update(overrides)
    return row


def _workbook_with_sheets(sheets: list[tuple[str, list[dict[str, str]]]]) -> BytesIO:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, rows in sheets:
            pd.DataFrame(rows).to_excel(writer, sheet_name=sheet_name, index=False)
    buffer.seek(0)
    return buffer


def test_get_excel_sheet_names_returns_visible_sheets_in_workbook_order() -> None:
    workbook = _workbook_with_sheets(
        [
            ("Revised LS", [_base_row()]),
            ("Rates", [_base_row()]),
            ("Load Sheet", [_base_row()]),
        ]
    )
    excel_workbook = load_workbook(workbook)
    excel_workbook["Rates"].sheet_state = "hidden"
    workbook = BytesIO()
    excel_workbook.save(workbook)
    workbook.seek(0)

    sheet_names = get_excel_sheet_names(workbook)

    assert sheet_names == ["Revised LS", "Load Sheet"]
    assert workbook.tell() == 0


def test_parse_andersons_excel_parses_explicit_worksheet() -> None:
    workbook = _workbook_with_sheets(
        [
            ("Load Sheet", [_base_row(**{"PO Number": "PO-A"})]),
            ("Revised LS", [_base_row(**{"PO Number": "PO-B"})]),
        ]
    )

    labels = parse_andersons_excel(workbook, worksheet_name="Revised LS")

    assert len(labels) == 1
    assert labels[0].po_number == "PO-B"


def test_parse_andersons_excel_explicit_worksheet_overrides_default_sheet() -> None:
    workbook = _workbook_with_sheets(
        [
            ("Load Sheet", [_base_row(**{"PO Number": "PO-DEFAULT"})]),
            ("Revised LS", [_base_row(**{"PO Number": "PO-SELECTED"})]),
        ]
    )

    labels = parse_andersons_excel(workbook, worksheet_name="Revised LS")

    assert labels[0].po_number == "PO-SELECTED"


def test_parse_andersons_excel_missing_explicit_worksheet_lists_available_sheets() -> None:
    workbook = _workbook_with_sheets(
        [
            ("Revised LS", [_base_row()]),
            ("Load Sheet", [_base_row()]),
        ]
    )

    with pytest.raises(
        ValueError,
        match=r"Worksheet 'Missing' was not found\. Available worksheets: Revised LS, Load Sheet\.",
    ):
        parse_andersons_excel(workbook, worksheet_name="Missing")


def test_parse_andersons_excel_keeps_default_behavior_when_worksheet_is_none() -> None:
    workbook = _workbook_with_sheets(
        [
            ("First", [_base_row(**{"PO Number": "PO-FIRST"})]),
            ("Revised LS", [_base_row(**{"PO Number": "PO-REVISED"})]),
        ]
    )

    labels = parse_andersons_excel(workbook, worksheet_name=None)

    assert labels[0].po_number == "PO-FIRST"
