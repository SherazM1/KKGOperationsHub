"""Tests for Sam's GCI Excel parsing."""

from __future__ import annotations

from io import BytesIO

import pandas as pd

from app.services.excel_reader_sams_gci import read_excel_sams_gci


def _workbook(rows: list[dict[str, str]]) -> BytesIO:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(writer, index=False)
    buffer.seek(0)
    return buffer


def test_read_excel_sams_gci_accepts_current_mdg_template_headers() -> None:
    mdg_file = _workbook(
        [
            {
                "SHIPPER NAME": "KKG",
                "SHIPPER ADDRESS": "123 Main",
                "SHIPPER CITY": "Green Bay",
                "SHIPPER STATE": "WI",
                "SHIPPER ZIP": "54301",
                "Item #": "1001",
                "Desc": "Display",
                "CLUB#": "6612",
                "WHSE": "WH1",
                "SHIP TO NAME": "Sam's Club",
                "SHIP TO ADDERSS": "456 Club",
                "CITY": "Bentonville",
                "STATE ": "AR",
                "ZIP": "72712",
                "PO #": "PO-1",
                "QTY": "24",
            }
        ]
    )
    gci_file = _workbook(
        [
            {
                "PROGRAM NAME": "Program A",
                "ITEM #": "1001",
                "QTY": "24",
                "UPC": "123456789012",
                "DESC": "Display",
            }
        ]
    )

    payload = read_excel_sams_gci(mdg_file, gci_file)

    assert payload.mdg_labels[0].shipper_address == "123 Main"
    assert payload.mdg_labels[0].ship_to_address == "456 Club"
    assert payload.bottom_rows[0].barcode_value == "123456789012"


def test_read_excel_sams_gci_accepts_loose_header_variations() -> None:
    mdg_file = _workbook(
        [
            {
                "Shipper": "KKG",
                "SHIPPER ADDRERSS": "123 Main",
                "SHIPPER_CITY": "Green Bay",
                "SHIPPER_STATE": "WI",
                "SHIPPER_POSTAL_CODE": "54301",
                "ShipTo Name": "Sam's Club",
                "SHIPTO_ADDRERSS": "456 Club",
                "SHIPTO_CITY": "Bentonville",
                "SHIPTO_STATE": "AR",
                "SHIPTO_ZIP_CODE": "72712",
                "P.O. #": "PO-1",
                "Club No": "6612",
                "Warehouse #": "WH1",
                "Item Num": "1001",
                "Item Description": "Display",
                "Ordered Quantity": "24",
            }
        ]
    )
    gci_file = _workbook(
        [
            {
                "PROGRAM_NAME": "Program A",
                "ITEM_NUM": "1001",
                "ORDER_QTY": "24",
                "UPC/BARCODE": "123456789012",
                "Product Description": "Display",
            }
        ]
    )

    payload = read_excel_sams_gci(mdg_file, gci_file)

    assert payload.mdg_labels[0].shipper_name == "KKG"
    assert payload.mdg_labels[0].shipper_address == "123 Main"
    assert payload.mdg_labels[0].po_number == "PO-1"
    assert payload.bottom_rows[0].program_name == "Program A"
    assert payload.bottom_rows[0].barcode_value == "123456789012"


def test_read_excel_sams_gci_accepts_second_workbook_without_program_or_qty_header() -> None:
    mdg_file = _workbook(
        [
            {
                "SHIPPER NAME": "KKG",
                "SHIPPER ADDRESS": "123 Main",
                "SHIPPER CITY": "Green Bay",
                "SHIPPER STATE": "WI",
                "SHIPPER ZIP": "54301",
                "SHIP TO NAME": "Sam's Club",
                "SHIP TO ADDRESS": "456 Club",
                "CITY": "Bentonville",
                "STATE": "AR",
                "ZIP": "72712",
                "PO #": "PO-1",
                "CLUB#": "6612",
                "WHSE": "WH1",
                "ITEM #": "1001",
                "DESC": "Display",
                "QTY": "24",
            }
        ]
    )
    gci_file = _workbook(
        [
            {
                "SAM'S ITEM #": "990553354",
                "UPC": "849219096356",
                "ITEM DESCRIPTION": "2Pk Boxes Trees",
                "PER DISPLAY": "12",
            }
        ]
    )

    payload = read_excel_sams_gci(mdg_file, gci_file)

    assert payload.bottom_rows[0].program_name == ""
    assert payload.bottom_rows[0].item_number == "990553354"
    assert payload.bottom_rows[0].barcode_value == "849219096356"
    assert payload.bottom_rows[0].description == "2Pk Boxes Trees"
    assert payload.bottom_rows[0].quantity == "12"
