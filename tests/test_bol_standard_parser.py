from __future__ import annotations

from io import BytesIO

import pandas as pd
import pytest

from app.services.bol_standard_parser import parse_standard_bol_excel
from app.services.bol_standard_mapper import map_standard_rows_to_records


def _standard_load_row() -> dict[str, object]:
    return {
        "KK Load": "KL-001",
        "Carrier": "Test Carrier",
        "load#": "LOAD-001",
        "KK PO#": "KKPO-001",
        "BOL #": "BOL-001",
        "ship date": "2026-05-13",
        "DC #": "1234",
        "DC NAME": "Test DC",
        "DC STREET": "123 Test St",
        "DC CITY": "Dallas",
        "DC ST": "TX",
        "DC ZIP": "75001",
        "TGT PO #": "TGT-001",
        "ITEM #": "ITEM-001",
        "UPC": "000111222333",
        "Pallet Description": "Test pallet",
        "QTY": 10,
        "TOTAL PALLETS": 2,
        "weight each": 50,
        "Total Weight": 100,
    }


def _workbook_with_sheet(sheet_name: str, rows: list[dict[str, object]]) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output


def test_parse_standard_bol_excel_accepts_normalized_main_load_sheet_name() -> None:
    workbook = _workbook_with_sheet("MAIN LOAD SHEET ", [_standard_load_row()])

    rows = parse_standard_bol_excel(workbook)

    assert len(rows) == 1
    assert rows[0].bol_number == "BOL-001"


def test_parse_standard_bol_excel_accepts_load_sheet_with_trailing_space() -> None:
    workbook = _workbook_with_sheet("LOAD SHEET ", [_standard_load_row()])

    rows = parse_standard_bol_excel(workbook)

    assert len(rows) == 1
    assert rows[0].wm_po == "TGT-001"
    assert rows[0].dc_city_state_zip == "Dallas, TX 75001"


def test_parse_standard_bol_excel_falls_back_to_header_based_sheet_detection() -> None:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame({"Other": ["not load data"]}).to_excel(
            writer,
            sheet_name="RATES",
            index=False,
        )
        pd.DataFrame([_standard_load_row()]).to_excel(
            writer,
            sheet_name="Operational Data",
            index=False,
        )
    output.seek(0)

    rows = parse_standard_bol_excel(output)

    assert len(rows) == 1
    assert rows[0].kk_load == "LOAD-001"


def test_parse_standard_bol_excel_normalizes_line_break_headers() -> None:
    row = _standard_load_row()
    row["DC\nST"] = row.pop("DC ST")
    row["DC\nZIP"] = row.pop("DC ZIP")
    row["Pallet\nDescription"] = row.pop("Pallet Description")

    workbook = _workbook_with_sheet("LOAD SHEET", [row])

    rows = parse_standard_bol_excel(workbook)

    assert len(rows) == 1
    assert rows[0].dc_city_state_zip == "Dallas, TX 75001"
    assert rows[0].item_description == "Test pallet"


def test_parse_standard_bol_excel_rejects_invalid_workbook_with_clear_error() -> None:
    workbook = _workbook_with_sheet("RATES", [{"Other": "value"}])

    with pytest.raises(ValueError, match="Could not find a BOL load sheet"):
        parse_standard_bol_excel(workbook)


def test_parse_standard_bol_excel_prefers_populated_bol_number_over_po_fallback() -> None:
    row = _standard_load_row()
    row["BOL #"] = "BOL-PRIMARY"
    row["TGT PO #"] = "TGT-FALLBACK"

    workbook = _workbook_with_sheet("LOAD SHEET", [row])

    rows = parse_standard_bol_excel(workbook)

    assert len(rows) == 1
    assert rows[0].bol_number == "BOL-PRIMARY"
    assert rows[0].wm_po == "TGT-FALLBACK"


def test_parse_standard_bol_excel_uses_tgt_po_as_effective_bol_when_bol_blank() -> None:
    row = _standard_load_row()
    row["BOL #"] = ""
    row["TGT PO #"] = "10001859231-0551"

    workbook = _workbook_with_sheet("LOAD SHEET", [row])

    rows = parse_standard_bol_excel(workbook)

    assert len(rows) == 1
    assert rows[0].bol_number == "10001859231-0551"
    assert rows[0].wm_po == "10001859231-0551"


def test_standard_bol_mapping_groups_blank_bol_rows_by_effective_po_fallback() -> None:
    first_row = _standard_load_row()
    first_row["BOL #"] = ""
    first_row["TGT PO #"] = "10001859231-0551"
    first_row["load#"] = "LOAD-0551"
    first_row["DC #"] = "0551"
    first_row["DC NAME"] = "DC 0551"

    second_row = _standard_load_row()
    second_row["BOL #"] = ""
    second_row["TGT PO #"] = "10001859231-0553"
    second_row["load#"] = "LOAD-0553"
    second_row["DC #"] = "0553"
    second_row["DC NAME"] = "DC 0553"

    workbook = _workbook_with_sheet("LOAD SHEET", [first_row, second_row])

    rows = parse_standard_bol_excel(workbook)
    records = map_standard_rows_to_records(rows)

    assert [record.bol_number for record in records] == [
        "10001859231-0551",
        "10001859231-0553",
    ]
    assert all(record.status == "Ready" for record in records)


def test_standard_bol_mapping_keeps_missing_bol_when_bol_and_po_are_blank() -> None:
    row = _standard_load_row()
    row["BOL #"] = ""
    row["TGT PO #"] = ""

    workbook = _workbook_with_sheet("LOAD SHEET", [row])

    rows = parse_standard_bol_excel(workbook)
    records = map_standard_rows_to_records(rows)

    assert len(records) == 1
    assert records[0].bol_number == ""
    assert "BOL #" in records[0].missing_required_fields
