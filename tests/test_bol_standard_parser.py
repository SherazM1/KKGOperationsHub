from __future__ import annotations

from io import BytesIO
from time import perf_counter

import pandas as pd
import pytest
from openpyxl import load_workbook

from app.services.bol_standard_parser import get_excel_sheet_names, parse_standard_bol_excel
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


def _workbook_with_sheets(sheets: list[tuple[str, list[dict[str, object]]]]) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, rows in sheets:
            pd.DataFrame(rows).to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output


def test_get_excel_sheet_names_returns_visible_sheets_in_workbook_order() -> None:
    workbook = _workbook_with_sheets(
        [
            ("Revised LS", [_standard_load_row()]),
            ("Load Sheet", [_standard_load_row()]),
            ("Rates", [{"Rate": "1.00"}]),
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


def test_parse_standard_bol_excel_accepts_normalized_main_load_sheet_name() -> None:
    workbook = _workbook_with_sheet("MAIN LOAD SHEET ", [_standard_load_row()])

    rows = parse_standard_bol_excel(workbook)

    assert len(rows) == 1
    assert rows[0].bol_number == "BOL-001"
    assert rows[0].total_weight == "100"
    assert rows[0].carrier_pro_number == "LOAD-001"


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
    assert rows[0].kk_load == "KL-001"


def test_parse_standard_bol_excel_can_parse_explicit_worksheet() -> None:
    selected_row = _standard_load_row()
    selected_row["BOL #"] = "BOL-REVISED"
    workbook = _workbook_with_sheets(
        [
            ("Load Sheet", [_standard_load_row()]),
            ("Revised LS", [selected_row]),
        ]
    )

    rows = parse_standard_bol_excel(workbook, worksheet_name="Revised LS")

    assert len(rows) == 1
    assert rows[0].bol_number == "BOL-REVISED"


def test_parse_standard_bol_excel_explicit_worksheet_overrides_auto_detect() -> None:
    load_sheet_row = _standard_load_row()
    load_sheet_row["BOL #"] = "BOL-AUTO"
    revised_row = _standard_load_row()
    revised_row["BOL #"] = "BOL-SELECTED"
    workbook = _workbook_with_sheets(
        [
            ("Load Sheet", [load_sheet_row]),
            ("Revised LS", [revised_row]),
        ]
    )

    rows = parse_standard_bol_excel(workbook, worksheet_name="Revised LS")

    assert [row.bol_number for row in rows] == ["BOL-SELECTED"]


def test_parse_standard_bol_excel_missing_explicit_worksheet_lists_available_sheets() -> None:
    workbook = _workbook_with_sheets(
        [
            ("Revised LS", [_standard_load_row()]),
            ("Load Sheet", [_standard_load_row()]),
        ]
    )

    with pytest.raises(
        ValueError,
        match=r"Worksheet 'Missing' was not found\. Available worksheets: Revised LS, Load Sheet\.",
    ):
        parse_standard_bol_excel(workbook, worksheet_name="Missing")


def test_parse_standard_bol_excel_keeps_auto_detect_when_worksheet_is_none() -> None:
    auto_row = _standard_load_row()
    auto_row["BOL #"] = "BOL-AUTO"
    revised_row = _standard_load_row()
    revised_row["BOL #"] = "BOL-REVISED"
    workbook = _workbook_with_sheets(
        [
            ("Load Sheet", [auto_row]),
            ("Revised LS", [revised_row]),
        ]
    )

    rows = parse_standard_bol_excel(workbook, worksheet_name=None)

    assert [row.bol_number for row in rows] == ["BOL-AUTO"]


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


def test_standard_bol_mapping_uses_kk_load_for_kkg_load_when_bol_falls_back_to_tgt_po() -> None:
    row = _standard_load_row()
    row["KK Load"] = "1"
    row["load#"] = "1073839"
    row["BOL #"] = ""
    row["TGT PO #"] = "10001859231-0551"

    workbook = _workbook_with_sheet("LOAD SHEET", [row])

    rows = parse_standard_bol_excel(workbook)
    records = map_standard_rows_to_records(rows)

    assert len(records) == 1
    assert records[0].bol_number == "10001859231-0551"
    assert records[0].kk_load_number == "1"
    assert records[0].carrier_pro_number == "1073839"


def test_parse_standard_bol_excel_does_not_use_load_number_as_kk_load() -> None:
    row = _standard_load_row()
    row.pop("KK Load")
    row["load#"] = "BASE-LOAD-001"

    workbook = _workbook_with_sheet("MAIN LOAD SHEET", [row])

    with pytest.raises(ValueError, match="kk_load"):
        parse_standard_bol_excel(workbook)


def test_parse_standard_bol_excel_uses_next_load_source_when_dedicated_load_is_blank() -> None:
    row = _standard_load_row()
    row["KKG Load #"] = ""
    row["KK Load"] = "1073839"
    row["load#"] = "CARRIER-LOAD-001"

    workbook = _workbook_with_sheet("LOAD SHEET", [row])

    rows = parse_standard_bol_excel(workbook)

    assert len(rows) == 1
    assert rows[0].kk_load == "1073839"


def test_parse_standard_bol_excel_uses_qty_as_plt_qty_when_no_pallet_quantity_column() -> None:
    row = _standard_load_row()
    row.pop("TOTAL PALLETS")
    row["QTY"] = 10

    workbook = _workbook_with_sheet("LOAD SHEET", [row])

    rows = parse_standard_bol_excel(workbook)

    assert len(rows) == 1
    assert rows[0].unit_qty == "10"
    assert rows[0].plt_qty == "10"


def test_parse_standard_bol_excel_prefers_explicit_pallet_quantity_over_qty() -> None:
    row = _standard_load_row()
    row["PLT QTY"] = 3
    row["QTY"] = 10
    row.pop("TOTAL PALLETS")

    workbook = _workbook_with_sheet("LOAD SHEET", [row])

    rows = parse_standard_bol_excel(workbook)

    assert len(rows) == 1
    assert rows[0].unit_qty == "10"
    assert rows[0].plt_qty == "3"


def test_parse_standard_bol_excel_prefers_pallet_qty_alias_over_generic_qty() -> None:
    row = _standard_load_row()
    row.pop("TOTAL PALLETS")
    row["Pallet Qty"] = 4
    row["QTY"] = 10

    workbook = _workbook_with_sheet("LOAD SHEET", [row])

    rows = parse_standard_bol_excel(workbook)

    assert len(rows) == 1
    assert rows[0].unit_qty == "10"
    assert rows[0].plt_qty == "4"


def test_parse_standard_bol_excel_resolves_qty_fallback_with_wide_headers_quickly() -> None:
    row = _standard_load_row()
    row.pop("TOTAL PALLETS")
    row["QTY"] = 10
    wide_row = {f"Unused Column {index}": "" for index in range(500)}
    wide_row.update(row)
    workbook = _workbook_with_sheet("LOAD SHEET", [wide_row])

    started_at = perf_counter()
    rows = parse_standard_bol_excel(workbook)
    elapsed = perf_counter() - started_at

    assert len(rows) == 1
    assert rows[0].plt_qty == "10"
    assert elapsed < 2.0


def test_parse_standard_bol_excel_accepts_total_weight_alias_and_pickup() -> None:
    row = _standard_load_row()
    row["Line Weight"] = row.pop("Total Weight")
    row["Delivery Appointment Number"] = "APPT-123"

    workbook = _workbook_with_sheet("LOAD SHEET", [row])

    rows = parse_standard_bol_excel(workbook)

    assert len(rows) == 1
    assert rows[0].total_weight == "100"
    assert rows[0].pickup_number == "APPT-123"


def test_parse_standard_bol_excel_allows_missing_optional_total_weight_and_pickup() -> None:
    row = _standard_load_row()
    row.pop("Total Weight")

    workbook = _workbook_with_sheet("LOAD SHEET", [row])

    rows = parse_standard_bol_excel(workbook)

    assert len(rows) == 1
    assert rows[0].total_weight == ""
    assert rows[0].pickup_number == ""


def test_standard_bol_mapping_preserves_total_weight_and_first_pickup() -> None:
    row = _standard_load_row()
    row["Total Weight"] = "306 lbs."
    row["Pick Up #"] = "PU-001"

    workbook = _workbook_with_sheet("LOAD SHEET", [row])

    rows = parse_standard_bol_excel(workbook)
    records = map_standard_rows_to_records(rows)

    assert len(records) == 1
    assert records[0].pickup_number == "PU-001"
    assert records[0].item_lines[0].total_weight == "306 lbs."
