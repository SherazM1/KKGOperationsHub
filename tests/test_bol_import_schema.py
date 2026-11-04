"""Behavioral checks for the operational load-sheet layouts and unsafe imports."""

import csv
from io import BytesIO, StringIO

import pytest
from openpyxl import Workbook

from app.services.bol_excel_headers import ColumnMappingRequired
from app.services.bol_standard_parser import parse_standard_bol_excel
from app.services.bol_multistop_parser import parse_multistop_bol_excel
from app.services.bol_standard_mapper import map_standard_rows_to_records
from app.services.bol_multistop_mapper import map_multistop_rows_to_records
from tests.test_bol_standard_parser import _standard_load_row


def upload(rows, csv_file=False, sheet_name="Load Sheet"):
    output = BytesIO()
    if csv_file:
        text = StringIO()
        csv.writer(text).writerows(rows)
        output.write(text.getvalue().encode())
        output.name = "loads.csv"
    else:
        book = Workbook()
        sheet = book.active
        sheet.title = sheet_name
        for row in rows:
            sheet.append(row)
        book.save(output)
        output.name = "loads.xlsx"
    output.seek(0)
    return output


def group_rows(*, mismatch=False, item_counts=True):
    base = _standard_load_row()
    base["TRACKERS"] = "TRACK-1"
    base["Units"] = base.pop("QTY")
    base["PLT Weight"] = base.pop("weight each")
    base.pop("TOTAL PALLETS")
    if item_counts:
        base["PLT QTY"] = 8
    base["Total PLT QTY"] = 26
    rows = []
    for index, (bol, load, qty, declared, weight) in enumerate([
        ("BOL-A", "LOAD-1", 8, 26, 7800),
        ("BOL-A", "LOAD-1", 17 if mismatch else 18, None, None),
        ("BOL-B", "LOAD-2", 26, 26, 7800),
    ]):
        row = dict(base, **{"BOL #": bol, "KK Load": load, "Units": qty,
                           "Total PLT QTY": declared, "PLT Weight": 300,
                           "Total Weight": weight, "ITEM #": f"ITEM-{index}"})
        if item_counts:
            row["PLT QTY"] = qty
        rows.append(list(row.values()))
    return [list(base), *rows]


@pytest.mark.parametrize("csv_file", [False, True])
def test_unknown_header_continuation_does_not_break_other_fields(csv_file):
    base = _standard_load_row()
    keys = list(base)
    headers = [{"Pallet Description": "Pallet", "DC ST": "DC", "DC ZIP": "DC"}.get(k, k) for k in keys]
    continuation = [{"Pallet Description": "Description", "DC ST": "ST", "DC ZIP": "ZIP"}.get(k) for k in keys]
    headers.append("Campaign")
    continuation.append("Details")
    rows = parse_standard_bol_excel(upload([headers, continuation, [*base.values(), "Wave 1"]], csv_file))
    assert len(rows) == 1
    assert rows[0].source_row_number == 3
    assert rows[0].item_description == "Test pallet"
    assert rows[0].dc_city_state_zip == "Dallas, TX 75001"
    assert rows[0].source_values["Campaign Details"] == "Wave 1"


@pytest.mark.parametrize("csv_file", [False, True])
def test_title_and_single_split_header_preserve_first_item(csv_file):
    base = _standard_load_row()
    keys = list(base)
    headers = ["Pallet" if key == "Pallet Description" else key for key in keys]
    second = ["Description" if key == "Pallet Description" else None for key in keys]
    rows = parse_standard_bol_excel(upload([
        ["Shipment report"], [None] * len(keys), headers, second, list(base.values()),
    ], csv_file, sheet_name="Operational Data"))
    assert len(rows) == 1
    assert rows[0].source_row_number == 5
    assert rows[0].bol_number == "BOL-001"


@pytest.mark.parametrize("csv_file", [False, True])
@pytest.mark.parametrize("multistop", [False, True])
def test_item_pallets_and_declared_group_total_are_not_double_counted(csv_file, multistop):
    parser = parse_multistop_bol_excel if multistop else parse_standard_bol_excel
    mapper = map_multistop_rows_to_records if multistop else map_standard_rows_to_records
    rows = parser(upload(group_rows(), csv_file))
    records = mapper(rows)
    assert len(rows) == 3
    assert len(records) == 2
    assert all(record.is_ready for record in records)
    assert [line.skids for line in records[0].item_lines] == ["8", "18"]
    assert [line.total_weight for line in records[0].item_lines] == ["2400", "5400"]
    assert records[0].total_pallet == 26
    assert records[0].total_ship_weight == 7800
    assert rows[0].source_values["Total Weight"] == "7800"
    assert rows[1].source_values["Total Weight"] == ""
    assert rows[0].source_values["Total PLT QTY"] == "26"


@pytest.mark.parametrize("multistop", [False, True])
def test_group_mismatch_requires_review(multistop):
    parser = parse_multistop_bol_excel if multistop else parse_standard_bol_excel
    mapper = map_multistop_rows_to_records if multistop else map_standard_rows_to_records
    records = mapper(parser(upload(group_rows(mismatch=True))))
    assert not records[0].is_ready
    assert any("sum to 25" in issue for issue in records[0].issues)
    assert records[1].is_ready


def test_load_pallet_group_can_span_multiple_destination_bols():
    data = group_rows()
    headers = data[0]
    data[2][headers.index("BOL #")] = "BOL-SECOND-STOP"
    data[2][headers.index("DC #")] = "5678"
    records = map_multistop_rows_to_records(parse_multistop_bol_excel(upload(data)))
    assert records[0].is_ready
    assert records[0].stop_count == 2
    assert records[0].total_pallet == 26
    assert records[0].total_ship_weight == 7800


def test_explicit_load_weight_is_not_summed_as_a_line_weight():
    row = _standard_load_row()
    row["Total Load Weight"] = 999
    rows = parse_standard_bol_excel(upload([list(row), list(row.values())]))
    assert rows[0].total_weight == "100"
    assert rows[0].group_weight_total == "999"
    assert not map_standard_rows_to_records(rows)[0].is_ready


def test_multi_item_group_total_is_not_distributed_without_item_counts():
    rows = parse_standard_bol_excel(upload(group_rows(item_counts=False)))
    assert [row.plt_qty for row in rows[:2]] == ["", ""]
    assert not map_standard_rows_to_records(rows)[0].is_ready
    assert rows[2].plt_qty == "26"  # A single-item group is unambiguous.


@pytest.mark.parametrize("csv_file", [False, True])
def test_duplicate_quantity_columns_require_and_honor_explicit_choices(csv_file):
    row = _standard_load_row()
    source_rows = [[*row.keys(), "QTY"], [*row.values(), 99]]
    with pytest.raises(ColumnMappingRequired) as caught:
        parse_standard_bol_excel(upload(source_rows, csv_file))
    options = caught.value.requests["unit_qty"]
    assert len(options) == 2
    assert all("[column " in value for value in options)
    assert caught.value.samples[options[0]] == ["10"]
    parsed = parse_standard_bol_excel(upload(source_rows, csv_file),
                                     column_overrides={"unit_qty": options[0]})
    assert parsed[0].unit_qty == "10"
    assert [parsed[0].source_values[c] for c in options] == ["10", "99"]


def test_duplicate_qty_fallback_has_its_own_pallet_mapping_request():
    row = _standard_load_row()
    row.pop("TOTAL PALLETS")
    data = [[*row.keys(), "QTY"], [*row.values(), 99]]
    with pytest.raises(ColumnMappingRequired) as caught:
        parse_standard_bol_excel(upload(data))
    choice = caught.value.requests["unit_qty"][0]
    with pytest.raises(ColumnMappingRequired) as caught:
        parse_standard_bol_excel(upload(data), column_overrides={"unit_qty": choice})
    assert "plt_qty" in caught.value.requests
    parsed = parse_standard_bol_excel(upload(data), column_overrides={"unit_qty": choice, "plt_qty": choice})
    assert parsed[0].unit_qty == parsed[0].plt_qty == "10"


@pytest.mark.parametrize("csv_file", [False, True])
def test_retailer_po_does_not_use_pod_or_po_date(csv_file):
    row = _standard_load_row()
    row["RETAILER PO #"] = row.pop("TGT PO #")
    data = [["POD", "Purchase Order Date", *row.keys()], [True, "2026-11-04", *row.values()]]
    parsed = parse_standard_bol_excel(upload(data, csv_file))
    assert parsed[0].wm_po == "TGT-001"
    assert parsed[0].column_mapping["wm_po"] == "RETAILER PO #"


def test_explicit_po_has_priority_over_legacy_retailer_metadata_in_both_modes():
    row = _standard_load_row()
    row["WMT PO #"] = row.pop("TGT PO #")
    row["RETAILER"] = "Store name, not a PO"
    row["TRACKERS"] = "TRACK-1"
    data = [list(row), list(row.values())]
    assert parse_standard_bol_excel(upload(data))[0].wm_po == "TGT-001"
    assert parse_multistop_bol_excel(upload(data))[0].target_po_number == "TGT-001"


@pytest.mark.parametrize("header", ["POD", "Purchase Order Date", "Cust_Order_PO_Value"])
def test_unapproved_po_label_is_not_silently_guessed(header):
    row = _standard_load_row()
    row[header] = row.pop("TGT PO #")
    data = [list(row), list(row.values())]
    assert parse_standard_bol_excel(upload(data))[0].wm_po == ""
    assert parse_standard_bol_excel(upload(data), column_overrides={"wm_po": header})[0].wm_po == "TGT-001"


def test_weight_footer_is_not_an_extra_bol_and_long_blank_gap_does_not_drop_items():
    row = _standard_load_row()
    footer = [None] * len(row)
    footer[list(row).index("Total Weight")] = 10000
    rows = parse_standard_bol_excel(upload([
        list(row), list(row.values()), *([[None] * len(row)] * 100), list(row.values()), footer,
    ]))
    assert len(rows) == 2
    assert [r.source_row_number for r in rows] == [2, 103]
    assert any("summary rows" in note for note in rows[0].parsing_notes)


@pytest.mark.parametrize("value", ["not a weight", "nan", "inf", "-10"])
def test_invalid_numeric_weight_cannot_be_marked_ready(value):
    row = _standard_load_row()
    row["weight each"] = value
    records = map_standard_rows_to_records(parse_standard_bol_excel(upload([list(row), list(row.values())])))
    assert not records[0].is_ready
    assert records[0].import_issues


def test_full_width_headers_preserve_identifier_values():
    row = _standard_load_row()
    def wide(text):
        return "".join(chr(ord(c) + 0xFEE0) if "!" <= c <= "~" else c for c in text)
    parsed = parse_standard_bol_excel(upload([[wide(key) for key in row], list(row.values())]))
    assert parsed[0].upc == "000111222333"
    assert parsed[0].bol_number == "BOL-001"


@pytest.mark.parametrize("mode", ["Standard", "No Recourse"])
def test_reconciled_wave_group_reaches_generated_pdf_totals(tmp_path, mode):
    from pypdf import PdfReader
    from app.services.bol_standard_docx_generator import generate_standard_docx_set, resolve_template_path_for_mode
    from app.services.bol_pdf_template_stamper import stamp_bol_pdf_set
    from app.utils.bol_facilities import BOL_FACILITY_LOOKUP
    rows = parse_standard_bol_excel(upload(group_rows()))
    records = map_standard_rows_to_records(rows)
    facility = BOL_FACILITY_LOOKUP["Green Bay Packaging"]
    docs = generate_standard_docx_set(records, facility, template_path=resolve_template_path_for_mode(mode),
                                      output_dir=tmp_path / "docx")
    assert not docs.failed_records
    pdfs = stamp_bol_pdf_set(records, facility, docs.generated_files, mode=mode,
                             output_dir=tmp_path / "pdf")
    assert not pdfs.failed_conversions
    first = next(file for file in pdfs.converted_files if file.bol_number == "BOL-A")
    text = " ".join(PdfReader(first.file_path).pages[0].extract_text().split())
    assert "7800" in text
    assert "26" in text
    assert text.count("ITEM-0") == text.count("ITEM-1") == 1


@pytest.mark.parametrize("multistop", [False, True])
def test_group_validation_cannot_be_bypassed_by_incomplete_generation(tmp_path, multistop):
    from app.services.bol_standard_docx_generator import generate_standard_docx_set, resolve_template_path_for_mode
    from app.services.bol_multistop_docx_generator import generate_multistop_docx_set
    from app.utils.bol_facilities import BOL_FACILITY_LOOKUP
    parser = parse_multistop_bol_excel if multistop else parse_standard_bol_excel
    mapper = map_multistop_rows_to_records if multistop else map_standard_rows_to_records
    records = mapper(parser(upload(group_rows(mismatch=True))))
    assert records[0].import_issues
    generator = generate_multistop_docx_set if multistop else generate_standard_docx_set
    result = generator(records, BOL_FACILITY_LOOKUP["Green Bay Packaging"],
                       template_path=resolve_template_path_for_mode("Standard"),
                       output_dir=tmp_path, allow_incomplete_records=True)
    assert not result.failed_records
    assert result.skipped_count == 1
    assert "sum to 25" in result.skipped_records[0].reason
    assert all(file.bol_number == "BOL-B" for file in result.generated_files)


def test_reused_bol_cannot_silently_combine_different_loads():
    first = _standard_load_row()
    second = dict(first, **{"KK Load": "ANOTHER-LOAD"})
    records = map_standard_rows_to_records(parse_standard_bol_excel(
        upload([list(first), list(first.values()), list(second.values())])))
    assert len(records) == 1
    assert not records[0].is_ready
    assert any("multiple KK loads" in issue for issue in records[0].import_issues)
