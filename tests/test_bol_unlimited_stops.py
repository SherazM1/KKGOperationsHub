"""Stop counts and optional identifiers survive import and complete BOL output."""

from pathlib import Path

import pytest
from docx import Document
from pypdf import PdfReader

from app.services.bol_multistop_docx_generator import generate_multistop_docx_set
from app.services.bol_multistop_mapper import map_multistop_rows_to_records
from app.services.bol_multistop_parser import parse_multistop_bol_excel
from app.services.bol_pdf_template_stamper import stamp_bol_pdf_set
from app.services.bol_standard_docx_generator import generate_standard_docx_set, resolve_template_path_for_mode
from app.services.bol_standard_mapper import map_standard_rows_to_records
from app.services.bol_standard_parser import parse_standard_bol_excel
from app.utils.bol_facilities import BOL_FACILITY_LOOKUP
from tests.test_bol_import_schema import upload
from tests.test_bol_multistop_grouping import _row
from tests.test_bol_multistop_parser import _multistop_load_row
from tests.test_bol_standard_parser import _standard_load_row
from tests.test_bol_standard_docx_generator import _ready_record


@pytest.mark.parametrize("multistop", [False, True])
@pytest.mark.parametrize("csv_file", [False, True])
@pytest.mark.parametrize("omit_columns", [False, True])
def test_item_number_and_upc_are_optional(multistop, csv_file, omit_columns):
    row = _multistop_load_row() if multistop else _standard_load_row()
    for field in ("ITEM #", "UPC"):
        if omit_columns:
            row.pop(field)
        else:
            row[field] = ""
    parser = parse_multistop_bol_excel if multistop else parse_standard_bol_excel
    mapper = map_multistop_rows_to_records if multistop else map_standard_rows_to_records
    rows = parser(upload([list(row), list(row.values())], csv_file=csv_file))
    assert rows[0].item_number == rows[0].upc == ""
    record, = mapper(rows)
    assert record.is_ready, record.issues
    assert not record.import_issues


@pytest.mark.parametrize("mode", ["Standard", "No Recourse"])
@pytest.mark.parametrize("missing", ["item_number", "upc"])
def test_word_keeps_the_identifier_that_was_supplied(tmp_path, mode, missing):
    record = _ready_record()
    line = record.item_lines[0]
    line.item_number, line.upc = "PRESENT-ITEM-123", "000123456789"
    setattr(line, missing, "")
    result = generate_standard_docx_set(
        [record], BOL_FACILITY_LOOKUP["Green Bay Packaging"],
        template_path=resolve_template_path_for_mode(mode), output_dir=tmp_path,
    )
    assert result.generated_count == 1, result.failed_records
    doc = Document(result.generated_files[0].file_path)
    text = "\n".join(node.text or "" for node in doc.element.body.iter()
                     if node.tag.endswith("}t"))
    assert line.item_description in text
    if missing == "item_number":
        assert "000123456789" in text
        assert "Item #:" not in text
    else:
        assert "PRESENT-ITEM-123" in text
        assert "UPC" not in text


@pytest.mark.parametrize("count", [4, 16, 100])
def test_all_stop_numbers_are_mapped_and_totalled(count):
    rows = [_row(kk_load="1", stop=stop, bol_number=f"BOL-{stop:03d}")
            for stop in range(count, 0, -1)]
    record, = map_multistop_rows_to_records(rows)
    assert record.is_ready, record.issues
    assert record.is_supported
    assert record.stop_count == count
    assert [stop.stop_number for stop in record.stops] == list(range(1, count + 1))
    assert len(record.item_lines) == count
    assert record.total_case == count * 10
    assert record.total_pallet == count * 2
    assert record.total_ship_weight == count * 100


def test_inferred_stops_can_exceed_three():
    rows = []
    for number in range(1, 7):
        row = _multistop_load_row()
        row.pop("Stop")
        row.update({"DC #": str(number), "BOL #": f"BOL-{number}"})
        rows.append(row)
    data = [list(rows[0]), *[list(row.values()) for row in rows]]
    record, = map_multistop_rows_to_records(parse_multistop_bol_excel(upload(data)))
    assert record.is_ready, record.issues
    assert [stop.stop_number for stop in record.stops] == list(range(1, 7))


def test_missing_stop_still_requires_review_without_iterating_the_gap():
    rows = [_row(kk_load="1", stop=1, bol_number="A"),
            _row(kk_load="1", stop=1_000_000, bol_number="B")]
    record, = map_multistop_rows_to_records(rows)
    assert not record.is_ready
    assert any("Missing Stops 2-999999" in issue for issue in record.issues)
    assert len(record.stops) == 2


@pytest.mark.parametrize("count,mode,template", [
    (4, "Standard", "multistop_bol_template.docx"),
    (16, "Standard", "standard_bol_template.docx"),
    (16, "No Recourse", "no_recourse_bol_template.docx"),
])
def test_every_delivery_and_child_bol_survives_word_and_pdf(tmp_path, count, mode, template):
    rows = [_row(kk_load="1", stop=number, bol_number=f"BOL-{number:03d}")
            for number in range(1, count + 1)]
    for row in rows:
        row.item_number = row.upc = ""
        row.dc_address = f"{row.stop_number} Delivery Marker Avenue"
    records = map_multistop_rows_to_records(rows)
    facility = BOL_FACILITY_LOOKUP["Green Bay Packaging"]
    docs = generate_multistop_docx_set(
        records, facility, output_dir=tmp_path / "word",
        template_path=Path("app/templates") / template,
        individual_stop_template_path=resolve_template_path_for_mode(mode),
        master_template_mode=mode,
    )
    assert not docs.failed_records
    assert not docs.skipped_records
    assert len(docs.generated_files) == count + 1
    assert {file.stop_number for file in docs.generated_files if file.document_type == "stop"} == set(range(1, count + 1))
    master = next(file for file in docs.generated_files if file.document_type == "combined")
    doc = Document(master.file_path)
    text = "\n".join(node.text or "" for node in doc.element.body.iter()
                     if node.tag.endswith("}t"))
    assert "Delivery continuation" in text
    for number in range(1, count + 1):
        assert f"{number} Delivery Marker Avenue" in text
        assert f"Pallet {number}" in text
    assert "Item #:" not in text and "UPC #:" not in text

    pdfs = stamp_bol_pdf_set(
        records, facility, docs.generated_files, mode="Multistop",
        multistop_no_recourse=mode == "No Recourse", output_dir=tmp_path / "pdf",
    )
    assert not pdfs.failed_conversions
    assert len(pdfs.converted_files) == count + 1
    master_pdf = next(file for file in pdfs.converted_files if file.document_type == "combined")
    pages = PdfReader(master_pdf.file_path).pages
    assert len(pages) == (count + 14) // 15 + (count + 7) // 8
    text = "\n".join(page.extract_text() for page in pages)
    for number in range(1, count + 1):
        assert f"{number} Delivery Marker Avenue" in text
        assert text.count(f"Pallet {number}\n") == 1
    for page in pages[(count + 14) // 15:]:
        assert "Delivery continuation" in page.extract_text()
        assert "Master BOL: BOL-001" in page.extract_text()
    schedule = "\n".join(page.extract_text() for page in pages[(count + 14) // 15:])
    for number in range(1, count + 1):
        assert f"{number} Delivery Marker Avenue" in schedule
    assert "Item #:" not in text and "UPC #:" not in text
