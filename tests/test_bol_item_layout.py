from dataclasses import replace
from pathlib import Path

import pytest
from docx import Document
from docx.oxml.ns import qn
from pypdf import PdfReader

from app.services.bol_standard_docx_generator import generate_standard_docx_set, resolve_template_path_for_mode
from app.services.bol_multistop_docx_generator import generate_multistop_docx_set
from app.services.bol_pdf_template_stamper import stamp_bol_pdf_set
from app.utils.bol_facilities import BOL_FACILITY_LOOKUP
from tests.test_bol_standard_docx_generator import _ready_record
from tests.test_bol_multistop_grouping import _three_stop_record


def _assert_word_pages(file_path, markers):
    doc = Document(file_path)
    item_pages = []
    for table in doc.tables:
        rows = list(table.rows)
        header = next((i for i, row in enumerate(rows)
                       if "ITEM DESCRIPTION" in " ".join(c.text for c in row.cells).upper()), None)
        if header is None:
            continue
        totals = next(i for i in range(header + 1, len(rows))
                      if "TOTALS" in " ".join(c.text for c in rows[i].cells))
        assert totals - header - 1 == 15
        item_pages.append("\n".join(" ".join(dict.fromkeys(c.text for c in row.cells))
                                    for row in rows[header + 1:totals]))
    assert len(item_pages) == (len(markers) + 14) // 15
    assert len(doc.element.findall(".//w:pageBreakBefore", doc.element.nsmap)) == len(item_pages) - 1
    for index, marker in enumerate(markers):
        assert sum(text.count(marker) for text in item_pages) == 1
        assert marker in item_pages[index // 15]
    ids = [node.get("id") for node in doc.element.iter(qn("wp:docPr"))]
    assert len(ids) == len(set(ids))


@pytest.mark.parametrize("mode", ["Standard", "No Recourse"])
@pytest.mark.parametrize("count", [1, 15, 16, 31])
def test_standard_word_and_pdf_default_to_fifteen_rows(tmp_path, mode, count):
    record = _ready_record()
    markers = [f"Product{index:03d}" for index in range(count)]
    record.item_lines = [replace(record.item_lines[0], item_description=marker,
                                 item_number=f"ITEM{index:03d}")
                         for index, marker in enumerate(markers)]
    facility = BOL_FACILITY_LOOKUP["Green Bay Packaging"]
    docs = generate_standard_docx_set([record], facility,
        template_path=resolve_template_path_for_mode(mode), output_dir=tmp_path)
    assert not docs.failed_records
    _assert_word_pages(docs.generated_files[0].file_path, markers)
    pdfs = stamp_bol_pdf_set([record], facility, docs.generated_files,
                            mode=mode, output_dir=tmp_path / "pdf")
    assert not pdfs.failed_conversions
    pages = PdfReader(pdfs.converted_files[0].file_path).pages
    assert len(pages) == (count + 14) // 15
    texts = [page.extract_text() for page in pages]
    for index, marker in enumerate(markers):
        assert sum(text.count(marker) for text in texts) == 1
        assert marker in texts[index // 15]


@pytest.mark.parametrize("mode,template", [
    ("Standard", "standard_bol_template.docx"),
    ("No Recourse", "no_recourse_bol_template.docx"),
    ("Standard", "multistop_bol_template.docx"),
])
def test_multistop_master_and_children_continue_without_losing_items(tmp_path, mode, template):
    records = _three_stop_record()
    record = records[0]
    first_stop = record.stops[0]
    markers = [f"Product{index:03d}" for index in range(16)]
    first_stop.item_lines = [replace(record.item_lines[0], item_description=marker,
                                    item_number=f"ITEM{index:03d}")
                            for index, marker in enumerate(markers)]
    for number, stop in enumerate(record.stops[1:], start=2):
        stop.item_lines = [replace(record.item_lines[0], item_description=f"A marker for stop{number}")]
    record.item_lines = [line for stop in record.stops for line in stop.item_lines]
    facility = BOL_FACILITY_LOOKUP["Green Bay Packaging"]
    docs = generate_multistop_docx_set(records, facility,
        template_path=Path("app/templates") / template,
        individual_stop_template_path=resolve_template_path_for_mode(mode),
        master_template_mode=mode, output_dir=tmp_path,
        master_comment_enabled=True, master_comment_text="Dispatch note", comment_scope="All BOLs")
    assert not docs.failed_records
    child = next(f for f in docs.generated_files if f.document_type == "stop" and f.stop_number == 1)
    master = next(f for f in docs.generated_files if f.document_type == "combined")
    _assert_word_pages(child.file_path, markers)
    _assert_word_pages(master.file_path, markers + ["A marker for stop2", "A marker for stop3"])
    pdfs = stamp_bol_pdf_set(records, facility, docs.generated_files, mode="Multistop",
        multistop_no_recourse=mode == "No Recourse", output_dir=tmp_path / "pdf")
    assert not pdfs.failed_conversions
    for file in pdfs.converted_files:
        if file.document_type == "combined" or file.stop_number == 1:
            pages = PdfReader(file.file_path).pages
            assert len(pages) == 2
            text = "\n".join(page.extract_text() for page in pages)
            for marker in markers:
                assert text.count(marker) == 1
