"""Delivery schedules use the available page instead of a fixed stop count."""

from io import BytesIO

import pytest
from docx import Document
from pypdf import PdfReader, PdfWriter

from app.services.bol_multistop_mapper import map_multistop_rows_to_records
from app.services.bol_stop_continuation import (
    append_delivery_continuation_docx,
    append_delivery_continuation_pdf,
)
from tests.test_bol_multistop_grouping import _row


def _record(count, long_addresses=False):
    rows = [_row(kk_load="1", stop=number, bol_number=f"BOL-{number:03d}")
            for number in range(1, count + 1)]
    for row in rows:
        row.dc_address = f"{row.stop_number} Delivery Marker Avenue"
        if long_addresses:
            row.dc_address += ", Receiving Building North, " * 4
    return map_multistop_rows_to_records(rows)[0]


@pytest.mark.parametrize("count", [4, 9, 16, 20])
def test_delivery_list_fits_one_pdf_page_and_keeps_all_stops(tmp_path, count):
    path = tmp_path / "master.pdf"
    writer = PdfWriter()
    writer.add_blank_page(width=612, height=792)
    writer.write(path)
    append_delivery_continuation_pdf(path, _record(count))
    pages = PdfReader(path).pages
    assert len(pages) == 2  # Original master plus one delivery schedule.
    text = pages[1].extract_text()
    assert "Delivery continuation" in text
    for number in range(1, count + 1):
        assert f"{number} Delivery Marker Avenue" in text
        assert f"BOL-{number:03d}" in text


def test_long_large_schedule_flows_without_dropping_stops_or_tiny_type(tmp_path):
    path = tmp_path / "master.pdf"
    writer = PdfWriter()
    writer.add_blank_page(width=612, height=792)
    writer.write(path)
    append_delivery_continuation_pdf(path, _record(60, long_addresses=True))
    pages = PdfReader(path).pages[1:]
    assert len(pages) > 1
    text = "\n".join(page.extract_text() for page in pages)
    for number in range(1, 61):
        assert f"BOL-{number:03d}" in text
    sizes = []
    for page in pages:
        assert "Delivery continuation" in page.extract_text()
        assert "Master BOL: BOL-001" in page.extract_text()
        page.extract_text(visitor_text=lambda text, cm, tm, font, size: sizes.append(size) if text.strip() else None)
    assert min(sizes) >= 7


def test_word_uses_one_compact_flowing_table_for_sixteen_stops():
    doc = Document("app/templates/standard_bol_template.docx")
    before = len(doc.tables)
    append_delivery_continuation_docx(doc, _record(16))
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    doc = Document(buffer)
    assert len(doc.tables) == before + 1
    table = doc.tables[-1]
    assert len(table.rows) == 17
    text = "\n".join(cell.text for row in table.rows for cell in row.cells)
    for number in range(1, 17):
        assert f"{number} Delivery Marker Avenue" in text
    sizes = [run.font.size.pt for row in table.rows for cell in row.cells
             for paragraph in cell.paragraphs for run in paragraph.runs if run.text.strip()]
    assert min(sizes) >= 7


def test_three_stops_do_not_add_a_delivery_schedule(tmp_path):
    record = _record(3)
    doc = Document()
    append_delivery_continuation_docx(doc, record)
    assert not doc.tables
    path = tmp_path / "master.pdf"
    writer = PdfWriter()
    writer.add_blank_page(width=612, height=792)
    writer.write(path)
    original = path.read_bytes()
    append_delivery_continuation_pdf(path, record)
    assert path.read_bytes() == original
