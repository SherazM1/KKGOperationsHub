"""Tests for Spec Sheet Extractor PDF inventory helpers."""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO

import pytest
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

from app.spec_sheet_extractor import extractor
from app.spec_sheet_extractor.extractor import extract_header_fields_from_uploads, inventory_pdf_uploads
from app.spec_sheet_extractor.models import SPEC_SHEET_FIXED_FIELDS
from app.spec_sheet_extractor.zones import HEADER_FIELD_ZONES


@dataclass(frozen=True)
class UploadedFileStub:
    name: str
    data: bytes

    def getvalue(self) -> bytes:
        return self.data


def _pdf_bytes(page_count: int) -> bytes:
    buffer = BytesIO()
    canv = canvas.Canvas(buffer, pagesize=letter)
    for page_number in range(1, page_count + 1):
        canv.drawString(72, 720, f"Page {page_number}")
        canv.showPage()
    canv.save()
    return buffer.getvalue()


def _representative_values() -> dict[str, str]:
    return {
        "Customer": "Fresh Step",
        "Design": "KK260127-02",
        "Revision": "0",
        "Part": "Graphic Box",
        "Opportunity/Project #": "3356",
        "Pieces per set": "2",
        "Board": "200 B",
        "Corr direction": "Vertical",
        "View": "outside",
        "Production/Project Manager": "Marta Espina",
        "Designer": "Andy Bales",
        "ID": "FS-01",
        "Area": "1766.81",
        "Blank width": "53+9/16",
        "Blank height": "38+3/4",
        "Inches of rule": "522+25/32",
        "Date": "07/20/2026",
    }


def _label_for_field(field_name: str, manager_label: str) -> str:
    if field_name == "Production/Project Manager":
        return manager_label
    return field_name


def _header_pdf_bytes(
    pages: list[dict[str, str]],
    *,
    manager_label: str = "Production Mngr",
    rotate_degrees: int | None = None,
) -> bytes:
    buffer = BytesIO()
    canv = canvas.Canvas(buffer, pagesize=letter)
    page_width, page_height = letter
    zone_by_field = {zone.field_name: zone for zone in HEADER_FIELD_ZONES}

    for page_values in pages:
        canv.setFont("Helvetica", 9)
        for field_name in SPEC_SHEET_FIXED_FIELDS:
            zone = zone_by_field[field_name]
            x = zone.left * page_width + 4
            y = ((zone.bottom + zone.top) / 2) * page_height
            label = _label_for_field(field_name, manager_label)
            value = page_values.get(field_name, "")
            canv.drawString(x, y, f"{label}: {value}")
        canv.setFont("Helvetica", 11)
        canv.drawString(72, 420, "BODY DIELINE MEASUREMENT Blank width: 999+1/2")
        canv.drawString(72, 400, "Inches of rule: 999 Date: 01/01/1900")
        canv.showPage()

    canv.save()
    pdf_bytes = buffer.getvalue()
    if rotate_degrees is None:
        return pdf_bytes

    reader = PdfReader(BytesIO(pdf_bytes))
    writer = PdfWriter()
    for page in reader.pages:
        page.rotate(rotate_degrees)
        writer.add_page(page)
    rotated = BytesIO()
    writer.write(rotated)
    return rotated.getvalue()


def test_inventories_one_valid_single_page_pdf() -> None:
    file_results, page_inventory = inventory_pdf_uploads(
        [UploadedFileStub("single.pdf", _pdf_bytes(1))]
    )

    assert len(file_results) == 1
    assert file_results[0].source_filename == "single.pdf"
    assert file_results[0].page_count == 1
    assert file_results[0].status == "Ready"
    assert len(page_inventory) == 1
    assert page_inventory[0].source_filename == "single.pdf"
    assert page_inventory[0].page_number == 1


def test_inventories_one_valid_multi_page_pdf() -> None:
    file_results, page_inventory = inventory_pdf_uploads(
        [UploadedFileStub("multi.pdf", _pdf_bytes(3))]
    )

    assert file_results[0].page_count == 3
    assert [record.page_number for record in page_inventory] == [1, 2, 3]


def test_multiple_pdfs_preserve_file_and_page_order() -> None:
    file_results, page_inventory = inventory_pdf_uploads(
        [
            UploadedFileStub("a.pdf", _pdf_bytes(2)),
            UploadedFileStub("b.pdf", _pdf_bytes(3)),
        ]
    )

    assert [result.source_filename for result in file_results] == ["a.pdf", "b.pdf"]
    assert [result.source_file_index for result in file_results] == [0, 1]
    assert [
        (record.source_filename, record.source_file_index, record.page_number)
        for record in page_inventory
    ] == [
        ("a.pdf", 0, 1),
        ("a.pdf", 0, 2),
        ("b.pdf", 1, 1),
        ("b.pdf", 1, 2),
        ("b.pdf", 1, 3),
    ]


def test_invalid_pdf_mixed_with_valid_pdf_records_failure_and_valid_pages() -> None:
    file_results, page_inventory = inventory_pdf_uploads(
        [
            UploadedFileStub("bad.pdf", b"not a pdf"),
            UploadedFileStub("good.pdf", _pdf_bytes(2)),
        ]
    )

    assert file_results[0].source_filename == "bad.pdf"
    assert file_results[0].status == "Failed"
    assert file_results[0].page_count == 0
    assert file_results[0].error_message
    assert file_results[1].source_filename == "good.pdf"
    assert file_results[1].status == "Ready"
    assert [record.source_filename for record in page_inventory] == ["good.pdf", "good.pdf"]


def test_page_numbering_starts_at_one() -> None:
    _, page_inventory = inventory_pdf_uploads([UploadedFileStub("pages.pdf", _pdf_bytes(4))])

    assert page_inventory[0].page_number == 1
    assert [record.page_number for record in page_inventory] == [1, 2, 3, 4]


def test_failed_files_do_not_prevent_valid_files_from_being_inventoried() -> None:
    file_results, page_inventory = inventory_pdf_uploads(
        [
            UploadedFileStub("first.pdf", _pdf_bytes(1)),
            UploadedFileStub("bad.pdf", b"%PDF broken"),
            UploadedFileStub("last.pdf", _pdf_bytes(1)),
        ]
    )

    assert [result.status for result in file_results] == ["Ready", "Failed", "Ready"]
    assert [
        (record.source_filename, record.page_number)
        for record in page_inventory
    ] == [("first.pdf", 1), ("last.pdf", 1)]


def test_extracts_all_17_fields_from_representative_page() -> None:
    values = _representative_values()

    results = extract_header_fields_from_uploads(
        [UploadedFileStub("fresh-step.pdf", _header_pdf_bytes([values]))]
    )

    assert len(results) == 1
    row = results[0].to_preview_row()
    assert results[0].extraction_status == "Extracted"
    for field_name, expected_value in values.items():
        assert row[field_name] == expected_value


def test_multi_page_extraction_preserves_order() -> None:
    first = {**_representative_values(), "Design": "KK260002-01"}
    second = {**_representative_values(), "Design": "KK260002-02", "Customer": "Energizer"}

    results = extract_header_fields_from_uploads(
        [UploadedFileStub("multi-header.pdf", _header_pdf_bytes([first, second]))]
    )

    assert [(result.source_filename, result.page_number) for result in results] == [
        ("multi-header.pdf", 1),
        ("multi-header.pdf", 2),
    ]
    assert [result.design for result in results] == ["KK260002-01", "KK260002-02"]
    assert [result.customer for result in results] == ["Fresh Step", "Energizer"]


def test_rotation_normalization_uses_same_header_zones() -> None:
    values = {**_representative_values(), "Customer": "Rotated Customer"}

    results = extract_header_fields_from_uploads(
        [UploadedFileStub("rotated.pdf", _header_pdf_bytes([values], rotate_degrees=90))]
    )

    assert results[0].extraction_status == "Extracted"
    assert results[0].customer == "Rotated Customer"
    assert results[0].blank_width == "53+9/16"


def test_blank_id_field_remains_blank() -> None:
    values = {**_representative_values(), "ID": ""}

    results = extract_header_fields_from_uploads(
        [UploadedFileStub("blank-id.pdf", _header_pdf_bytes([values]))]
    )

    assert results[0].id == ""
    assert results[0].extraction_status == "Extracted with blanks"


def test_fractional_values_are_preserved_exactly() -> None:
    values = {
        **_representative_values(),
        "Blank width": "53+9/16",
        "Blank height": "38+3/4",
        "Inches of rule": "522+25/32",
    }

    results = extract_header_fields_from_uploads(
        [UploadedFileStub("fractions.pdf", _header_pdf_bytes([values]))]
    )

    assert results[0].blank_width == "53+9/16"
    assert results[0].blank_height == "38+3/4"
    assert results[0].inches_of_rule == "522+25/32"


def test_opportunity_project_prefix_is_preserved() -> None:
    values = {**_representative_values(), "Opportunity/Project #": "Oppty-3075"}

    results = extract_header_fields_from_uploads(
        [UploadedFileStub("oppty.pdf", _header_pdf_bytes([values]))]
    )

    assert results[0].opportunity_project_number == "Oppty-3075"


def test_pieces_per_set_preserves_wording() -> None:
    values = {**_representative_values(), "Pieces per set": "1 Set Required"}

    results = extract_header_fields_from_uploads(
        [UploadedFileStub("pieces.pdf", _header_pdf_bytes([values]))]
    )

    assert results[0].pieces_per_set == "1 Set Required"


@pytest.mark.parametrize("manager_label", ["Production Mngr", "Project Mngr"])
def test_manager_label_variants_map_to_single_output_field(manager_label: str) -> None:
    values = {**_representative_values(), "Production/Project Manager": "Marta Espina"}

    results = extract_header_fields_from_uploads(
        [UploadedFileStub("manager.pdf", _header_pdf_bytes([values], manager_label=manager_label))]
    )

    assert results[0].production_project_manager == "Marta Espina"


def test_one_failed_page_does_not_block_remaining_pages(monkeypatch: pytest.MonkeyPatch) -> None:
    values = _representative_values()
    original_extract_page = extractor._extract_page_header_fields
    call_count = 0

    def fail_second_page(page: object) -> dict[str, str]:
        nonlocal call_count
        call_count += 1
        if call_count == 2:
            raise ValueError("page could not be processed")
        return original_extract_page(page)

    monkeypatch.setattr(extractor, "_extract_page_header_fields", fail_second_page)

    results = extractor.extract_header_fields_from_uploads(
        [UploadedFileStub("one-bad-page.pdf", _header_pdf_bytes([values, values, values]))]
    )

    assert [result.extraction_status for result in results] == [
        "Extracted",
        "Failed",
        "Extracted",
    ]
    assert results[1].error_message == "page could not be processed"
    assert results[2].customer == "Fresh Step"


def test_body_dieline_measurement_text_does_not_leak_into_header_fields() -> None:
    values = {**_representative_values(), "Blank width": "", "Inches of rule": ""}

    results = extract_header_fields_from_uploads(
        [UploadedFileStub("body-text.pdf", _header_pdf_bytes([values]))]
    )

    assert results[0].blank_width == ""
    assert results[0].inches_of_rule == ""
