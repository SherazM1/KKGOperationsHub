"""Tests for Spec Sheet Extractor PDF inventory helpers."""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

from app.spec_sheet_extractor.extractor import inventory_pdf_uploads


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
