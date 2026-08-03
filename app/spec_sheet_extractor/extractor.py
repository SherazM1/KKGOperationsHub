"""PDF upload inventory helpers for the Spec Sheet Extractor module."""

from __future__ import annotations

from io import BytesIO
from typing import Protocol, Sequence

from pypdf import PdfReader

from app.spec_sheet_extractor.models import PdfFileInventoryResult, PdfPageInventoryRecord


class UploadedPdf(Protocol):
    """Minimal interface required from Streamlit uploaded files."""

    name: str

    def getvalue(self) -> bytes:
        """Return uploaded file bytes."""


def read_uploaded_pdf_bytes(uploaded_file: UploadedPdf) -> bytes:
    """Read uploaded PDF bytes in memory."""
    return uploaded_file.getvalue()


def inventory_pdf_uploads(
    uploaded_files: Sequence[UploadedPdf],
) -> tuple[list[PdfFileInventoryResult], list[PdfPageInventoryRecord]]:
    """Validate PDFs and create one inventory record per readable PDF page."""
    file_results: list[PdfFileInventoryResult] = []
    page_inventory: list[PdfPageInventoryRecord] = []

    for file_index, uploaded_file in enumerate(uploaded_files):
        source_filename = uploaded_file.name
        try:
            pdf_bytes = read_uploaded_pdf_bytes(uploaded_file)
            reader = PdfReader(BytesIO(pdf_bytes))
            page_count = len(reader.pages)
        except Exception as exc:
            file_results.append(
                PdfFileInventoryResult(
                    source_filename=source_filename,
                    source_file_index=file_index,
                    page_count=0,
                    status="Failed",
                    error_message=str(exc),
                )
            )
            continue

        file_results.append(
            PdfFileInventoryResult(
                source_filename=source_filename,
                source_file_index=file_index,
                page_count=page_count,
                status="Ready",
                error_message=None,
            )
        )
        page_inventory.extend(
            PdfPageInventoryRecord(
                source_filename=source_filename,
                source_file_index=file_index,
                page_number=page_number,
                status="Ready",
                error_message=None,
            )
            for page_number in range(1, page_count + 1)
        )

    return file_results, page_inventory
