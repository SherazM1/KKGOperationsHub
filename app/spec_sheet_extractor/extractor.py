"""PDF upload inventory helpers for the Spec Sheet Extractor module."""

from __future__ import annotations

from io import BytesIO
import re
from typing import Protocol, Sequence

from pypdf import PdfReader

from app.spec_sheet_extractor.models import (
    FIELD_ATTRIBUTE_BY_LABEL,
    SPEC_SHEET_FIXED_FIELDS,
    PdfFileInventoryResult,
    PdfHeaderExtractionResult,
    PdfPageInventoryRecord,
)
from app.spec_sheet_extractor.zones import HEADER_FIELD_ZONES, NormalizedTextZone


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


_FIELD_LABEL_PATTERNS: dict[str, tuple[str, ...]] = {
    "Customer": (r"Customer",),
    "Design": (r"Design",),
    "Revision": (r"Revision", r"Rev"),
    "Part": (r"Part",),
    "Opportunity/Project #": (
        r"Opportunity\s*/\s*Project\s*#",
        r"Opportunity\s*#",
        r"Project\s*#",
        r"Oppty\s*#",
    ),
    "Pieces per set": (r"Pieces\s+per\s+set", r"Pieces\s*/\s*set"),
    "Board": (r"Board",),
    "Corr direction": (r"Corr\s+direction", r"Corr\s+dir(?:ection)?"),
    "View": (r"View",),
    "Production/Project Manager": (
        r"Production\s*/\s*Project\s+Manager",
        r"Production\s+Mngr",
        r"Project\s+Mngr",
        r"Production\s+Manager",
        r"Project\s+Manager",
    ),
    "Designer": (r"Designer",),
    "ID": (r"ID",),
    "Area": (r"Area",),
    "Blank width": (r"Blank\s+width",),
    "Blank height": (r"Blank\s+height",),
    "Inches of rule": (r"Inches\s+of\s+rule",),
    "Date": (r"Date",),
}


def _clean_pdf_text(text: str) -> str:
    """Collapse obvious PDF whitespace artifacts without changing value text."""
    cleaned = text.replace("\xa0", " ")
    cleaned = re.sub(r"[ \t\r\f\v]+", " ", cleaned)
    cleaned = re.sub(r"\s*\n\s*", " ", cleaned)
    return cleaned.strip()


def _strip_printed_field_label(field_name: str, text: str) -> str:
    """Remove the template label from a zone while preserving the value."""
    value = _clean_pdf_text(text)
    for label_pattern in _FIELD_LABEL_PATTERNS[field_name]:
        value = re.sub(
            rf"^\s*{label_pattern}\s*:?\s*",
            "",
            value,
            count=1,
            flags=re.IGNORECASE,
        )
    return value.strip()


def _collect_zone_text(page: object, zone: NormalizedTextZone) -> str:
    page_width = float(page.mediabox.width)
    page_height = float(page.mediabox.height)
    text_chunks: list[tuple[float, float, str]] = []

    def visitor_text(text: str, _cm: object, tm: Sequence[float], _font: object, _font_size: float) -> None:
        if not text.strip():
            return
        x = float(tm[4])
        y = float(tm[5])
        if zone.contains(x, y, page_width, page_height):
            text_chunks.append((y, x, text))

    page.extract_text(visitor_text=visitor_text)
    text_chunks.sort(key=lambda chunk: (-chunk[0], chunk[1]))
    return _clean_pdf_text(" ".join(chunk[2] for chunk in text_chunks))


def _extract_page_header_fields(page: object) -> dict[str, str]:
    extracted_fields: dict[str, str] = {}
    for zone in HEADER_FIELD_ZONES:
        zone_text = _collect_zone_text(page, zone)
        extracted_fields[zone.field_name] = _strip_printed_field_label(zone.field_name, zone_text)
    return extracted_fields


def _header_result_from_fields(
    *,
    source_filename: str,
    source_file_index: int,
    page_number: int,
    extraction_status: str,
    error_message: str | None,
    extracted_fields: dict[str, str] | None = None,
) -> PdfHeaderExtractionResult:
    values = {FIELD_ATTRIBUTE_BY_LABEL[label]: "" for label in SPEC_SHEET_FIXED_FIELDS}
    for label, value in (extracted_fields or {}).items():
        values[FIELD_ATTRIBUTE_BY_LABEL[label]] = value
    return PdfHeaderExtractionResult(
        source_filename=source_filename,
        source_file_index=source_file_index,
        page_number=page_number,
        extraction_status=extraction_status,
        error_message=error_message,
        **values,
    )


def extract_header_fields_from_uploads(
    uploaded_files: Sequence[UploadedPdf],
    *,
    source_file_indexes: set[int] | None = None,
) -> list[PdfHeaderExtractionResult]:
    """Extract fixed top-section fields from each readable uploaded PDF page."""
    results: list[PdfHeaderExtractionResult] = []

    for file_index, uploaded_file in enumerate(uploaded_files):
        if source_file_indexes is not None and file_index not in source_file_indexes:
            continue

        source_filename = uploaded_file.name
        try:
            reader = PdfReader(BytesIO(read_uploaded_pdf_bytes(uploaded_file)))
        except Exception as exc:
            results.append(
                _header_result_from_fields(
                    source_filename=source_filename,
                    source_file_index=file_index,
                    page_number=0,
                    extraction_status="Failed",
                    error_message=str(exc),
                )
            )
            continue

        for page_index, page in enumerate(reader.pages):
            page_number = page_index + 1
            try:
                extracted_fields = _extract_page_header_fields(page)
                blank_count = sum(1 for value in extracted_fields.values() if value == "")
                status = "Extracted with blanks" if blank_count else "Extracted"
                results.append(
                    _header_result_from_fields(
                        source_filename=source_filename,
                        source_file_index=file_index,
                        page_number=page_number,
                        extraction_status=status,
                        error_message=None,
                        extracted_fields=extracted_fields,
                    )
                )
            except Exception as exc:
                results.append(
                    _header_result_from_fields(
                        source_filename=source_filename,
                        source_file_index=file_index,
                        page_number=page_number,
                        extraction_status="Failed",
                        error_message=str(exc),
                    )
                )

    return results
