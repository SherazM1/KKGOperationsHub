"""PDF upload inventory helpers for the Spec Sheet Extractor module."""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO, StringIO
import contextlib
import re
from typing import Any, Protocol, Sequence

from pypdf import PdfReader
from pypdf._text_extraction import mult

from app.spec_sheet_extractor.models import (
    FIELD_ATTRIBUTE_BY_LABEL,
    SPEC_SHEET_FIXED_FIELDS,
    PdfFileInventoryResult,
    PdfHeaderExtractionResult,
    PdfPageInventoryRecord,
)
from app.spec_sheet_extractor.zones import (
    HEADER_FIELD_ZONES,
    HEADER_REGION_ZONE,
    NormalizedTextZone,
)


class UploadedPdf(Protocol):
    """Minimal interface required from Streamlit uploaded files."""

    name: str

    def getvalue(self) -> bytes:
        """Return uploaded file bytes."""


@dataclass(frozen=True)
class TextFragment:
    """A positioned text fragment extracted from a PDF page."""

    text: str
    raw_x: float
    raw_y: float
    x: float
    y: float
    left: float
    bottom: float
    right: float
    top: float
    current_matrix: tuple[float, ...]
    text_matrix: tuple[float, ...]
    transformed_matrix: tuple[float, ...]


@dataclass(frozen=True)
class HeaderFieldExtraction:
    """Extracted header fields plus non-fatal extraction warnings."""

    fields: dict[str, str]
    warnings: tuple[str, ...] = ()


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
        r"Oppty\s*/\s*Proj\.?\s*#",
        r"Oppty\s*/\s*Project\s*#",
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

_LABEL_LOOKAHEAD = "|".join(
    rf"(?:{label_pattern})\s*:"
    for patterns in _FIELD_LABEL_PATTERNS.values()
    for label_pattern in patterns
)
_STRICT_DATE_PATTERN = re.compile(r"\b\d{2}/\d{2}/\d{4}\b")


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
        stripped_value = re.sub(
            rf"^\s*{label_pattern}\s*:?\s*",
            "",
            value,
            count=1,
            flags=re.IGNORECASE,
        )
        if stripped_value != value:
            return _clean_extracted_field_value(field_name, stripped_value)
    return ""


def _clean_extracted_field_value(field_name: str, value: str) -> str:
    if field_name == "Date":
        return _extract_strict_date(value)
    if field_name == "Inches of rule":
        return _trim_trailing_known_label(value, "Date")
    return _trim_contaminating_labels(value)


def _extract_strict_date(value: str) -> str:
    match = _STRICT_DATE_PATTERN.search(_clean_pdf_text(value))
    return match.group(0) if match else ""


def _trim_trailing_known_label(value: str, label_field_name: str) -> str:
    cleaned = _clean_pdf_text(value)
    for label_pattern in _FIELD_LABEL_PATTERNS[label_field_name]:
        cleaned = re.sub(
            rf"\s+(?:{label_pattern})\s*:\s*$",
            "",
            cleaned,
            count=1,
            flags=re.IGNORECASE,
        )
    return cleaned.strip()


def _trim_contaminating_labels(value: str) -> str:
    cleaned = _clean_pdf_text(value)
    for label_pattern in _FIELD_LABEL_PATTERNS.values():
        for pattern in label_pattern:
            match = re.search(rf"\s+(?:{pattern})\s*:", cleaned, flags=re.IGNORECASE)
            if match:
                return cleaned[: match.start()].strip()
    return cleaned


def _text_box_for_fragment(
    text: str,
    transformed_matrix: Sequence[float],
    font_size: float,
) -> tuple[float, float, float, float, float, float]:
    x = float(transformed_matrix[4])
    y = float(transformed_matrix[5])
    scale_x = max(abs(float(transformed_matrix[0])), abs(float(transformed_matrix[2])), 1.0)
    scale_y = max(abs(float(transformed_matrix[1])), abs(float(transformed_matrix[3])), 1.0)
    height = max(float(font_size) * scale_y, 1.0)
    width = max(len(_clean_pdf_text(text)) * float(font_size) * 0.55 * scale_x, 1.0)
    return x, y, x, y - height * 0.25, x + width, y + height


def _raw_page_dimensions(page: object) -> tuple[float, float]:
    cropbox = getattr(page, "cropbox", None)
    box = cropbox or page.mediabox
    return float(box.width), float(box.height)


def _display_page_dimensions(page: object) -> tuple[float, float]:
    raw_width, raw_height = _raw_page_dimensions(page)
    rotation = int(getattr(page, "rotation", 0) or 0) % 360
    if rotation in (90, 270):
        return raw_height, raw_width
    return raw_width, raw_height


def _map_point_to_display(page: object, x: float, y: float) -> tuple[float, float]:
    raw_width, raw_height = _raw_page_dimensions(page)
    rotation = int(getattr(page, "rotation", 0) or 0) % 360
    if rotation == 90:
        return y, raw_width - x
    if rotation == 180:
        return raw_width - x, raw_height - y
    if rotation == 270:
        return raw_height - y, x
    return x, y


def _map_box_to_display(
    page: object,
    left: float,
    bottom: float,
    right: float,
    top: float,
) -> tuple[float, float, float, float]:
    points = [
        _map_point_to_display(page, left, bottom),
        _map_point_to_display(page, left, top),
        _map_point_to_display(page, right, bottom),
        _map_point_to_display(page, right, top),
    ]
    xs = [point[0] for point in points]
    ys = [point[1] for point in points]
    return min(xs), min(ys), max(xs), max(ys)


def _iter_text_fragments(page: object) -> list[TextFragment]:
    fragments: list[TextFragment] = []

    def visitor_text(
        text: str,
        cm: Sequence[float],
        tm: Sequence[float],
        _font: object,
        font_size: float,
    ) -> None:
        cleaned = _clean_pdf_text(text)
        if not cleaned:
            return
        transformed = mult(tm, cm)
        raw_x, raw_y, raw_left, raw_bottom, raw_right, raw_top = _text_box_for_fragment(
            cleaned,
            transformed,
            font_size,
        )
        x, y = _map_point_to_display(page, raw_x, raw_y)
        left, bottom, right, top = _map_box_to_display(
            page,
            raw_left,
            raw_bottom,
            raw_right,
            raw_top,
        )
        fragments.append(
            TextFragment(
                text=cleaned,
                raw_x=raw_x,
                raw_y=raw_y,
                x=x,
                y=y,
                left=left,
                bottom=bottom,
                right=right,
                top=top,
                current_matrix=tuple(float(value) for value in cm),
                text_matrix=tuple(float(value) for value in tm),
                transformed_matrix=tuple(float(value) for value in transformed),
            )
        )

    page.extract_text(visitor_text=visitor_text)
    return fragments


def _fragments_in_zone(page: object, zone: NormalizedTextZone) -> list[TextFragment]:
    page_width, page_height = _display_page_dimensions(page)
    return [
        fragment
        for fragment in _iter_text_fragments(page)
        if (
            zone.contains(fragment.x, fragment.y, page_width, page_height)
            or zone.contains(
                (fragment.left + fragment.right) / 2,
                (fragment.bottom + fragment.top) / 2,
                page_width,
                page_height,
            )
        )
    ]


def _sort_fragments_for_reading(fragments: list[TextFragment]) -> list[TextFragment]:
    return sorted(fragments, key=lambda fragment: (-round(fragment.y, 1), fragment.x))


def _collect_zone_text(page: object, zone: NormalizedTextZone) -> str:
    fragments = _sort_fragments_for_reading(_fragments_in_zone(page, zone))
    return _clean_pdf_text(" ".join(fragment.text for fragment in fragments))


def _top_region_text_from_fragments(page: object) -> str:
    fragments = _sort_fragments_for_reading(_fragments_in_zone(page, HEADER_REGION_ZONE))
    return "\n".join(fragment.text for fragment in fragments)


def _top_region_layout_text(page: object) -> str:
    try:
        with contextlib.redirect_stdout(StringIO()), contextlib.redirect_stderr(StringIO()):
            text = page.extract_text(extraction_mode="layout") or ""
    except Exception:
        text = page.extract_text() or ""
    lines = text.splitlines()
    if not lines:
        return ""
    line_limit = max(1, int(len(lines) * 0.35))
    return "\n".join(lines[:line_limit])


def _parse_labeled_header_text(text: str) -> dict[str, str]:
    parsed = {label: "" for label in SPEC_SHEET_FIXED_FIELDS}
    normalized = _clean_pdf_text(text)
    if not normalized:
        return parsed

    for field_name in SPEC_SHEET_FIXED_FIELDS:
        for label_pattern in _FIELD_LABEL_PATTERNS[field_name]:
            match = re.search(
                rf"(?:^|\s){label_pattern}\s*:?\s*(.*?)\s*(?=(?:{_LABEL_LOOKAHEAD})(?:\s|:)|$)",
                normalized,
                flags=re.IGNORECASE,
            )
            if match:
                parsed[field_name] = _clean_extracted_field_value(field_name, match.group(1))
                break
    return parsed


def _merge_fallback_fields(
    primary: dict[str, str],
    fallback: dict[str, str],
    warnings: list[str],
) -> dict[str, str]:
    merged = {
        field_name: _clean_extracted_field_value(field_name, value)
        for field_name, value in primary.items()
    }
    for field_name, value in fallback.items():
        if value and (not merged.get(field_name) or _value_contains_field_label(merged[field_name])):
            if merged.get(field_name) and merged[field_name] != value:
                warnings.append(f"{field_name} was corrected by fallback parsing.")
            merged[field_name] = value
    return merged


def _value_contains_field_label(value: str) -> bool:
    return any(
        re.search(rf"(?:^|\s){label_pattern}\s*:?", value, flags=re.IGNORECASE)
        for patterns in _FIELD_LABEL_PATTERNS.values()
        for label_pattern in patterns
    )


def _extract_page_header_fields(page: object) -> HeaderFieldExtraction:
    extracted_fields: dict[str, str] = {}
    warnings: list[str] = []
    for zone in HEADER_FIELD_ZONES:
        zone_text = _collect_zone_text(page, zone)
        extracted_fields[zone.field_name] = _strip_printed_field_label(zone.field_name, zone_text)

    top_region_text = _top_region_text_from_fragments(page)
    fallback_fields = _parse_labeled_header_text(top_region_text)
    extracted_fields = _merge_fallback_fields(extracted_fields, fallback_fields, warnings)
    if not extracted_fields.get("Date"):
        extracted_fields["Date"] = _extract_strict_date(top_region_text)
    if int(getattr(page, "rotation", 0) or 0) % 360 == 0:
        layout_text = _top_region_layout_text(page)
        layout_fallback_fields = _parse_labeled_header_text(layout_text)
        extracted_fields = _merge_fallback_fields(extracted_fields, layout_fallback_fields, warnings)
        if not extracted_fields.get("Date"):
            extracted_fields["Date"] = _extract_strict_date(layout_text)
    return HeaderFieldExtraction(extracted_fields, tuple(warnings))


def inspect_pdf_page_text_structure(pdf_bytes: bytes, page_number: int = 1) -> dict[str, Any]:
    """Return developer diagnostics for one PDF page's text structure."""
    reader = PdfReader(BytesIO(pdf_bytes))
    page = reader.pages[page_number - 1]
    fragments = _iter_text_fragments(page)
    return {
        "page_number": page_number,
        "mediabox": tuple(float(value) for value in page.mediabox),
        "cropbox": tuple(float(value) for value in page.cropbox),
        "rotation": int(page.rotation or 0),
        "fragment_count": len(fragments),
        "fragments": [
            {
                "text": fragment.text,
                "x": fragment.x,
                "y": fragment.y,
                "raw_x": fragment.raw_x,
                "raw_y": fragment.raw_y,
                "bbox": (fragment.left, fragment.bottom, fragment.right, fragment.top),
                "current_matrix": fragment.current_matrix,
                "text_matrix": fragment.text_matrix,
                "transformed_matrix": fragment.transformed_matrix,
            }
            for fragment in _sort_fragments_for_reading(fragments)
        ],
        "top_region_text": _top_region_text_from_fragments(page),
        "layout_top_text": _top_region_layout_text(page),
    }


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
                extraction = _extract_page_header_fields(page)
                extracted_fields = extraction.fields
                populated_count = sum(1 for value in extracted_fields.values() if value)
                if populated_count == 0:
                    status = "Failed extraction"
                    error_message = "No fixed header fields could be extracted."
                else:
                    status = "Extracted with blanks" if extraction.warnings else "Extracted"
                    error_message = " | ".join(extraction.warnings) if extraction.warnings else None
                results.append(
                    _header_result_from_fields(
                        source_filename=source_filename,
                        source_file_index=file_index,
                        page_number=page_number,
                        extraction_status=status,
                        error_message=error_message,
                        extracted_fields=extracted_fields,
                    )
                )
            except Exception as exc:
                results.append(
                    _header_result_from_fields(
                        source_filename=source_filename,
                        source_file_index=file_index,
                        page_number=page_number,
                        extraction_status="Failed extraction",
                        error_message=str(exc),
                    )
                )

    return results
