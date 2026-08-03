"""Data models for the Spec Sheet Extractor module."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class PdfFileInventoryResult:
    """Inventory result for one uploaded spec-sheet PDF."""

    source_filename: str
    source_file_index: int
    page_count: int
    status: str
    error_message: str | None = None


@dataclass(frozen=True)
class PdfPageInventoryRecord:
    """Inventory record for one source PDF page."""

    source_filename: str
    source_file_index: int
    page_number: int
    status: str
    error_message: str | None = None
