"""Data models for the Spec Sheet Extractor module."""

from __future__ import annotations

from dataclasses import dataclass


SPEC_SHEET_FIXED_FIELDS: tuple[str, ...] = (
    "Customer",
    "Design",
    "Revision",
    "Part",
    "Opportunity/Project #",
    "Pieces per set",
    "Board",
    "Corr direction",
    "View",
    "Production/Project Manager",
    "Designer",
    "ID",
    "Area",
    "Blank width",
    "Blank height",
    "Inches of rule",
    "Date",
)

FIELD_ATTRIBUTE_BY_LABEL: dict[str, str] = {
    "Customer": "customer",
    "Design": "design",
    "Revision": "revision",
    "Part": "part",
    "Opportunity/Project #": "opportunity_project_number",
    "Pieces per set": "pieces_per_set",
    "Board": "board",
    "Corr direction": "corr_direction",
    "View": "view",
    "Production/Project Manager": "production_project_manager",
    "Designer": "designer",
    "ID": "id",
    "Area": "area",
    "Blank width": "blank_width",
    "Blank height": "blank_height",
    "Inches of rule": "inches_of_rule",
    "Date": "date",
}


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


@dataclass(frozen=True)
class PdfHeaderExtractionResult:
    """Fixed header-field extraction result for one source PDF page."""

    source_filename: str
    source_file_index: int
    page_number: int
    extraction_status: str
    error_message: str | None
    customer: str = ""
    design: str = ""
    revision: str = ""
    part: str = ""
    opportunity_project_number: str = ""
    pieces_per_set: str = ""
    board: str = ""
    corr_direction: str = ""
    view: str = ""
    production_project_manager: str = ""
    designer: str = ""
    id: str = ""
    area: str = ""
    blank_width: str = ""
    blank_height: str = ""
    inches_of_rule: str = ""
    date: str = ""
    upper_special_text: str = ""
    lower_special_text: str = ""

    def to_preview_row(self) -> dict[str, str | int | None]:
        """Return a UI-friendly row with final field labels."""
        row: dict[str, str | int | None] = {
            "Source File": self.source_filename,
            "Page Number": self.page_number,
            "Status": self.extraction_status,
            "Error": self.error_message or "",
        }
        for label in SPEC_SHEET_FIXED_FIELDS:
            row[label] = getattr(self, FIELD_ATTRIBUTE_BY_LABEL[label])
        row["Upper Special Text"] = self.upper_special_text
        row["Lower Special Text"] = self.lower_special_text
        return row
