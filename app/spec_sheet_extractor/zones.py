"""Zone definitions for the Spec Sheet Extractor module."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class NormalizedTextZone:
    """A rectangle in normalized PDF page coordinates.

    Coordinates are fractions of the page mediabox after reading text positions
    from the PDF. Values use the PDF origin convention: left/bottom is 0.0 and
    right/top is 1.0. Keeping zones normalized lets the same logical field map
    apply across supported page sizes and orientations.
    """

    field_name: str
    left: float
    bottom: float
    right: float
    top: float

    def contains(self, x: float, y: float, page_width: float, page_height: float) -> bool:
        """Return True when a PDF text point falls inside this zone."""
        normalized_x = x / page_width if page_width else 0.0
        normalized_y = y / page_height if page_height else 0.0
        return (
            self.left <= normalized_x <= self.right
            and self.bottom <= normalized_y <= self.top
        )

    def intersects(
        self,
        left: float,
        bottom: float,
        right: float,
        top: float,
        page_width: float,
        page_height: float,
    ) -> bool:
        """Return True when a text span box overlaps this zone."""
        normalized_left = left / page_width if page_width else 0.0
        normalized_right = right / page_width if page_width else 0.0
        normalized_bottom = bottom / page_height if page_height else 0.0
        normalized_top = top / page_height if page_height else 0.0
        return not (
            normalized_right < self.left
            or normalized_left > self.right
            or normalized_top < self.bottom
            or normalized_bottom > self.top
        )


# Zones target only the fixed top header section of the Kendal King spec sheet.
# They intentionally stop above the drawing/body area so dieline measurements do
# not become candidates for fixed header fields.
HEADER_FIELD_ZONES: tuple[NormalizedTextZone, ...] = (
    NormalizedTextZone("Customer", 0.05, 0.925, 0.34, 0.965),
    NormalizedTextZone("Design", 0.34, 0.925, 0.56, 0.965),
    NormalizedTextZone("Revision", 0.56, 0.925, 0.72, 0.965),
    NormalizedTextZone("Part", 0.72, 0.925, 0.95, 0.965),
    NormalizedTextZone("Opportunity/Project #", 0.05, 0.875, 0.34, 0.918),
    NormalizedTextZone("Pieces per set", 0.34, 0.875, 0.56, 0.918),
    NormalizedTextZone("Board", 0.56, 0.875, 0.72, 0.918),
    NormalizedTextZone("Corr direction", 0.72, 0.875, 0.95, 0.918),
    NormalizedTextZone("View", 0.05, 0.825, 0.24, 0.868),
    NormalizedTextZone("Production/Project Manager", 0.24, 0.825, 0.52, 0.868),
    NormalizedTextZone("Designer", 0.52, 0.825, 0.72, 0.868),
    NormalizedTextZone("ID", 0.72, 0.825, 0.95, 0.868),
    NormalizedTextZone("Area", 0.05, 0.775, 0.24, 0.818),
    NormalizedTextZone("Blank width", 0.24, 0.775, 0.42, 0.818),
    NormalizedTextZone("Blank height", 0.42, 0.775, 0.60, 0.818),
    NormalizedTextZone("Inches of rule", 0.60, 0.775, 0.80, 0.818),
    NormalizedTextZone("Date", 0.80, 0.775, 0.95, 0.818),
)

HEADER_REGION_ZONE = NormalizedTextZone("Header Region", 0.0, 0.72, 1.0, 1.0)

# Visual instruction area below the fixed header and above the lower caption band.
# This broad zone catches color/print/material notes while excluding the fixed
# header rows at the top of the page.
UPPER_SPECIAL_TEXT_ZONE = NormalizedTextZone(
    "Upper Special Text",
    0.0,
    0.13,
    1.0,
    0.86,
)

# Visual lower caption/instruction band above the copyright footer. The bottom
# extends slightly below 0 because some production PDFs emit bottom captions with
# transformed coordinates just outside the visible mediabox.
LOWER_SPECIAL_TEXT_ZONE = NormalizedTextZone(
    "Lower Special Text",
    0.0,
    -0.60,
    1.0,
    0.24,
)
