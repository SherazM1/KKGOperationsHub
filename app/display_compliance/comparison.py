"""Future inspection-comparison boundary for Display Compliance.

This module will compare aligned observed regions against baseline expected
regions.
"""

from __future__ import annotations

from app.display_compliance.models import DisplayBaseline, InspectionResult


def inspect_image(*, baseline: DisplayBaseline, image_bytes: bytes) -> InspectionResult:
    """Inspect one image against a selected display baseline."""
    raise NotImplementedError("Display Compliance inspection comparison is not implemented yet.")

