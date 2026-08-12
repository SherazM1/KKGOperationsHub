"""Future image-alignment boundary for Display Compliance.

This module will align an inspection image geometrically to its selected
baseline before region comparison.
"""

from __future__ import annotations

from app.display_compliance.models import DisplayBaseline


def align_to_baseline(*, baseline: DisplayBaseline, image_bytes: bytes) -> bytes:
    """Align an inspection image to the baseline coordinate space."""
    raise NotImplementedError("Display Compliance image alignment is not implemented yet.")

