"""Future region-detection boundary for Display Compliance.

This module will detect display/product-placement regions from a perfect
baseline image.
"""

from __future__ import annotations

from app.display_compliance.models import DisplayBaseline, ProductRegion


def detect_baseline_regions(*, baseline: DisplayBaseline, image_bytes: bytes) -> list[ProductRegion]:
    """Detect expected visual regions from a baseline image."""
    raise NotImplementedError("Display Compliance region detection is not implemented yet.")

