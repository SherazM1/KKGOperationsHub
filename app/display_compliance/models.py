"""Domain models for Display Compliance."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class ProductRegion:
    """Expected visual or product-placement region in a baseline display."""

    region_id: str
    bbox: tuple[int, int, int, int] | None = None
    polygon: tuple[tuple[int, int], ...] = ()
    label: str = ""
    visual_signature: str | None = None


@dataclass(frozen=True)
class DisplayBaseline:
    """Approved reference display used as the source of truth."""

    baseline_id: str
    name: str
    reference_filename: str
    reference_width: int
    reference_height: int
    regions: list[ProductRegion] = field(default_factory=list)


@dataclass(frozen=True)
class InspectionIssue:
    """Future-compatible record of a display inspection issue."""

    region_id: str
    issue_type: str
    confidence: float
    message: str


@dataclass(frozen=True)
class InspectionResult:
    """Future-compatible result for an inspection image."""

    status: str
    confidence: float
    issues: list[InspectionIssue] = field(default_factory=list)

