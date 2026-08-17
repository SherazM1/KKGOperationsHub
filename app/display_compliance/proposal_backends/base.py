"""Generic learned region-proposal backend contracts."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol

import numpy as np

from app.display_compliance.models import ProductRegion

BBox = tuple[int, int, int, int]
Polygon = tuple[tuple[int, int], ...]


class RegionProposalBackendUnavailable(ValueError):
    """Raised when a learned proposal backend cannot run in this environment."""


@dataclass(frozen=True)
class LearnedProposal:
    """Product-agnostic proposal from a learned segmentation backend."""

    bbox: BBox
    polygon: Polygon
    area: int
    predicted_iou: float | None
    stability_score: float | None
    source_backend: str
    mask: np.ndarray | None = None
    solidity: float = 0.0
    size_cluster_size: int = 1
    row_alignment_size: int = 1
    column_alignment_size: int = 1
    score: float = 0.0


@dataclass(frozen=True)
class LearnedSegmentationDiagnostics:
    """Serializable counters for one learned proposal run."""

    backend: str
    original_width: int
    original_height: int
    working_width: int
    working_height: int
    device: str
    model_config: str
    checkpoint: str
    model_load_seconds: float
    inference_seconds: float
    total_seconds: float
    raw_mask_count: int
    rejected_degenerate: int
    rejected_too_small: int
    rejected_too_large: int
    rejected_aspect_ratio: int
    rejected_low_confidence: int
    rejected_low_solidity: int
    proposals_after_basic_filtering: int
    removed_by_iou_deduplication: int
    removed_by_nested_deduplication: int
    removed_by_deduplication: int
    proposals_after_duplicate_cleanup: int
    size_cluster_count: int
    repeated_size_member_count: int
    alignment_supported_count: int
    final_region_count: int


@dataclass(frozen=True)
class RegionProposalResult:
    """Complete learned proposal result, including lightweight diagnostics."""

    regions: list[ProductRegion]
    diagnostics: LearnedSegmentationDiagnostics
    diagnostic_images: dict[str, bytes]
    proposals: list[LearnedProposal]


class RegionProposalBackend(Protocol):
    """Backend boundary for product-agnostic region proposals."""

    backend_name: str

    def propose(self, *, image_bytes: bytes) -> RegionProposalResult:
        """Generate product-agnostic candidate regions from one baseline image."""
