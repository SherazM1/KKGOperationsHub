"""Region proposal backends for Display Compliance."""

from app.display_compliance.proposal_backends.base import (
    LearnedProposal,
    LearnedSegmentationDiagnostics,
    RegionProposalBackend,
    RegionProposalBackendUnavailable,
    RegionProposalResult,
)
from app.display_compliance.proposal_backends.sam2_backend import (
    DEFAULT_LEARNED_MAX_SIDE_ENV,
    DEFAULT_SAM2_CACHE_DIR_ENV,
    DEFAULT_SAM2_CHECKPOINT_ENV,
    DEFAULT_SAM2_MODEL_CONFIG,
    DEFAULT_SAM2_MODEL_CONFIG_ENV,
    DEFAULT_SAM2_TINY_CHECKPOINT_URL,
    Sam2AutomaticMaskBackend,
    analyze_learned_candidate_regions,
)

__all__ = [
    "DEFAULT_LEARNED_MAX_SIDE_ENV",
    "DEFAULT_SAM2_CACHE_DIR_ENV",
    "DEFAULT_SAM2_CHECKPOINT_ENV",
    "DEFAULT_SAM2_MODEL_CONFIG",
    "DEFAULT_SAM2_MODEL_CONFIG_ENV",
    "DEFAULT_SAM2_TINY_CHECKPOINT_URL",
    "LearnedProposal",
    "LearnedSegmentationDiagnostics",
    "RegionProposalBackend",
    "RegionProposalBackendUnavailable",
    "RegionProposalResult",
    "Sam2AutomaticMaskBackend",
    "analyze_learned_candidate_regions",
]
