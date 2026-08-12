"""Future scoring boundary for Display Compliance.

This module will convert region-comparison signals into PASS / REVIEW / FAIL
decisions.
"""

from __future__ import annotations

from app.display_compliance.models import InspectionIssue, InspectionResult


def score_inspection(*, issues: list[InspectionIssue]) -> InspectionResult:
    """Score inspection issues into a future PASS / REVIEW / FAIL result."""
    raise NotImplementedError("Display Compliance scoring is not implemented yet.")

