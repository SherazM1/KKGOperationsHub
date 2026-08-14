"""Candidate region detection for Display Compliance.

This module will detect display/product-placement regions from a perfect
baseline image.
"""

from __future__ import annotations

from dataclasses import dataclass
from dataclasses import replace
from io import BytesIO
import math

import cv2
import numpy as np

from app.display_compliance.models import DisplayBaseline, ProductRegion


class DisplayComplianceSegmentationError(ValueError):
    """Raised when candidate-region detection cannot process an image."""


@dataclass(frozen=True)
class RegionDetectionConfig:
    """Deterministic thresholds for candidate product-region detection."""

    max_working_side: int = 1200
    min_relative_area: float = 0.0008
    max_relative_area: float = 0.82
    min_side_pixels: int = 10
    max_aspect_ratio: float = 8.0
    overlap_iou_threshold: float = 0.55
    overlap_coverage_threshold: float = 0.88
    morphology_kernel_size: int = 5
    canny_low_threshold: int = 40
    canny_high_threshold: int = 130


@dataclass(frozen=True)
class SegmentationDiagnostics:
    """Serializable counters for one candidate-region detection run.

    Rejection counts are mutually exclusive and follow the detector's existing
    geometry-filter order so raw_proposal_count reconciles with filtered output.
    """

    original_width: int
    original_height: int
    working_width: int
    working_height: int
    raw_contour_count: int
    raw_proposal_count: int
    rejected_degenerate: int
    rejected_too_small: int
    rejected_too_large: int
    rejected_aspect_ratio: int
    rejected_near_whole_image: int
    proposals_after_geometry_filter: int
    deduplication_input_count: int
    removed_by_iou_deduplication: int
    removed_by_coverage_deduplication: int
    removed_by_deduplication: int
    final_region_count: int
    strategy_a_raw_proposal_count: int
    strategy_a_proposals_after_geometry_filter: int
    strategy_a_rejected_degenerate: int
    strategy_a_rejected_too_small: int
    strategy_a_rejected_too_large: int
    strategy_a_rejected_aspect_ratio: int
    strategy_a_rejected_near_whole_image: int
    strategy_b_raw_proposal_count: int
    strategy_b_proposals_after_geometry_filter: int
    strategy_b_rejected_degenerate: int
    strategy_b_rejected_too_small: int
    strategy_b_rejected_too_large: int
    strategy_b_rejected_aspect_ratio: int
    strategy_b_rejected_near_whole_image: int
    strategy_c_raw_proposal_count: int
    strategy_c_proposals_after_geometry_filter: int
    strategy_c_rejected_degenerate: int
    strategy_c_rejected_too_small: int
    strategy_c_rejected_too_large: int
    strategy_c_rejected_aspect_ratio: int
    strategy_c_rejected_near_whole_image: int
    merged_pool_count_before_dedup: int
    size_cluster_count: int
    repeated_size_member_count: int
    alignment_boosted_count: int
    multi_strategy_supported_count: int


@dataclass(frozen=True)
class ProposalDiagnosticRow:
    """Small serializable sample of geometry-filtered proposal details."""

    proposal: int
    x: int
    y: int
    width: int
    height: int
    area_percent: float
    aspect_ratio: float


@dataclass(frozen=True)
class SegmentationDiagnosticResult:
    """Diagnostic candidate-region output without exposing OpenCV arrays."""

    regions: list[ProductRegion]
    diagnostics: SegmentationDiagnostics
    diagnostic_images: dict[str, bytes]
    proposal_sample: list[ProposalDiagnosticRow]


DEFAULT_DETECTION_CONFIG = RegionDetectionConfig()
BBox = tuple[int, int, int, int]


@dataclass(frozen=True)
class _ProposalStages:
    normalized: np.ndarray
    edges: np.ndarray
    threshold: np.ndarray
    morphology: np.ndarray


@dataclass(frozen=True)
class _GeometryFilterResult:
    filtered: list[BBox]
    rejected_degenerate: int
    rejected_too_small: int
    rejected_too_large: int
    rejected_aspect_ratio: int
    rejected_near_whole_image: int


@dataclass(frozen=True)
class _DeduplicationResult:
    kept: list["_Proposal"]
    removed_by_iou: int
    removed_by_coverage: int


@dataclass(frozen=True)
class _Proposal:
    bbox: BBox
    sources: tuple[str, ...]
    rectangularity: float = 0.0
    edge_support: float = 0.0
    size_cluster_size: int = 1
    row_alignment_size: int = 1
    column_alignment_size: int = 1
    score: float = 0.0


@dataclass(frozen=True)
class _StrategyResult:
    name: str
    raw_contour_count: int
    raw_proposals: list[_Proposal]
    geometry_result: _GeometryFilterResult


def detect_baseline_regions(*, baseline: DisplayBaseline, image_bytes: bytes) -> list[ProductRegion]:
    """Detect expected visual regions from a baseline image."""
    return detect_candidate_regions(image_bytes=image_bytes)


def detect_candidate_regions(
    *,
    image_bytes: bytes,
    config: RegionDetectionConfig = DEFAULT_DETECTION_CONFIG,
) -> list[ProductRegion]:
    """Detect product-agnostic candidate regions from a perfect baseline image."""
    return analyze_candidate_regions(image_bytes=image_bytes, config=config).regions


def analyze_candidate_regions(
    *,
    image_bytes: bytes,
    config: RegionDetectionConfig = DEFAULT_DETECTION_CONFIG,
    proposal_sample_limit: int = 50,
) -> SegmentationDiagnosticResult:
    """Detect product-agnostic candidate regions and return diagnostic evidence."""
    image = _decode_image(image_bytes)
    original_height, original_width = image.shape[:2]
    working_image, scale = _resize_for_detection(image, config.max_working_side)
    stages = _build_proposal_stages(working_image, config)
    strategy_results = [
        _strategy_a_morphology_contours(stages, working_image, config),
        _strategy_b_cleaned_edges(stages, working_image, config),
        _strategy_c_structural_rectangles(stages, working_image, config),
    ]
    merged_pool = _merge_strategy_proposals(strategy_results, stages.edges)
    scored_pool, scoring_summary = _score_structural_proposals(merged_pool)
    dedup_result = _deduplicate_proposals_with_diagnostics(scored_pool, config=config)
    original_bboxes = [
        _map_bbox_to_original(
            proposal.bbox,
            scale=scale,
            original_width=original_width,
            original_height=original_height,
        )
        for proposal in dedup_result.kept
    ]
    ordered_bboxes = _spatially_sort_bboxes(original_bboxes)
    regions = [
        ProductRegion(
            region_id=f"region_{index:03}",
            bbox=bbox,
            polygon=_bbox_polygon(bbox),
        )
        for index, bbox in enumerate(ordered_bboxes, start=1)
    ]
    strategy_a = strategy_results[0]
    strategy_b = strategy_results[1]
    strategy_c = strategy_results[2]
    strategy_a_geometry = strategy_a.geometry_result
    removed_by_deduplication = dedup_result.removed_by_iou + dedup_result.removed_by_coverage
    diagnostics = SegmentationDiagnostics(
        original_width=original_width,
        original_height=original_height,
        working_width=working_image.shape[1],
        working_height=working_image.shape[0],
        raw_contour_count=strategy_a.raw_contour_count,
        raw_proposal_count=len(strategy_a.raw_proposals),
        rejected_degenerate=strategy_a_geometry.rejected_degenerate,
        rejected_too_small=strategy_a_geometry.rejected_too_small,
        rejected_too_large=strategy_a_geometry.rejected_too_large,
        rejected_aspect_ratio=strategy_a_geometry.rejected_aspect_ratio,
        rejected_near_whole_image=strategy_a_geometry.rejected_near_whole_image,
        proposals_after_geometry_filter=len(strategy_a_geometry.filtered),
        deduplication_input_count=len(scored_pool),
        removed_by_iou_deduplication=dedup_result.removed_by_iou,
        removed_by_coverage_deduplication=dedup_result.removed_by_coverage,
        removed_by_deduplication=removed_by_deduplication,
        final_region_count=len(regions),
        strategy_a_raw_proposal_count=len(strategy_a.raw_proposals),
        strategy_a_proposals_after_geometry_filter=len(strategy_a.geometry_result.filtered),
        strategy_a_rejected_degenerate=strategy_a.geometry_result.rejected_degenerate,
        strategy_a_rejected_too_small=strategy_a.geometry_result.rejected_too_small,
        strategy_a_rejected_too_large=strategy_a.geometry_result.rejected_too_large,
        strategy_a_rejected_aspect_ratio=strategy_a.geometry_result.rejected_aspect_ratio,
        strategy_a_rejected_near_whole_image=strategy_a.geometry_result.rejected_near_whole_image,
        strategy_b_raw_proposal_count=len(strategy_b.raw_proposals),
        strategy_b_proposals_after_geometry_filter=len(strategy_b.geometry_result.filtered),
        strategy_b_rejected_degenerate=strategy_b.geometry_result.rejected_degenerate,
        strategy_b_rejected_too_small=strategy_b.geometry_result.rejected_too_small,
        strategy_b_rejected_too_large=strategy_b.geometry_result.rejected_too_large,
        strategy_b_rejected_aspect_ratio=strategy_b.geometry_result.rejected_aspect_ratio,
        strategy_b_rejected_near_whole_image=strategy_b.geometry_result.rejected_near_whole_image,
        strategy_c_raw_proposal_count=len(strategy_c.raw_proposals),
        strategy_c_proposals_after_geometry_filter=len(strategy_c.geometry_result.filtered),
        strategy_c_rejected_degenerate=strategy_c.geometry_result.rejected_degenerate,
        strategy_c_rejected_too_small=strategy_c.geometry_result.rejected_too_small,
        strategy_c_rejected_too_large=strategy_c.geometry_result.rejected_too_large,
        strategy_c_rejected_aspect_ratio=strategy_c.geometry_result.rejected_aspect_ratio,
        strategy_c_rejected_near_whole_image=strategy_c.geometry_result.rejected_near_whole_image,
        merged_pool_count_before_dedup=len(scored_pool),
        size_cluster_count=scoring_summary["size_cluster_count"],
        repeated_size_member_count=scoring_summary["repeated_size_member_count"],
        alignment_boosted_count=scoring_summary["alignment_boosted_count"],
        multi_strategy_supported_count=scoring_summary["multi_strategy_supported_count"],
    )
    return SegmentationDiagnosticResult(
        regions=regions,
        diagnostics=diagnostics,
        diagnostic_images={
            "normalized": _encode_png(stages.normalized),
            "edges": _encode_png(stages.edges),
            "threshold": _encode_png(stages.threshold),
            "morphology": _encode_png(stages.morphology),
            "strategy_a_proposals": _render_bbox_overlay(
                working_image,
                strategy_a.geometry_result.filtered,
            ),
            "strategy_b_proposals": _render_bbox_overlay(
                working_image,
                strategy_b.geometry_result.filtered,
            ),
            "strategy_c_proposals": _render_bbox_overlay(
                working_image,
                strategy_c.geometry_result.filtered,
            ),
            "merged_proposals": _render_bbox_overlay(
                working_image,
                [proposal.bbox for proposal in scored_pool],
            ),
            "raw_proposals": _render_bbox_overlay(
                working_image,
                [proposal.bbox for proposal in merged_pool],
            ),
            "filtered_proposals": _render_bbox_overlay(
                working_image,
                [proposal.bbox for proposal in scored_pool],
            ),
            "final_proposals": _render_bbox_overlay(
                working_image,
                [proposal.bbox for proposal in dedup_result.kept],
            ),
        },
        proposal_sample=_proposal_sample(
            [proposal.bbox for proposal in scored_pool],
            image_width=working_image.shape[1],
            image_height=working_image.shape[0],
            limit=proposal_sample_limit,
        ),
    )


def render_annotated_preview(
    *,
    image_bytes: bytes,
    regions: list[ProductRegion],
) -> bytes:
    """Render a PNG preview with detected candidate regions overlaid."""
    image = _decode_image(image_bytes)
    annotated = image.copy()
    for region in regions:
        if region.bbox is None:
            continue
        x, y, width, height = region.bbox
        cv2.rectangle(
            annotated,
            (x, y),
            (x + width, y + height),
            (0, 128, 255),
            max(2, round(min(image.shape[:2]) / 350)),
        )
        label = region.region_id.replace("region_", "")
        cv2.putText(
            annotated,
            label,
            (x + 4, max(y + 18, 18)),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.55,
            (0, 0, 0),
            3,
            cv2.LINE_AA,
        )
        cv2.putText(
            annotated,
            label,
            (x + 4, max(y + 18, 18)),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.55,
            (255, 255, 255),
            1,
            cv2.LINE_AA,
        )

    return _encode_png(annotated)


def _decode_image(image_bytes: bytes) -> np.ndarray:
    if not image_bytes:
        raise DisplayComplianceSegmentationError("A reference image is required for detection.")
    image_array = np.frombuffer(image_bytes, dtype=np.uint8)
    image = cv2.imdecode(image_array, cv2.IMREAD_COLOR)
    if image is None or image.size == 0:
        raise DisplayComplianceSegmentationError("OpenCV could not decode the reference image.")
    return image


def _resize_for_detection(image: np.ndarray, max_side: int) -> tuple[np.ndarray, float]:
    height, width = image.shape[:2]
    largest_side = max(width, height)
    if largest_side <= max_side:
        return image, 1.0
    scale = max_side / largest_side
    resized = cv2.resize(
        image,
        (round(width * scale), round(height * scale)),
        interpolation=cv2.INTER_AREA,
    )
    return resized, scale


def _build_proposal_mask(image: np.ndarray, config: RegionDetectionConfig) -> np.ndarray:
    return _build_proposal_stages(image, config).morphology


def _build_proposal_stages(image: np.ndarray, config: RegionDetectionConfig) -> _ProposalStages:
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    normalized = clahe.apply(gray)
    blurred = cv2.GaussianBlur(normalized, (3, 3), 0)

    edges = cv2.Canny(
        blurred,
        threshold1=config.canny_low_threshold,
        threshold2=config.canny_high_threshold,
    )
    adaptive = cv2.adaptiveThreshold(
        blurred,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY_INV,
        31,
        5,
    )
    combined = cv2.bitwise_or(edges, adaptive)
    kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT,
        (config.morphology_kernel_size, config.morphology_kernel_size),
    )
    closed = cv2.morphologyEx(combined, cv2.MORPH_CLOSE, kernel, iterations=2)
    morphology = cv2.dilate(closed, kernel, iterations=1)
    return _ProposalStages(
        normalized=normalized,
        edges=edges,
        threshold=adaptive,
        morphology=morphology,
    )


def _contour_bboxes(mask: np.ndarray) -> tuple[list[np.ndarray], list[BBox]]:
    contours, _hierarchy = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    return contours, [cv2.boundingRect(contour) for contour in contours]


def _strategy_a_morphology_contours(
    stages: _ProposalStages,
    image: np.ndarray,
    config: RegionDetectionConfig,
) -> _StrategyResult:
    contours, bboxes = _contour_bboxes(stages.morphology)
    raw_proposals = [
        _proposal_from_contour(
            contour,
            source="strategy_a_morphology",
            edges=stages.edges,
        )
        for contour in contours
    ]
    geometry_result = _filter_bboxes_with_diagnostics(
        bboxes,
        image_width=image.shape[1],
        image_height=image.shape[0],
        config=config,
    )
    return _StrategyResult(
        name="strategy_a_morphology",
        raw_contour_count=len(contours),
        raw_proposals=raw_proposals,
        geometry_result=geometry_result,
    )


def _strategy_b_cleaned_edges(
    stages: _ProposalStages,
    image: np.ndarray,
    config: RegionDetectionConfig,
) -> _StrategyResult:
    light_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    cleaned = cv2.dilate(stages.edges, light_kernel, iterations=1)
    cleaned = cv2.morphologyEx(cleaned, cv2.MORPH_CLOSE, light_kernel, iterations=1)
    contours, _hierarchy = cv2.findContours(cleaned, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    raw_proposals = [
        _proposal_from_contour(
            contour,
            source="strategy_b_cleaned_edges",
            edges=stages.edges,
        )
        for contour in contours
    ]
    bboxes = [proposal.bbox for proposal in raw_proposals]
    geometry_result = _filter_bboxes_with_diagnostics(
        bboxes,
        image_width=image.shape[1],
        image_height=image.shape[0],
        config=config,
    )
    return _StrategyResult(
        name="strategy_b_cleaned_edges",
        raw_contour_count=len(contours),
        raw_proposals=raw_proposals,
        geometry_result=geometry_result,
    )


def _strategy_c_structural_rectangles(
    stages: _ProposalStages,
    image: np.ndarray,
    config: RegionDetectionConfig,
) -> _StrategyResult:
    vertical_kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT,
        (1, max(8, image.shape[0] // 45)),
    )
    horizontal_kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT,
        (max(8, image.shape[1] // 45), 1),
    )
    vertical = cv2.morphologyEx(stages.edges, cv2.MORPH_CLOSE, vertical_kernel, iterations=1)
    horizontal = cv2.morphologyEx(stages.edges, cv2.MORPH_CLOSE, horizontal_kernel, iterations=1)
    structural = cv2.bitwise_or(vertical, horizontal)
    structural = cv2.dilate(
        structural,
        cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3)),
        iterations=1,
    )
    contours, _hierarchy = cv2.findContours(structural, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    raw_proposals = [
        _proposal_from_contour(
            contour,
            source="strategy_c_structural",
            edges=stages.edges,
        )
        for contour in contours
    ]
    bboxes = [proposal.bbox for proposal in raw_proposals]
    geometry_result = _filter_bboxes_with_diagnostics(
        bboxes,
        image_width=image.shape[1],
        image_height=image.shape[0],
        config=config,
    )
    return _StrategyResult(
        name="strategy_c_structural",
        raw_contour_count=len(contours),
        raw_proposals=raw_proposals,
        geometry_result=geometry_result,
    )


def _proposal_from_contour(
    contour: np.ndarray,
    *,
    source: str,
    edges: np.ndarray,
) -> _Proposal:
    bbox = cv2.boundingRect(contour)
    x, y, width, height = bbox
    contour_area = cv2.contourArea(contour)
    bbox_area = max(1, width * height)
    rectangularity = min(1.0, float(contour_area) / bbox_area)
    return _Proposal(
        bbox=bbox,
        sources=(source,),
        rectangularity=round(rectangularity, 4),
        edge_support=round(_edge_support(edges, bbox), 4),
    )


def _edge_support(edges: np.ndarray, bbox: BBox) -> float:
    x, y, width, height = bbox
    if width <= 0 or height <= 0:
        return 0.0
    roi = edges[y : y + height, x : x + width]
    if roi.size == 0:
        return 0.0
    perimeter_scale = max(1, 2 * (width + height))
    return min(1.0, float(cv2.countNonZero(roi)) / perimeter_scale)


def _merge_strategy_proposals(
    strategy_results: list[_StrategyResult],
    edges: np.ndarray,
) -> list[_Proposal]:
    merged: list[_Proposal] = []
    for result in strategy_results:
        proposal_by_bbox = {proposal.bbox: proposal for proposal in result.raw_proposals}
        for bbox in result.geometry_result.filtered:
            proposal = proposal_by_bbox.get(bbox)
            if proposal is None:
                proposal = _Proposal(
                    bbox=bbox,
                    sources=(result.name,),
                    edge_support=round(_edge_support(edges, bbox), 4),
                )
            merge_index = _find_multi_strategy_match(merged, proposal)
            if merge_index is None:
                merged.append(proposal)
            else:
                existing = merged[merge_index]
                sources = tuple(sorted(set(existing.sources + proposal.sources)))
                stronger = proposal if proposal.rectangularity > existing.rectangularity else existing
                merged[merge_index] = replace(
                    stronger,
                    sources=sources,
                    edge_support=max(existing.edge_support, proposal.edge_support),
                )
    return merged


def _find_multi_strategy_match(
    proposals: list[_Proposal],
    candidate: _Proposal,
) -> int | None:
    for index, proposal in enumerate(proposals):
        if set(proposal.sources) == set(candidate.sources):
            continue
        if _bbox_iou(proposal.bbox, candidate.bbox) >= 0.72:
            return index
        if _smaller_coverage(proposal.bbox, candidate.bbox) >= 0.92:
            return index
    return None


def _score_structural_proposals(
    proposals: list[_Proposal],
) -> tuple[list[_Proposal], dict[str, int]]:
    if not proposals:
        return proposals, {
            "size_cluster_count": 0,
            "repeated_size_member_count": 0,
            "alignment_boosted_count": 0,
            "multi_strategy_supported_count": 0,
        }

    size_clusters = _group_similar_sizes(proposals)
    row_clusters = _group_similar_centers(proposals, axis="y")
    column_clusters = _group_similar_centers(proposals, axis="x")
    size_cluster_by_index = _index_cluster_sizes(size_clusters)
    row_cluster_by_index = _index_cluster_sizes(row_clusters)
    column_cluster_by_index = _index_cluster_sizes(column_clusters)

    scored: list[_Proposal] = []
    for index, proposal in enumerate(proposals):
        size_cluster_size = size_cluster_by_index[index]
        row_alignment_size = row_cluster_by_index[index]
        column_alignment_size = column_cluster_by_index[index]
        alignment_size = max(row_alignment_size, column_alignment_size)
        repeated_bonus = min(3, size_cluster_size - 1) * 0.35
        alignment_bonus = min(3, alignment_size - 1) * 0.25
        source_bonus = (len(proposal.sources) - 1) * 0.45
        score = (
            proposal.rectangularity * 1.2
            + proposal.edge_support * 1.6
            + repeated_bonus
            + alignment_bonus
            + source_bonus
        )
        scored.append(
            replace(
                proposal,
                size_cluster_size=size_cluster_size,
                row_alignment_size=row_alignment_size,
                column_alignment_size=column_alignment_size,
                score=round(score, 4),
            )
        )

    scored.sort(
        key=lambda proposal: (
            -proposal.score,
            proposal.bbox[1],
            proposal.bbox[0],
            proposal.bbox[2] * proposal.bbox[3],
        )
    )
    return scored, {
        "size_cluster_count": sum(1 for cluster in size_clusters if len(cluster) > 1),
        "repeated_size_member_count": sum(
            1 for proposal in scored if proposal.size_cluster_size > 1
        ),
        "alignment_boosted_count": sum(
            1
            for proposal in scored
            if proposal.row_alignment_size > 1 or proposal.column_alignment_size > 1
        ),
        "multi_strategy_supported_count": sum(1 for proposal in scored if len(proposal.sources) > 1),
    }


def _group_similar_sizes(proposals: list[_Proposal]) -> list[list[int]]:
    clusters: list[list[int]] = []
    for index, proposal in enumerate(proposals):
        width = proposal.bbox[2]
        height = proposal.bbox[3]
        for cluster in clusters:
            cluster_width = sum(proposals[item].bbox[2] for item in cluster) / len(cluster)
            cluster_height = sum(proposals[item].bbox[3] for item in cluster) / len(cluster)
            width_tolerance = max(8, cluster_width * 0.18)
            height_tolerance = max(8, cluster_height * 0.18)
            if (
                abs(width - cluster_width) <= width_tolerance
                and abs(height - cluster_height) <= height_tolerance
            ):
                cluster.append(index)
                break
        else:
            clusters.append([index])
    return clusters


def _group_similar_centers(proposals: list[_Proposal], *, axis: str) -> list[list[int]]:
    clusters: list[list[int]] = []
    for index, proposal in enumerate(proposals):
        x, y, width, height = proposal.bbox
        center = x + width / 2 if axis == "x" else y + height / 2
        span = width if axis == "x" else height
        for cluster in clusters:
            cluster_center = sum(
                (
                    proposals[item].bbox[0] + proposals[item].bbox[2] / 2
                    if axis == "x"
                    else proposals[item].bbox[1] + proposals[item].bbox[3] / 2
                )
                for item in cluster
            ) / len(cluster)
            cluster_span = sum(
                proposals[item].bbox[2] if axis == "x" else proposals[item].bbox[3]
                for item in cluster
            ) / len(cluster)
            tolerance = max(10, min(span, cluster_span) * 0.35)
            if abs(center - cluster_center) <= tolerance:
                cluster.append(index)
                break
        else:
            clusters.append([index])
    return clusters


def _index_cluster_sizes(clusters: list[list[int]]) -> dict[int, int]:
    return {index: len(cluster) for cluster in clusters for index in cluster}


def _filter_bboxes(
    bboxes: list[BBox],
    *,
    image_width: int,
    image_height: int,
    config: RegionDetectionConfig,
) -> list[BBox]:
    return _filter_bboxes_with_diagnostics(
        bboxes,
        image_width=image_width,
        image_height=image_height,
        config=config,
    ).filtered


def _filter_bboxes_with_diagnostics(
    bboxes: list[BBox],
    *,
    image_width: int,
    image_height: int,
    config: RegionDetectionConfig,
) -> _GeometryFilterResult:
    image_area = image_width * image_height
    filtered: list[BBox] = []
    rejected_degenerate = 0
    rejected_too_small = 0
    rejected_too_large = 0
    rejected_aspect_ratio = 0
    rejected_near_whole_image = 0
    for x, y, width, height in bboxes:
        if width <= 0 or height <= 0:
            rejected_degenerate += 1
            continue
        if width <= config.min_side_pixels or height <= config.min_side_pixels:
            rejected_too_small += 1
            continue
        area_ratio = (width * height) / image_area
        if area_ratio < config.min_relative_area:
            rejected_too_small += 1
            continue
        if area_ratio > config.max_relative_area:
            rejected_too_large += 1
            continue
        aspect_ratio = max(width / height, height / width)
        if aspect_ratio > config.max_aspect_ratio:
            rejected_aspect_ratio += 1
            continue
        if width >= image_width * 0.97 and height >= image_height * 0.97:
            rejected_near_whole_image += 1
            continue
        filtered.append((x, y, width, height))
    return _GeometryFilterResult(
        filtered=filtered,
        rejected_degenerate=rejected_degenerate,
        rejected_too_small=rejected_too_small,
        rejected_too_large=rejected_too_large,
        rejected_aspect_ratio=rejected_aspect_ratio,
        rejected_near_whole_image=rejected_near_whole_image,
    )


def _deduplicate_bboxes(
    bboxes: list[BBox],
    *,
    config: RegionDetectionConfig,
) -> list[BBox]:
    return [proposal.bbox for proposal in _deduplicate_bboxes_with_diagnostics(bboxes, config=config).kept]


def _deduplicate_bboxes_with_diagnostics(
    bboxes: list[BBox],
    *,
    config: RegionDetectionConfig,
) -> _DeduplicationResult:
    proposals = [_Proposal(bbox=bbox, sources=("legacy",)) for bbox in bboxes]
    return _deduplicate_proposals_with_diagnostics(proposals, config=config)


def _deduplicate_proposals_with_diagnostics(
    proposals: list[_Proposal],
    *,
    config: RegionDetectionConfig,
) -> _DeduplicationResult:
    ordered = sorted(
        proposals,
        key=lambda proposal: (
            -proposal.score,
            -(proposal.bbox[2] * proposal.bbox[3]),
            proposal.bbox[1],
            proposal.bbox[0],
        ),
    )
    kept: list[_Proposal] = []
    removed_by_iou = 0
    removed_by_coverage = 0
    for candidate in ordered:
        rejected_by_iou = any(
            _bbox_iou(candidate.bbox, existing.bbox) >= config.overlap_iou_threshold
            for existing in kept
        )
        rejected_by_coverage = False
        if not rejected_by_iou:
            rejected_by_coverage = any(
                _smaller_coverage(candidate.bbox, existing.bbox)
                >= config.overlap_coverage_threshold
                for existing in kept
            )
        if rejected_by_iou:
            removed_by_iou += 1
            continue
        if rejected_by_coverage:
            removed_by_coverage += 1
            continue
        kept.append(candidate)
    return _DeduplicationResult(
        kept=kept,
        removed_by_iou=removed_by_iou,
        removed_by_coverage=removed_by_coverage,
    )


def _proposal_sample(
    bboxes: list[BBox],
    *,
    image_width: int,
    image_height: int,
    limit: int,
) -> list[ProposalDiagnosticRow]:
    image_area = image_width * image_height
    rows: list[ProposalDiagnosticRow] = []
    for index, (x, y, width, height) in enumerate(bboxes[:limit], start=1):
        rows.append(
            ProposalDiagnosticRow(
                proposal=index,
                x=x,
                y=y,
                width=width,
                height=height,
                area_percent=round((width * height / image_area) * 100, 4),
                aspect_ratio=round(max(width / height, height / width), 4),
            )
        )
    return rows


def _render_bbox_overlay(image: np.ndarray, bboxes: list[BBox]) -> bytes:
    overlay = image.copy()
    line_width = max(1, round(min(image.shape[:2]) / 350))
    label_limit = 120
    for index, (x, y, width, height) in enumerate(bboxes, start=1):
        cv2.rectangle(
            overlay,
            (x, y),
            (x + width, y + height),
            (0, 128, 255),
            line_width,
        )
        if index > label_limit:
            continue
        label = str(index)
        cv2.putText(
            overlay,
            label,
            (x + 3, max(y + 14, 14)),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.42,
            (0, 0, 0),
            2,
            cv2.LINE_AA,
        )
        cv2.putText(
            overlay,
            label,
            (x + 3, max(y + 14, 14)),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.42,
            (255, 255, 255),
            1,
            cv2.LINE_AA,
        )
    return _encode_png(overlay)


def _encode_png(image: np.ndarray) -> bytes:
    success, encoded = cv2.imencode(".png", image)
    if not success:
        raise DisplayComplianceSegmentationError("Could not render diagnostic preview.")
    return BytesIO(encoded.tobytes()).getvalue()


def _bbox_iou(first: BBox, second: BBox) -> float:
    intersection = _intersection_area(first, second)
    if intersection == 0:
        return 0.0
    first_area = first[2] * first[3]
    second_area = second[2] * second[3]
    return intersection / (first_area + second_area - intersection)


def _smaller_coverage(first: BBox, second: BBox) -> float:
    intersection = _intersection_area(first, second)
    if intersection == 0:
        return 0.0
    smaller_area = min(first[2] * first[3], second[2] * second[3])
    return intersection / smaller_area


def _intersection_area(first: BBox, second: BBox) -> int:
    left = max(first[0], second[0])
    top = max(first[1], second[1])
    right = min(first[0] + first[2], second[0] + second[2])
    bottom = min(first[1] + first[3], second[1] + second[3])
    if right <= left or bottom <= top:
        return 0
    return (right - left) * (bottom - top)


def _map_bbox_to_original(
    bbox: BBox,
    *,
    scale: float,
    original_width: int,
    original_height: int,
) -> BBox:
    if math.isclose(scale, 1.0):
        mapped = bbox
    else:
        x, y, width, height = bbox
        mapped = (
            round(x / scale),
            round(y / scale),
            round(width / scale),
            round(height / scale),
        )
    return _clamp_bbox(mapped, image_width=original_width, image_height=original_height)


def _clamp_bbox(bbox: BBox, *, image_width: int, image_height: int) -> BBox:
    x, y, width, height = bbox
    x = max(0, min(x, image_width - 1))
    y = max(0, min(y, image_height - 1))
    right = max(x + 1, min(x + width, image_width))
    bottom = max(y + 1, min(y + height, image_height))
    return x, y, right - x, bottom - y


def _spatially_sort_bboxes(bboxes: list[BBox]) -> list[BBox]:
    if not bboxes:
        return []
    median_height = sorted(bbox[3] for bbox in bboxes)[len(bboxes) // 2]
    row_tolerance = max(12, median_height * 0.55)
    rows: list[list[BBox]] = []

    for bbox in sorted(bboxes, key=lambda item: (item[1] + item[3] / 2, item[0])):
        center_y = bbox[1] + bbox[3] / 2
        for row in rows:
            row_center = sum(item[1] + item[3] / 2 for item in row) / len(row)
            if abs(center_y - row_center) <= row_tolerance:
                row.append(bbox)
                break
        else:
            rows.append([bbox])

    rows.sort(key=lambda row: sum(item[1] + item[3] / 2 for item in row) / len(row))
    ordered: list[BBox] = []
    for row in rows:
        ordered.extend(sorted(row, key=lambda item: item[0]))
    return ordered


def _bbox_polygon(bbox: BBox) -> tuple[tuple[int, int], ...]:
    x, y, width, height = bbox
    return (
        (x, y),
        (x + width, y),
        (x + width, y + height),
        (x, y + height),
    )
