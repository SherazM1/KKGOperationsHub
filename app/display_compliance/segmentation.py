"""Candidate region detection for Display Compliance.

This module will detect display/product-placement regions from a perfect
baseline image.
"""

from __future__ import annotations

from dataclasses import dataclass
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
    kept: list[BBox]
    removed_by_iou: int
    removed_by_coverage: int


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
    contours, proposals = _contour_bboxes(stages.morphology)
    geometry_result = _filter_bboxes_with_diagnostics(
        proposals,
        image_width=working_image.shape[1],
        image_height=working_image.shape[0],
        config=config,
    )
    dedup_result = _deduplicate_bboxes_with_diagnostics(geometry_result.filtered, config=config)
    original_bboxes = [
        _map_bbox_to_original(
            bbox,
            scale=scale,
            original_width=original_width,
            original_height=original_height,
        )
        for bbox in dedup_result.kept
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
    removed_by_deduplication = dedup_result.removed_by_iou + dedup_result.removed_by_coverage
    diagnostics = SegmentationDiagnostics(
        original_width=original_width,
        original_height=original_height,
        working_width=working_image.shape[1],
        working_height=working_image.shape[0],
        raw_contour_count=len(contours),
        raw_proposal_count=len(proposals),
        rejected_degenerate=geometry_result.rejected_degenerate,
        rejected_too_small=geometry_result.rejected_too_small,
        rejected_too_large=geometry_result.rejected_too_large,
        rejected_aspect_ratio=geometry_result.rejected_aspect_ratio,
        rejected_near_whole_image=geometry_result.rejected_near_whole_image,
        proposals_after_geometry_filter=len(geometry_result.filtered),
        deduplication_input_count=len(geometry_result.filtered),
        removed_by_iou_deduplication=dedup_result.removed_by_iou,
        removed_by_coverage_deduplication=dedup_result.removed_by_coverage,
        removed_by_deduplication=removed_by_deduplication,
        final_region_count=len(regions),
    )
    return SegmentationDiagnosticResult(
        regions=regions,
        diagnostics=diagnostics,
        diagnostic_images={
            "normalized": _encode_png(stages.normalized),
            "edges": _encode_png(stages.edges),
            "threshold": _encode_png(stages.threshold),
            "morphology": _encode_png(stages.morphology),
            "raw_proposals": _render_bbox_overlay(working_image, proposals),
            "filtered_proposals": _render_bbox_overlay(working_image, geometry_result.filtered),
            "final_proposals": _render_bbox_overlay(working_image, dedup_result.kept),
        },
        proposal_sample=_proposal_sample(
            geometry_result.filtered,
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
    return _deduplicate_bboxes_with_diagnostics(bboxes, config=config).kept


def _deduplicate_bboxes_with_diagnostics(
    bboxes: list[BBox],
    *,
    config: RegionDetectionConfig,
) -> _DeduplicationResult:
    ordered = sorted(bboxes, key=lambda bbox: bbox[2] * bbox[3], reverse=True)
    kept: list[BBox] = []
    removed_by_iou = 0
    removed_by_coverage = 0
    for candidate in ordered:
        rejected_by_iou = any(
            _bbox_iou(candidate, existing) >= config.overlap_iou_threshold for existing in kept
        )
        rejected_by_coverage = False
        if not rejected_by_iou:
            rejected_by_coverage = any(
                _smaller_coverage(candidate, existing) >= config.overlap_coverage_threshold
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
