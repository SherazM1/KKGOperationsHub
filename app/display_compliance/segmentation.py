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


DEFAULT_DETECTION_CONFIG = RegionDetectionConfig()
BBox = tuple[int, int, int, int]


def detect_baseline_regions(*, baseline: DisplayBaseline, image_bytes: bytes) -> list[ProductRegion]:
    """Detect expected visual regions from a baseline image."""
    return detect_candidate_regions(image_bytes=image_bytes)


def detect_candidate_regions(
    *,
    image_bytes: bytes,
    config: RegionDetectionConfig = DEFAULT_DETECTION_CONFIG,
) -> list[ProductRegion]:
    """Detect product-agnostic candidate regions from a perfect baseline image."""
    image = _decode_image(image_bytes)
    original_height, original_width = image.shape[:2]
    working_image, scale = _resize_for_detection(image, config.max_working_side)
    mask = _build_proposal_mask(working_image, config)
    proposals = _contour_bboxes(mask)
    filtered = _filter_bboxes(
        proposals,
        image_width=working_image.shape[1],
        image_height=working_image.shape[0],
        config=config,
    )
    deduped = _deduplicate_bboxes(filtered, config=config)
    original_bboxes = [
        _map_bbox_to_original(
            bbox,
            scale=scale,
            original_width=original_width,
            original_height=original_height,
        )
        for bbox in deduped
    ]
    ordered_bboxes = _spatially_sort_bboxes(original_bboxes)

    return [
        ProductRegion(
            region_id=f"region_{index:03}",
            bbox=bbox,
            polygon=_bbox_polygon(bbox),
        )
        for index, bbox in enumerate(ordered_bboxes, start=1)
    ]


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

    success, encoded = cv2.imencode(".png", annotated)
    if not success:
        raise DisplayComplianceSegmentationError("Could not render annotated preview.")
    return BytesIO(encoded.tobytes()).getvalue()


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
    return cv2.dilate(closed, kernel, iterations=1)


def _contour_bboxes(mask: np.ndarray) -> list[BBox]:
    contours, _hierarchy = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    return [cv2.boundingRect(contour) for contour in contours]


def _filter_bboxes(
    bboxes: list[BBox],
    *,
    image_width: int,
    image_height: int,
    config: RegionDetectionConfig,
) -> list[BBox]:
    image_area = image_width * image_height
    filtered: list[BBox] = []
    for x, y, width, height in bboxes:
        if width <= 0 or height <= 0:
            continue
        if width <= config.min_side_pixels or height <= config.min_side_pixels:
            continue
        area_ratio = (width * height) / image_area
        if area_ratio < config.min_relative_area or area_ratio > config.max_relative_area:
            continue
        aspect_ratio = max(width / height, height / width)
        if aspect_ratio > config.max_aspect_ratio:
            continue
        if width >= image_width * 0.97 and height >= image_height * 0.97:
            continue
        filtered.append((x, y, width, height))
    return filtered


def _deduplicate_bboxes(
    bboxes: list[BBox],
    *,
    config: RegionDetectionConfig,
) -> list[BBox]:
    ordered = sorted(bboxes, key=lambda bbox: bbox[2] * bbox[3], reverse=True)
    kept: list[BBox] = []
    for candidate in ordered:
        if any(
            _bbox_iou(candidate, existing) >= config.overlap_iou_threshold
            or _smaller_coverage(candidate, existing) >= config.overlap_coverage_threshold
            for existing in kept
        ):
            continue
        kept.append(candidate)
    return kept


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
