"""Tests for Display Compliance candidate-region detection."""

from __future__ import annotations

import cv2
import numpy as np

from app.display_compliance.segmentation import (
    RegionDetectionConfig,
    _deduplicate_bboxes,
    detect_candidate_regions,
    render_annotated_preview,
)


def _encode_png(image: np.ndarray) -> bytes:
    success, encoded = cv2.imencode(".png", image)
    assert success
    return encoded.tobytes()


def _blank_image(width: int = 320, height: int = 220) -> np.ndarray:
    return np.full((height, width, 3), 255, dtype=np.uint8)


def _draw_product_box(
    image: np.ndarray,
    *,
    left: int,
    top: int,
    right: int,
    bottom: int,
) -> None:
    cv2.rectangle(image, (left, top), (right, bottom), (210, 210, 210), -1)
    cv2.rectangle(image, (left, top), (right, bottom), (30, 30, 30), 3)


def test_detector_accepts_valid_image_bytes() -> None:
    regions = detect_candidate_regions(image_bytes=_encode_png(_blank_image()))

    assert isinstance(regions, list)


def test_obvious_rectangular_objects_produce_candidate_regions() -> None:
    image = _blank_image()
    _draw_product_box(image, left=30, top=30, right=100, bottom=95)
    _draw_product_box(image, left=145, top=35, right=220, bottom=100)

    regions = detect_candidate_regions(image_bytes=_encode_png(image))

    assert len(regions) >= 2
    assert all(region.bbox is not None for region in regions)


def test_tiny_noise_is_filtered() -> None:
    image = _blank_image()
    cv2.rectangle(image, (10, 10), (13, 13), (0, 0, 0), -1)

    regions = detect_candidate_regions(image_bytes=_encode_png(image))

    assert regions == []


def test_very_large_whole_image_contours_are_filtered() -> None:
    image = _blank_image()
    cv2.rectangle(image, (1, 1), (318, 218), (0, 0, 0), 3)

    regions = detect_candidate_regions(image_bytes=_encode_png(image))

    assert regions == []


def test_duplicate_highly_overlapping_proposals_are_deduplicated() -> None:
    bboxes = [(20, 20, 60, 50), (22, 21, 58, 49), (160, 20, 50, 45)]

    deduped = _deduplicate_bboxes(bboxes, config=RegionDetectionConfig())

    assert len(deduped) == 2


def test_region_ids_are_deterministic_and_spatially_ordered() -> None:
    image = _blank_image(width=360, height=260)
    _draw_product_box(image, left=205, top=145, right=285, bottom=210)
    _draw_product_box(image, left=35, top=35, right=115, bottom=100)
    _draw_product_box(image, left=205, top=35, right=285, bottom=100)
    _draw_product_box(image, left=35, top=145, right=115, bottom=210)

    regions = detect_candidate_regions(image_bytes=_encode_png(image))

    bboxes = [region.bbox for region in regions]
    assert [region.region_id for region in regions] == [
        f"region_{index:03}" for index in range(1, len(regions) + 1)
    ]
    assert bboxes[:4] == sorted(bboxes[:4], key=lambda bbox: (bbox[1] // 80, bbox[0]))


def test_polygons_and_bboxes_remain_within_original_image_dimensions() -> None:
    width = 1800
    height = 1200
    image = _blank_image(width=width, height=height)
    _draw_product_box(image, left=120, top=100, right=520, bottom=420)
    _draw_product_box(image, left=760, top=520, right=1160, bottom=840)

    regions = detect_candidate_regions(image_bytes=_encode_png(image))

    assert regions
    for region in regions:
        assert region.bbox is not None
        x, y, box_width, box_height = region.bbox
        assert 0 <= x < width
        assert 0 <= y < height
        assert x + box_width <= width
        assert y + box_height <= height
        for point_x, point_y in region.polygon:
            assert 0 <= point_x <= width
            assert 0 <= point_y <= height


def test_rerunning_detection_returns_equivalent_geometry_and_order() -> None:
    image = _blank_image()
    _draw_product_box(image, left=35, top=35, right=100, bottom=95)
    _draw_product_box(image, left=140, top=35, right=205, bottom=95)
    image_bytes = _encode_png(image)

    first = detect_candidate_regions(image_bytes=image_bytes)
    second = detect_candidate_regions(image_bytes=image_bytes)

    assert second == first


def test_zero_region_result_does_not_fabricate_detections() -> None:
    regions = detect_candidate_regions(image_bytes=_encode_png(_blank_image()))

    assert regions == []


def test_annotated_preview_renders_png_bytes() -> None:
    image = _blank_image()
    _draw_product_box(image, left=30, top=30, right=100, bottom=95)
    image_bytes = _encode_png(image)
    regions = detect_candidate_regions(image_bytes=image_bytes)

    preview = render_annotated_preview(image_bytes=image_bytes, regions=regions)

    assert preview.startswith(b"\x89PNG\r\n\x1a\n")
