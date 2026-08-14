"""Tests for Display Compliance candidate-region detection."""

from __future__ import annotations

import cv2
import numpy as np

from app.display_compliance.segmentation import (
    RegionDetectionConfig,
    _deduplicate_bboxes,
    analyze_candidate_regions,
    detect_candidate_regions,
    render_annotated_preview,
)


def _encode_png(image: np.ndarray) -> bytes:
    success, encoded = cv2.imencode(".png", image)
    assert success
    return encoded.tobytes()


def _decode_png(image_bytes: bytes) -> np.ndarray:
    decoded = cv2.imdecode(np.frombuffer(image_bytes, dtype=np.uint8), cv2.IMREAD_COLOR)
    assert decoded is not None
    return decoded


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


def _draw_product_grid(
    image: np.ndarray,
    *,
    columns: int = 4,
    rows: int = 3,
    left: int = 28,
    top: int = 24,
    box_width: int = 46,
    box_height: int = 58,
    gap_x: int = 16,
    gap_y: int = 14,
) -> None:
    for row in range(rows):
        for column in range(columns):
            x = left + column * (box_width + gap_x)
            y = top + row * (box_height + gap_y)
            _draw_product_box(
                image,
                left=x,
                top=y,
                right=x + box_width,
                bottom=y + box_height,
            )


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


def test_diagnostics_return_dimensions_and_non_negative_counts() -> None:
    image_bytes = _encode_png(_blank_image(width=420, height=260))

    result = analyze_candidate_regions(image_bytes=image_bytes)
    diagnostics = result.diagnostics

    assert diagnostics.original_width == 420
    assert diagnostics.original_height == 260
    assert diagnostics.working_width == 420
    assert diagnostics.working_height == 260
    for value in diagnostics.__dict__.values():
        assert value >= 0


def test_diagnostic_counts_reconcile() -> None:
    image = _blank_image()
    _draw_product_box(image, left=30, top=30, right=100, bottom=95)
    result = analyze_candidate_regions(image_bytes=_encode_png(image))
    diagnostics = result.diagnostics

    geometry_rejections = (
        diagnostics.rejected_degenerate
        + diagnostics.rejected_too_small
        + diagnostics.rejected_too_large
        + diagnostics.rejected_aspect_ratio
        + diagnostics.rejected_near_whole_image
    )
    assert diagnostics.raw_proposal_count - geometry_rejections == (
        diagnostics.proposals_after_geometry_filter
    )
    assert (
        diagnostics.strategy_a_raw_proposal_count - geometry_rejections
        == diagnostics.strategy_a_proposals_after_geometry_filter
    )
    assert diagnostics.raw_proposal_count == diagnostics.strategy_a_raw_proposal_count
    assert (
        diagnostics.proposals_after_geometry_filter
        == diagnostics.strategy_a_proposals_after_geometry_filter
    )
    assert diagnostics.deduplication_input_count == diagnostics.merged_pool_count_before_dedup
    assert (
        diagnostics.removed_by_iou_deduplication
        + diagnostics.removed_by_coverage_deduplication
        == diagnostics.removed_by_deduplication
    )
    assert diagnostics.deduplication_input_count - diagnostics.removed_by_deduplication == (
        diagnostics.final_region_count
    )


def test_multi_strategy_diagnostics_expose_strategy_and_merge_counts() -> None:
    image = _blank_image(width=340, height=260)
    _draw_product_grid(image, columns=3, rows=2)

    result = analyze_candidate_regions(image_bytes=_encode_png(image))
    diagnostics = result.diagnostics

    assert diagnostics.strategy_a_raw_proposal_count >= 0
    assert diagnostics.strategy_b_raw_proposal_count > 0
    assert diagnostics.strategy_c_raw_proposal_count > 0
    assert diagnostics.merged_pool_count_before_dedup >= diagnostics.final_region_count
    assert diagnostics.deduplication_input_count == diagnostics.merged_pool_count_before_dedup
    assert diagnostics.removed_by_deduplication >= 0


def test_diagnostic_image_bytes_decode_as_pngs() -> None:
    result = analyze_candidate_regions(image_bytes=_encode_png(_blank_image()))

    assert set(result.diagnostic_images) == {
        "normalized",
        "edges",
        "threshold",
        "morphology",
        "strategy_a_proposals",
        "strategy_b_proposals",
        "strategy_c_proposals",
        "merged_proposals",
        "raw_proposals",
        "filtered_proposals",
        "final_proposals",
    }
    for image_bytes in result.diagnostic_images.values():
        assert image_bytes.startswith(b"\x89PNG\r\n\x1a\n")
        assert _decode_png(image_bytes).size > 0


def test_synthetic_rectangles_produce_raw_diagnostic_proposals() -> None:
    image = _blank_image()
    _draw_product_box(image, left=30, top=30, right=100, bottom=95)
    _draw_product_box(image, left=145, top=35, right=220, bottom=100)

    result = analyze_candidate_regions(image_bytes=_encode_png(image))

    assert result.diagnostics.raw_contour_count > 0
    assert result.diagnostics.raw_proposal_count > 0
    assert result.proposal_sample


def test_repeated_rectangle_grid_produces_multiple_plausible_proposals() -> None:
    image = _blank_image(width=340, height=280)
    _draw_product_grid(image, columns=4, rows=3)

    result = analyze_candidate_regions(image_bytes=_encode_png(image))

    assert len(result.regions) > 1
    assert result.diagnostics.repeated_size_member_count > 1
    assert result.diagnostics.size_cluster_count > 0


def test_alignment_and_repeated_size_scoring_retains_grid_proposals() -> None:
    image = _blank_image(width=420, height=320)
    _draw_product_grid(
        image,
        columns=5,
        rows=3,
        left=24,
        top=28,
        box_width=48,
        box_height=62,
        gap_x=18,
        gap_y=18,
    )

    result = analyze_candidate_regions(image_bytes=_encode_png(image))

    assert result.diagnostics.alignment_boosted_count > 1
    assert result.diagnostics.final_region_count >= 6


def test_zero_region_images_still_produce_diagnostic_stage_images() -> None:
    result = analyze_candidate_regions(image_bytes=_encode_png(_blank_image()))

    assert result.regions == []
    assert result.diagnostics.final_region_count == 0
    assert result.diagnostics.merged_pool_count_before_dedup == 0
    assert all(result.diagnostic_images.values())


def test_multi_strategy_merge_does_not_fabricate_blank_image_proposals() -> None:
    result = analyze_candidate_regions(image_bytes=_encode_png(_blank_image(width=420, height=260)))

    assert result.regions == []
    assert result.diagnostics.strategy_a_raw_proposal_count == 0
    assert result.diagnostics.strategy_b_raw_proposal_count == 0
    assert result.diagnostics.strategy_c_raw_proposal_count == 0
    assert result.diagnostics.merged_pool_count_before_dedup == 0


def test_diagnostics_do_not_change_deterministic_final_region_ids() -> None:
    image = _blank_image(width=360, height=260)
    _draw_product_box(image, left=205, top=145, right=285, bottom=210)
    _draw_product_box(image, left=35, top=35, right=115, bottom=100)
    _draw_product_box(image, left=205, top=35, right=285, bottom=100)
    _draw_product_box(image, left=35, top=145, right=115, bottom=210)
    image_bytes = _encode_png(image)

    regions = detect_candidate_regions(image_bytes=image_bytes)
    diagnostic_regions = analyze_candidate_regions(image_bytes=image_bytes).regions

    assert diagnostic_regions == regions
    assert [region.region_id for region in diagnostic_regions] == [
        f"region_{index:03}" for index in range(1, len(diagnostic_regions) + 1)
    ]


def test_repeated_diagnostic_runs_return_equivalent_counts() -> None:
    image = _blank_image()
    _draw_product_box(image, left=35, top=35, right=100, bottom=95)
    image_bytes = _encode_png(image)

    first = analyze_candidate_regions(image_bytes=image_bytes)
    second = analyze_candidate_regions(image_bytes=image_bytes)

    assert second.diagnostics == first.diagnostics
    assert second.regions == first.regions


def test_detect_candidate_regions_remains_backward_compatible() -> None:
    image = _blank_image()
    _draw_product_box(image, left=30, top=30, right=100, bottom=95)
    image_bytes = _encode_png(image)

    regions = detect_candidate_regions(image_bytes=image_bytes)

    assert isinstance(regions, list)
    assert regions == analyze_candidate_regions(image_bytes=image_bytes).regions
