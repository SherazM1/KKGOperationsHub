"""Tests for learned Display Compliance proposal backends."""

from __future__ import annotations

from dataclasses import asdict

import cv2
import numpy as np
import pytest

from app.display_compliance.proposal_backends import (
    RegionProposalBackend,
    RegionProposalBackendUnavailable,
    Sam2AutomaticMaskBackend,
)
from app.display_compliance.proposal_backends.sam2_backend import (
    DEFAULT_SAM2_CACHE_DIR_ENV,
    DEFAULT_SAM2_CHECKPOINT_ENV,
    DEFAULT_LEARNED_MAX_SIDE_ENV,
    DEFAULT_SAM2_MODEL_CONFIG_ENV,
    _CHECKPOINT_PATH_CACHE,
)
from app.display_compliance.segmentation import detect_candidate_regions


class _MockMaskGenerator:
    def __init__(self, masks: list[dict[str, object]]) -> None:
        self.masks = masks
        self.calls = 0

    def generate(self, image: np.ndarray) -> list[dict[str, object]]:
        self.calls += 1
        assert image.ndim == 3
        return self.masks


def _encode_png(image: np.ndarray) -> bytes:
    success, encoded = cv2.imencode(".png", image)
    assert success
    return encoded.tobytes()


def _blank_image(width: int = 240, height: int = 180) -> np.ndarray:
    return np.full((height, width, 3), 255, dtype=np.uint8)


def _rect_mask(
    *,
    image_width: int = 240,
    image_height: int = 180,
    left: int,
    top: int,
    width: int,
    height: int,
    predicted_iou: float = 0.9,
    stability_score: float = 0.92,
) -> dict[str, object]:
    mask = np.zeros((image_height, image_width), dtype=bool)
    mask[top : top + height, left : left + width] = True
    return {
        "segmentation": mask,
        "bbox": [left, top, width, height],
        "area": int(mask.sum()),
        "predicted_iou": predicted_iou,
        "stability_score": stability_score,
    }


def _proposal_result(masks: list[dict[str, object]]) -> object:
    backend = Sam2AutomaticMaskBackend(mask_generator=_MockMaskGenerator(masks))
    return backend.propose(image_bytes=_encode_png(_blank_image()))


def test_generic_learned_backend_interface() -> None:
    backend: RegionProposalBackend = Sam2AutomaticMaskBackend(mask_generator=_MockMaskGenerator([]))

    result = backend.propose(image_bytes=_encode_png(_blank_image()))

    assert result.diagnostics.backend == "sam2"
    assert result.regions == []


def test_unavailable_checkpoint_fails_cleanly(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv(DEFAULT_SAM2_CHECKPOINT_ENV, raising=False)
    monkeypatch.delenv(DEFAULT_SAM2_MODEL_CONFIG_ENV, raising=False)
    monkeypatch.delenv(DEFAULT_SAM2_CACHE_DIR_ENV, raising=False)

    def failing_download(url, destination, timeout_seconds) -> None:
        raise RegionProposalBackendUnavailable("download failed")

    backend = Sam2AutomaticMaskBackend(checkpoint_downloader=failing_download)

    with pytest.raises(RegionProposalBackendUnavailable, match="download failed"):
        backend.propose(image_bytes=_encode_png(_blank_image()))


def test_unavailable_model_dependency_fails_cleanly(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path,
) -> None:
    checkpoint = tmp_path / "sam2.pt"
    checkpoint.write_bytes(b"not a real checkpoint")
    monkeypatch.setenv(DEFAULT_SAM2_CHECKPOINT_ENV, str(checkpoint))
    monkeypatch.setenv(DEFAULT_SAM2_MODEL_CONFIG_ENV, "sam2_hiera_t.yaml")

    backend = Sam2AutomaticMaskBackend()

    with pytest.raises(RegionProposalBackendUnavailable, match="Learned region proposal backend"):
        backend.propose(image_bytes=_encode_png(_blank_image()))


def test_mock_sam_masks_convert_to_product_regions() -> None:
    result = _proposal_result([
        _rect_mask(left=24, top=20, width=44, height=58),
        _rect_mask(left=90, top=20, width=45, height=58),
    ])

    assert len(result.regions) == 2
    assert result.diagnostics.raw_mask_count == 2
    assert all(region.bbox is not None for region in result.regions)
    assert all(region.polygon for region in result.regions)


def test_explicit_checkpoint_override_is_preserved(tmp_path) -> None:
    checkpoint = tmp_path / "override.pt"
    checkpoint.write_bytes(b"checkpoint")

    def unexpected_download(url, destination, timeout_seconds) -> None:
        raise AssertionError("explicit checkpoint should not download")

    backend = Sam2AutomaticMaskBackend(
        checkpoint_path=str(checkpoint),
        checkpoint_downloader=unexpected_download,
    )

    assert backend._resolve_checkpoint() == checkpoint


def test_missing_checkpoint_uses_automatic_cache_resolution(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path,
) -> None:
    monkeypatch.setattr(
        "app.display_compliance.proposal_backends.sam2_backend.MIN_CHECKPOINT_BYTES",
        1,
    )
    _CHECKPOINT_PATH_CACHE.clear()

    def fake_download(url, destination, timeout_seconds) -> None:
        destination.write_bytes(b"cached checkpoint")

    backend = Sam2AutomaticMaskBackend(
        cache_dir=tmp_path / "cache",
        checkpoint_url="https://example.test/sam2.pt",
        checkpoint_downloader=fake_download,
    )

    checkpoint = backend._resolve_checkpoint()

    assert checkpoint.exists()
    assert checkpoint.parent == tmp_path / "cache"
    assert checkpoint.name == "sam2.1_hiera_tiny.pt"


def test_existing_cached_checkpoint_is_reused(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path,
) -> None:
    monkeypatch.setattr(
        "app.display_compliance.proposal_backends.sam2_backend.MIN_CHECKPOINT_BYTES",
        1,
    )
    _CHECKPOINT_PATH_CACHE.clear()
    cache_dir = tmp_path / "cache"
    cache_dir.mkdir()
    checkpoint = cache_dir / "sam2.1_hiera_tiny.pt"
    checkpoint.write_bytes(b"cached")
    downloads = 0

    def fake_download(url, destination, timeout_seconds) -> None:
        nonlocal downloads
        downloads += 1
        destination.write_bytes(b"new")

    backend = Sam2AutomaticMaskBackend(
        cache_dir=cache_dir,
        checkpoint_url="https://example.test/reuse.pt",
        checkpoint_downloader=fake_download,
    )

    assert backend._resolve_checkpoint() == checkpoint
    assert backend._resolve_checkpoint() == checkpoint
    assert downloads == 0


def test_partial_download_is_not_treated_as_valid(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path,
) -> None:
    monkeypatch.setattr(
        "app.display_compliance.proposal_backends.sam2_backend.MIN_CHECKPOINT_BYTES",
        100,
    )
    _CHECKPOINT_PATH_CACHE.clear()

    def partial_download(url, destination, timeout_seconds) -> None:
        destination.write_bytes(b"partial")

    backend = Sam2AutomaticMaskBackend(
        cache_dir=tmp_path,
        checkpoint_url="https://example.test/partial.pt",
        checkpoint_downloader=partial_download,
    )

    with pytest.raises(RegionProposalBackendUnavailable, match="incomplete"):
        backend._resolve_checkpoint()
    assert not (tmp_path / "sam2.1_hiera_tiny.pt").exists()
    assert not (tmp_path / "sam2.1_hiera_tiny.pt.download").exists()


def test_download_failure_produces_backend_unavailable_error(tmp_path) -> None:
    _CHECKPOINT_PATH_CACHE.clear()

    def failed_download(url, destination, timeout_seconds) -> None:
        raise RegionProposalBackendUnavailable("network failed")

    backend = Sam2AutomaticMaskBackend(
        cache_dir=tmp_path / "nested" / "cache",
        checkpoint_url="https://example.test/fail.pt",
        checkpoint_downloader=failed_download,
    )

    with pytest.raises(RegionProposalBackendUnavailable, match="network failed"):
        backend._resolve_checkpoint()


def test_learned_max_side_env_is_used(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(DEFAULT_LEARNED_MAX_SIDE_ENV, "320")

    backend = Sam2AutomaticMaskBackend(mask_generator=_MockMaskGenerator([]))

    assert backend.max_working_side == 320


def test_polygons_and_bboxes_remain_in_bounds() -> None:
    result = _proposal_result([_rect_mask(left=210, top=150, width=80, height=60)])

    assert result.regions
    for region in result.regions:
        assert region.bbox is not None
        x, y, width, height = region.bbox
        assert 0 <= x < 240
        assert 0 <= y < 180
        assert x + width <= 240
        assert y + height <= 180
        assert all(0 <= px <= 240 and 0 <= py <= 180 for px, py in region.polygon)


def test_inference_resolution_scaling_maps_regions_to_original_dimensions() -> None:
    class SizeAwareMaskGenerator:
        def generate(self, image: np.ndarray) -> list[dict[str, object]]:
            assert image.shape[:2] == (180, 240)
            return [
                _rect_mask(
                    image_width=240,
                    image_height=180,
                    left=24,
                    top=20,
                    width=44,
                    height=58,
                )
            ]

    backend = Sam2AutomaticMaskBackend(
        mask_generator=SizeAwareMaskGenerator(),
        max_working_side=240,
    )
    result = backend.propose(image_bytes=_encode_png(_blank_image(width=480, height=360)))

    assert result.diagnostics.original_width == 480
    assert result.diagnostics.working_width == 240
    assert result.regions[0].bbox == (48, 40, 88, 116)


def test_tiny_masks_are_filtered() -> None:
    result = _proposal_result([_rect_mask(left=10, top=10, width=3, height=3)])

    assert result.regions == []
    assert result.diagnostics.rejected_too_small == 1


def test_near_whole_image_masks_are_filtered() -> None:
    result = _proposal_result([_rect_mask(left=1, top=1, width=238, height=178)])

    assert result.regions == []
    assert result.diagnostics.rejected_too_large == 1


def test_duplicate_and_nested_masks_are_handled_deterministically() -> None:
    masks = [
        _rect_mask(left=30, top=30, width=60, height=70, stability_score=0.95),
        _rect_mask(left=32, top=32, width=57, height=67, stability_score=0.9),
        _rect_mask(left=48, top=48, width=18, height=18, stability_score=0.98),
    ]

    first = _proposal_result(masks)
    second = _proposal_result(masks)

    assert first.regions == second.regions
    assert first.diagnostics.removed_by_deduplication >= 1
    assert first.diagnostics.final_region_count == 1


def test_adjacent_masks_remain_distinct() -> None:
    result = _proposal_result([
        _rect_mask(left=24, top=30, width=42, height=58),
        _rect_mask(left=70, top=30, width=42, height=58),
        _rect_mask(left=116, top=30, width=42, height=58),
    ])

    assert result.diagnostics.final_region_count == 3


def test_repeated_size_alignment_scoring_boosts_generic_regions() -> None:
    result = _proposal_result([
        _rect_mask(left=22, top=24, width=40, height=54),
        _rect_mask(left=72, top=24, width=40, height=54),
        _rect_mask(left=122, top=24, width=40, height=54),
        _rect_mask(left=22, top=92, width=40, height=54),
        _rect_mask(left=72, top=92, width=40, height=54),
        _rect_mask(left=122, top=92, width=40, height=54),
    ])

    assert result.diagnostics.repeated_size_member_count == 6
    assert result.diagnostics.alignment_supported_count == 6
    assert result.diagnostics.final_region_count == 6


def test_final_product_region_ids_are_deterministic() -> None:
    masks = [
        _rect_mask(left=120, top=90, width=42, height=58),
        _rect_mask(left=24, top=24, width=42, height=58),
        _rect_mask(left=120, top=24, width=42, height=58),
    ]

    result = _proposal_result(masks)

    assert [region.region_id for region in result.regions] == [
        "region_001",
        "region_002",
        "region_003",
    ]
    assert [region.bbox for region in result.regions] == [
        (24, 24, 42, 58),
        (120, 24, 42, 58),
        (120, 90, 42, 58),
    ]


def test_existing_classical_detector_remains_functional() -> None:
    image = _blank_image()
    cv2.rectangle(image, (30, 30), (100, 95), (210, 210, 210), -1)
    cv2.rectangle(image, (30, 30), (100, 95), (30, 30, 30), 3)

    regions = detect_candidate_regions(image_bytes=_encode_png(image))

    assert isinstance(regions, list)


def test_learned_diagnostics_serialize_and_images_decode() -> None:
    result = _proposal_result([_rect_mask(left=24, top=20, width=44, height=58)])
    diagnostics = asdict(result.diagnostics)

    assert diagnostics["backend"] == "sam2"
    assert set(result.diagnostic_images) == {
        "learned_raw_masks",
        "learned_raw_bboxes",
        "learned_after_basic_filtering",
        "learned_after_duplicate_cleanup",
        "learned_structurally_supported",
        "learned_final_regions",
    }
    for image_bytes in result.diagnostic_images.values():
        assert image_bytes.startswith(b"\x89PNG\r\n\x1a\n")
        decoded = cv2.imdecode(np.frombuffer(image_bytes, dtype=np.uint8), cv2.IMREAD_COLOR)
        assert decoded is not None
