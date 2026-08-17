"""SAM 2 automatic-mask learned proposal backend.

The dependency is intentionally optional. Importing this module does not import
PyTorch or SAM 2; model loading happens only when the backend is used.
"""

from __future__ import annotations

from dataclasses import replace
from io import BytesIO
import math
import os
from pathlib import Path
import tempfile
import time
from typing import Any
from typing import Callable
from urllib.error import URLError
from urllib.request import Request
from urllib.request import urlopen

import cv2
import numpy as np

from app.display_compliance.models import ProductRegion
from app.display_compliance.proposal_backends.base import (
    BBox,
    LearnedProposal,
    LearnedSegmentationDiagnostics,
    RegionProposalBackendUnavailable,
    RegionProposalResult,
)
from app.display_compliance.segmentation import (
    DisplayComplianceSegmentationError,
    _bbox_iou,
    _clamp_bbox,
    _group_similar_centers,
    _group_similar_sizes,
    _index_cluster_sizes,
    _map_bbox_to_original,
    _smaller_coverage,
    _spatially_sort_bboxes,
)

DEFAULT_SAM2_CHECKPOINT_ENV = "DISPLAY_COMPLIANCE_SAM2_CHECKPOINT"
DEFAULT_SAM2_CACHE_DIR_ENV = "DISPLAY_COMPLIANCE_SAM2_CACHE_DIR"
DEFAULT_SAM2_MODEL_CONFIG_ENV = "DISPLAY_COMPLIANCE_SAM2_MODEL_CONFIG"
DEFAULT_LEARNED_MAX_SIDE_ENV = "DISPLAY_COMPLIANCE_LEARNED_MAX_SIDE"
DEFAULT_SAM2_MODEL_CONFIG = "configs/sam2.1/sam2.1_hiera_t.yaml"
DEFAULT_SAM2_TINY_CHECKPOINT_FILENAME = "sam2.1_hiera_tiny.pt"
DEFAULT_SAM2_TINY_CHECKPOINT_URL = (
    "https://dl.fbaipublicfiles.com/segment_anything_2/092824/"
    + DEFAULT_SAM2_TINY_CHECKPOINT_FILENAME
)
DEFAULT_MAX_WORKING_SIDE = 640
DEFAULT_DOWNLOAD_TIMEOUT_SECONDS = 120
MIN_CHECKPOINT_BYTES = 10 * 1024 * 1024

_MODEL_CACHE: dict[tuple[str, str, str], tuple[Any, float]] = {}
_CHECKPOINT_PATH_CACHE: dict[tuple[str, str], Path] = {}


class Sam2AutomaticMaskBackend:
    """Product-agnostic SAM 2 automatic-mask proposal backend."""

    backend_name = "sam2"

    def __init__(
        self,
        *,
        mask_generator: Any | None = None,
        checkpoint_path: str | None = None,
        cache_dir: str | Path | None = None,
        checkpoint_url: str = DEFAULT_SAM2_TINY_CHECKPOINT_URL,
        checkpoint_downloader: Callable[[str, Path, int], None] | None = None,
        model_config: str | None = None,
        device: str | None = None,
        max_working_side: int | None = None,
        download_timeout_seconds: int = DEFAULT_DOWNLOAD_TIMEOUT_SECONDS,
        min_area_fraction: float = 0.00025,
        max_area_fraction: float = 0.78,
        min_side_pixels: int = 6,
        max_aspect_ratio: float = 12.0,
        min_solidity: float = 0.08,
        min_predicted_iou: float = 0.0,
        min_stability_score: float = 0.0,
    ) -> None:
        self._provided_mask_generator = mask_generator
        self.checkpoint_path = checkpoint_path or os.getenv(DEFAULT_SAM2_CHECKPOINT_ENV, "")
        self.cache_dir = Path(
            cache_dir
            or os.getenv(DEFAULT_SAM2_CACHE_DIR_ENV, "")
            or Path(tempfile.gettempdir()) / "display_compliance_models"
        )
        self.checkpoint_url = checkpoint_url
        self._checkpoint_downloader = checkpoint_downloader
        self.model_config = (
            model_config
            or os.getenv(DEFAULT_SAM2_MODEL_CONFIG_ENV, "")
            or DEFAULT_SAM2_MODEL_CONFIG
        )
        self.device = device
        self.max_working_side = max_working_side or _env_int(
            DEFAULT_LEARNED_MAX_SIDE_ENV,
            DEFAULT_MAX_WORKING_SIDE,
        )
        self.download_timeout_seconds = download_timeout_seconds
        self.min_area_fraction = min_area_fraction
        self.max_area_fraction = max_area_fraction
        self.min_side_pixels = min_side_pixels
        self.max_aspect_ratio = max_aspect_ratio
        self.min_solidity = min_solidity
        self.min_predicted_iou = min_predicted_iou
        self.min_stability_score = min_stability_score

    def propose(self, *, image_bytes: bytes) -> RegionProposalResult:
        started_at = time.perf_counter()
        image = _decode_image(image_bytes)
        original_height, original_width = image.shape[:2]
        working_image, scale = _resize_for_detection(image, self.max_working_side)
        mask_generator, load_seconds, device, checkpoint = self._mask_generator()

        inference_started_at = time.perf_counter()
        masks = mask_generator.generate(cv2.cvtColor(working_image, cv2.COLOR_BGR2RGB))
        inference_seconds = time.perf_counter() - inference_started_at

        raw_proposals, filter_counts = _proposals_from_sam_masks(
            masks,
            image_width=working_image.shape[1],
            image_height=working_image.shape[0],
            backend_name=self.backend_name,
            min_area_fraction=self.min_area_fraction,
            max_area_fraction=self.max_area_fraction,
            min_side_pixels=self.min_side_pixels,
            max_aspect_ratio=self.max_aspect_ratio,
            min_solidity=self.min_solidity,
            min_predicted_iou=self.min_predicted_iou,
            min_stability_score=self.min_stability_score,
        )
        scored, scoring_summary = _score_learned_proposals(raw_proposals)
        deduped, dedup_counts = _remove_duplicate_and_nested_masks(scored)
        supported = _structurally_supported_proposals(deduped)

        original_proposals = [
            _proposal_to_original(
                proposal,
                scale=scale,
                original_width=original_width,
                original_height=original_height,
            )
            for proposal in supported
        ]
        ordered_bboxes = _spatially_sort_bboxes([proposal.bbox for proposal in original_proposals])
        proposal_by_bbox = {proposal.bbox: proposal for proposal in original_proposals}
        regions = [
            ProductRegion(
                region_id=f"region_{index:03}",
                bbox=bbox,
                polygon=proposal_by_bbox[bbox].polygon or _bbox_polygon(bbox),
            )
            for index, bbox in enumerate(ordered_bboxes, start=1)
        ]

        total_seconds = time.perf_counter() - started_at
        removed_by_deduplication = (
            dedup_counts["removed_by_iou"] + dedup_counts["removed_by_nested"]
        )
        diagnostics = LearnedSegmentationDiagnostics(
            backend=self.backend_name,
            original_width=original_width,
            original_height=original_height,
            working_width=working_image.shape[1],
            working_height=working_image.shape[0],
            device=device,
            model_config=self.model_config or "(provided generator)",
            checkpoint=str(checkpoint) if checkpoint is not None else "(provided generator)",
            model_load_seconds=round(load_seconds, 4),
            inference_seconds=round(inference_seconds, 4),
            total_seconds=round(total_seconds, 4),
            raw_mask_count=len(masks),
            rejected_degenerate=filter_counts["rejected_degenerate"],
            rejected_too_small=filter_counts["rejected_too_small"],
            rejected_too_large=filter_counts["rejected_too_large"],
            rejected_aspect_ratio=filter_counts["rejected_aspect_ratio"],
            rejected_low_confidence=filter_counts["rejected_low_confidence"],
            rejected_low_solidity=filter_counts["rejected_low_solidity"],
            proposals_after_basic_filtering=len(raw_proposals),
            removed_by_iou_deduplication=dedup_counts["removed_by_iou"],
            removed_by_nested_deduplication=dedup_counts["removed_by_nested"],
            removed_by_deduplication=removed_by_deduplication,
            proposals_after_duplicate_cleanup=len(deduped),
            size_cluster_count=scoring_summary["size_cluster_count"],
            repeated_size_member_count=scoring_summary["repeated_size_member_count"],
            alignment_supported_count=scoring_summary["alignment_supported_count"],
            final_region_count=len(regions),
        )
        return RegionProposalResult(
            regions=regions,
            diagnostics=diagnostics,
            diagnostic_images={
                "learned_raw_masks": _render_mask_overlay(working_image, raw_proposals),
                "learned_raw_bboxes": _render_bbox_overlay(
                    working_image,
                    [proposal.bbox for proposal in raw_proposals],
                ),
                "learned_after_basic_filtering": _render_mask_overlay(
                    working_image,
                    raw_proposals,
                ),
                "learned_after_duplicate_cleanup": _render_mask_overlay(
                    working_image,
                    deduped,
                ),
                "learned_structurally_supported": _render_bbox_overlay(
                    working_image,
                    [proposal.bbox for proposal in supported],
                ),
                "learned_final_regions": _render_bbox_overlay(
                    image,
                    [region.bbox for region in regions if region.bbox is not None],
                ),
            },
            proposals=original_proposals,
        )

    def _mask_generator(self) -> tuple[Any, float, str, Path | None]:
        if self._provided_mask_generator is not None:
            return self._provided_mask_generator, 0.0, self.device or "provided", None

        if not self.model_config:
            raise RegionProposalBackendUnavailable(
                "Learned region proposal backend is unavailable. "
                f"Set {DEFAULT_SAM2_MODEL_CONFIG_ENV} to a SAM 2 model config."
            )
        checkpoint = self._resolve_checkpoint()

        try:
            import torch
            from sam2.automatic_mask_generator import SAM2AutomaticMaskGenerator
            from sam2.build_sam import build_sam2
        except Exception as exc:
            raise RegionProposalBackendUnavailable(
                "Learned region proposal backend is unavailable. "
                "Install PyTorch and SAM 2 to use learned segmentation."
            ) from exc

        device = self.device
        if device is None:
            device = "cuda" if torch.cuda.is_available() else "cpu"
        cache_key = (str(checkpoint), self.model_config, device)
        cached = _MODEL_CACHE.get(cache_key)
        if cached is not None:
            return cached[0], 0.0, device, checkpoint

        load_started_at = time.perf_counter()
        try:
            model = build_sam2(self.model_config, str(checkpoint), device=device)
            generator = SAM2AutomaticMaskGenerator(model)
        except Exception as exc:
            raise RegionProposalBackendUnavailable(
                "Learned region proposal backend is unavailable. SAM 2 failed to load."
            ) from exc
        load_seconds = time.perf_counter() - load_started_at
        _MODEL_CACHE[cache_key] = (generator, load_seconds)
        return generator, load_seconds, device, checkpoint

    def _resolve_checkpoint(self) -> Path:
        if self.checkpoint_path:
            checkpoint = Path(self.checkpoint_path).expanduser()
            if not checkpoint.exists():
                raise RegionProposalBackendUnavailable(
                    "Learned region proposal backend is unavailable. "
                    f"SAM 2 checkpoint was not found: {checkpoint}"
                )
            if checkpoint.stat().st_size <= 0:
                raise RegionProposalBackendUnavailable(
                    "Learned region proposal backend is unavailable. "
                    f"SAM 2 checkpoint is empty: {checkpoint}"
                )
            return checkpoint

        cache_key = (str(self.cache_dir), self.checkpoint_url)
        cached_path = _CHECKPOINT_PATH_CACHE.get(cache_key)
        if cached_path is not None and _valid_cached_checkpoint(cached_path):
            return cached_path

        checkpoint = self.cache_dir / DEFAULT_SAM2_TINY_CHECKPOINT_FILENAME
        if _valid_cached_checkpoint(checkpoint):
            _CHECKPOINT_PATH_CACHE[cache_key] = checkpoint
            return checkpoint

        self.cache_dir.mkdir(parents=True, exist_ok=True)
        temporary_checkpoint = checkpoint.with_suffix(checkpoint.suffix + ".download")
        if temporary_checkpoint.exists():
            temporary_checkpoint.unlink()
        try:
            downloader = self._checkpoint_downloader or _download_checkpoint
            downloader(self.checkpoint_url, temporary_checkpoint, self.download_timeout_seconds)
            if not _valid_cached_checkpoint(temporary_checkpoint):
                raise RegionProposalBackendUnavailable(
                    "Learned region proposal backend is unavailable. "
                    "Downloaded SAM 2 checkpoint is incomplete."
                )
            os.replace(temporary_checkpoint, checkpoint)
        except RegionProposalBackendUnavailable:
            if temporary_checkpoint.exists():
                temporary_checkpoint.unlink()
            raise
        except OSError as exc:
            if temporary_checkpoint.exists():
                temporary_checkpoint.unlink()
            raise RegionProposalBackendUnavailable(
                "Learned region proposal backend is unavailable. "
                "Could not cache the SAM 2 checkpoint."
            ) from exc

        _CHECKPOINT_PATH_CACHE[cache_key] = checkpoint
        return checkpoint


def analyze_learned_candidate_regions(
    *,
    image_bytes: bytes,
    backend: Sam2AutomaticMaskBackend | None = None,
) -> RegionProposalResult:
    """Run the configured learned proposal backend."""
    return (backend or Sam2AutomaticMaskBackend()).propose(image_bytes=image_bytes)


def _env_int(name: str, default: int) -> int:
    raw_value = os.getenv(name)
    if not raw_value:
        return default
    try:
        value = int(raw_value)
    except ValueError:
        return default
    return value if value > 0 else default


def _valid_cached_checkpoint(path: Path) -> bool:
    return path.exists() and path.is_file() and path.stat().st_size >= MIN_CHECKPOINT_BYTES


def _download_checkpoint(url: str, destination: Path, timeout_seconds: int) -> None:
    request = Request(url, headers={"User-Agent": "KKGOperationsHub/DisplayCompliance"})
    try:
        with urlopen(request, timeout=timeout_seconds) as response:
            with destination.open("wb") as checkpoint_file:
                while True:
                    chunk = response.read(1024 * 1024)
                    if not chunk:
                        break
                    checkpoint_file.write(chunk)
    except (OSError, URLError) as exc:
        raise RegionProposalBackendUnavailable(
            "Learned region proposal backend is unavailable. "
            "Could not download the SAM 2.1 Tiny checkpoint."
        ) from exc


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


def _proposals_from_sam_masks(
    masks: list[dict[str, Any]],
    *,
    image_width: int,
    image_height: int,
    backend_name: str,
    min_area_fraction: float,
    max_area_fraction: float,
    min_side_pixels: int,
    max_aspect_ratio: float,
    min_solidity: float,
    min_predicted_iou: float,
    min_stability_score: float,
) -> tuple[list[LearnedProposal], dict[str, int]]:
    image_area = image_width * image_height
    proposals: list[LearnedProposal] = []
    counts = {
        "rejected_degenerate": 0,
        "rejected_too_small": 0,
        "rejected_too_large": 0,
        "rejected_aspect_ratio": 0,
        "rejected_low_confidence": 0,
        "rejected_low_solidity": 0,
    }
    for mask_payload in masks:
        mask = np.asarray(mask_payload.get("segmentation"), dtype=np.uint8)
        if mask.ndim != 2 or mask.shape[0] == 0 or mask.shape[1] == 0:
            counts["rejected_degenerate"] += 1
            continue
        bbox = _bbox_from_mask_payload(mask_payload, mask)
        bbox = _clamp_bbox(bbox, image_width=image_width, image_height=image_height)
        x, y, width, height = bbox
        area = int(mask_payload.get("area") or int(mask.sum()))
        if area <= 0 or width <= 0 or height <= 0:
            counts["rejected_degenerate"] += 1
            continue
        if width <= min_side_pixels or height <= min_side_pixels:
            counts["rejected_too_small"] += 1
            continue
        area_fraction = area / image_area
        if area_fraction < min_area_fraction:
            counts["rejected_too_small"] += 1
            continue
        if area_fraction > max_area_fraction or (
            width >= image_width * 0.97 and height >= image_height * 0.97
        ):
            counts["rejected_too_large"] += 1
            continue
        aspect_ratio = max(width / height, height / width)
        if aspect_ratio > max_aspect_ratio:
            counts["rejected_aspect_ratio"] += 1
            continue

        predicted_iou = _optional_float(mask_payload.get("predicted_iou"))
        stability_score = _optional_float(mask_payload.get("stability_score"))
        if predicted_iou is not None and predicted_iou < min_predicted_iou:
            counts["rejected_low_confidence"] += 1
            continue
        if stability_score is not None and stability_score < min_stability_score:
            counts["rejected_low_confidence"] += 1
            continue

        polygon = _polygon_from_mask(mask)
        solidity = min(1.0, area / max(1, width * height))
        if solidity < min_solidity:
            counts["rejected_low_solidity"] += 1
            continue
        proposals.append(
            LearnedProposal(
                bbox=bbox,
                polygon=polygon or _bbox_polygon(bbox),
                area=area,
                predicted_iou=predicted_iou,
                stability_score=stability_score,
                source_backend=backend_name,
                mask=mask.astype(bool),
                solidity=round(solidity, 4),
            )
        )
    return proposals, counts


def _bbox_from_mask_payload(mask_payload: dict[str, Any], mask: np.ndarray) -> BBox:
    bbox = mask_payload.get("bbox")
    if bbox is not None and len(bbox) >= 4:
        x, y, width, height = bbox[:4]
        return round(x), round(y), round(width), round(height)
    coords = cv2.findNonZero(mask)
    if coords is None:
        return 0, 0, 0, 0
    return cv2.boundingRect(coords)


def _optional_float(value: object) -> float | None:
    if value is None:
        return None
    return float(value)


def _polygon_from_mask(mask: np.ndarray) -> tuple[tuple[int, int], ...]:
    contours, _hierarchy = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return ()
    contour = max(contours, key=cv2.contourArea)
    epsilon = max(1.0, cv2.arcLength(contour, True) * 0.015)
    approximation = cv2.approxPolyDP(contour, epsilon, True)
    points = approximation.reshape(-1, 2)
    return tuple((int(x), int(y)) for x, y in points)


def _score_learned_proposals(
    proposals: list[LearnedProposal],
) -> tuple[list[LearnedProposal], dict[str, int]]:
    if not proposals:
        return proposals, {
            "size_cluster_count": 0,
            "repeated_size_member_count": 0,
            "alignment_supported_count": 0,
        }
    size_clusters = _group_similar_sizes(_proposal_adapters(proposals))
    row_clusters = _group_similar_centers(_proposal_adapters(proposals), axis="y")
    column_clusters = _group_similar_centers(_proposal_adapters(proposals), axis="x")
    size_cluster_by_index = _index_cluster_sizes(size_clusters)
    row_cluster_by_index = _index_cluster_sizes(row_clusters)
    column_cluster_by_index = _index_cluster_sizes(column_clusters)

    scored: list[LearnedProposal] = []
    for index, proposal in enumerate(proposals):
        size_cluster_size = size_cluster_by_index[index]
        row_alignment_size = row_cluster_by_index[index]
        column_alignment_size = column_cluster_by_index[index]
        confidence = proposal.predicted_iou if proposal.predicted_iou is not None else 0.5
        stability = proposal.stability_score if proposal.stability_score is not None else 0.5
        repeated_bonus = min(3, size_cluster_size - 1) * 0.35
        alignment_bonus = min(3, max(row_alignment_size, column_alignment_size) - 1) * 0.25
        score = (
            confidence * 0.9
            + stability * 0.8
            + proposal.solidity * 0.6
            + repeated_bonus
            + alignment_bonus
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
    scored.sort(key=lambda item: (-item.score, item.bbox[1], item.bbox[0], item.area))
    return scored, {
        "size_cluster_count": sum(1 for cluster in size_clusters if len(cluster) > 1),
        "repeated_size_member_count": sum(
            1 for proposal in scored if proposal.size_cluster_size > 1
        ),
        "alignment_supported_count": sum(
            1
            for proposal in scored
            if proposal.row_alignment_size > 1 or proposal.column_alignment_size > 1
        ),
    }


def _proposal_adapters(proposals: list[LearnedProposal]) -> list[Any]:
    return [type("_ProposalAdapter", (), {"bbox": proposal.bbox}) for proposal in proposals]


def _remove_duplicate_and_nested_masks(
    proposals: list[LearnedProposal],
) -> tuple[list[LearnedProposal], dict[str, int]]:
    kept: list[LearnedProposal] = []
    removed_by_iou = 0
    removed_by_nested = 0
    for candidate in proposals:
        if any(_bbox_iou(candidate.bbox, existing.bbox) >= 0.82 for existing in kept):
            removed_by_iou += 1
            continue
        nested_replacement_index = _nested_replacement_index(candidate, kept)
        if nested_replacement_index is not None:
            kept[nested_replacement_index] = candidate
            removed_by_nested += 1
            continue
        if any(_is_nested_duplicate(candidate, existing) for existing in kept):
            removed_by_nested += 1
            continue
        kept.append(candidate)
    return kept, {
        "removed_by_iou": removed_by_iou,
        "removed_by_nested": removed_by_nested,
    }


def _is_nested_duplicate(candidate: LearnedProposal, existing: LearnedProposal) -> bool:
    coverage = _smaller_coverage(candidate.bbox, existing.bbox)
    if coverage < 0.94:
        return False
    candidate_area = candidate.bbox[2] * candidate.bbox[3]
    existing_area = existing.bbox[2] * existing.bbox[3]
    smaller_area = min(candidate_area, existing_area)
    larger_area = max(candidate_area, existing_area)
    if smaller_area / max(1, larger_area) >= 0.72:
        return False
    return candidate_area < existing_area


def _nested_replacement_index(
    candidate: LearnedProposal,
    kept: list[LearnedProposal],
) -> int | None:
    candidate_area = candidate.bbox[2] * candidate.bbox[3]
    for index, existing in enumerate(kept):
        coverage = _smaller_coverage(candidate.bbox, existing.bbox)
        if coverage < 0.94:
            continue
        existing_area = existing.bbox[2] * existing.bbox[3]
        smaller_area = min(candidate_area, existing_area)
        larger_area = max(candidate_area, existing_area)
        if smaller_area / max(1, larger_area) >= 0.72:
            continue
        if candidate_area > existing_area:
            return index
    return None


def _structurally_supported_proposals(
    proposals: list[LearnedProposal],
) -> list[LearnedProposal]:
    return sorted(
        proposals,
        key=lambda proposal: (
            proposal.bbox[1] + proposal.bbox[3] / 2,
            proposal.bbox[0] + proposal.bbox[2] / 2,
            -proposal.score,
        ),
    )


def _proposal_to_original(
    proposal: LearnedProposal,
    *,
    scale: float,
    original_width: int,
    original_height: int,
) -> LearnedProposal:
    bbox = _map_bbox_to_original(
        proposal.bbox,
        scale=scale,
        original_width=original_width,
        original_height=original_height,
    )
    if math.isclose(scale, 1.0):
        polygon = proposal.polygon
    else:
        polygon = tuple(
            (
                max(0, min(round(x / scale), original_width)),
                max(0, min(round(y / scale), original_height)),
            )
            for x, y in proposal.polygon
        )
    return replace(proposal, bbox=bbox, polygon=polygon, mask=None)


def _render_mask_overlay(image: np.ndarray, proposals: list[LearnedProposal]) -> bytes:
    overlay = image.copy()
    palette = [
        (0, 128, 255),
        (80, 200, 120),
        (220, 90, 90),
        (180, 110, 230),
        (40, 190, 220),
        (235, 170, 40),
    ]
    for index, proposal in enumerate(proposals):
        color = np.array(palette[index % len(palette)], dtype=np.uint8)
        if proposal.mask is not None and proposal.mask.shape[:2] == image.shape[:2]:
            mask = proposal.mask.astype(bool)
            overlay[mask] = (overlay[mask] * 0.55 + color * 0.45).astype(np.uint8)
        x, y, width, height = proposal.bbox
        cv2.rectangle(overlay, (x, y), (x + width, y + height), tuple(int(c) for c in color), 1)
    return _encode_png(overlay)


def _render_bbox_overlay(image: np.ndarray, bboxes: list[BBox]) -> bytes:
    overlay = image.copy()
    line_width = max(1, round(min(image.shape[:2]) / 350))
    for index, (x, y, width, height) in enumerate(bboxes, start=1):
        cv2.rectangle(overlay, (x, y), (x + width, y + height), (0, 128, 255), line_width)
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


def _bbox_polygon(bbox: BBox) -> tuple[tuple[int, int], ...]:
    x, y, width, height = bbox
    return (
        (x, y),
        (x + width, y),
        (x + width, y + height),
        (x, y + height),
    )
