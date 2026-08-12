"""Filesystem-backed baseline storage for Display Compliance."""

from __future__ import annotations

from dataclasses import asdict
import json
from pathlib import Path
import re

from app.display_compliance.models import DisplayBaseline, ProductRegion


DEFAULT_BASELINE_STORAGE_ROOT = Path("data") / "display_compliance" / "baselines"


class LocalBaselineStorage:
    """Local development storage for baseline metadata and reference images."""

    def __init__(self, root: Path = DEFAULT_BASELINE_STORAGE_ROOT) -> None:
        self.root = root

    def save_baseline(self, baseline: DisplayBaseline, image_bytes: bytes) -> None:
        baseline_dir = self.root / baseline.baseline_id
        baseline_dir.mkdir(parents=True, exist_ok=True)
        metadata_path = baseline_dir / "baseline.json"
        image_path = baseline_dir / _safe_reference_filename(baseline.reference_filename)

        image_path.write_bytes(image_bytes)
        metadata_path.write_text(
            json.dumps(asdict(baseline), indent=2, sort_keys=True),
            encoding="utf-8",
        )

    def save_baseline_metadata(self, baseline: DisplayBaseline) -> None:
        baseline_dir = self.root / baseline.baseline_id
        baseline_dir.mkdir(parents=True, exist_ok=True)
        metadata_path = baseline_dir / "baseline.json"
        metadata_path.write_text(
            json.dumps(asdict(baseline), indent=2, sort_keys=True),
            encoding="utf-8",
        )

    def load_baseline(self, baseline_id: str) -> DisplayBaseline:
        metadata_path = self.root / baseline_id / "baseline.json"
        with metadata_path.open("r", encoding="utf-8") as metadata_file:
            payload = json.load(metadata_file)
        return baseline_from_dict(payload)

    def load_reference_image_bytes(self, baseline: DisplayBaseline) -> bytes:
        image_path = self.root / baseline.baseline_id / _safe_reference_filename(
            baseline.reference_filename
        )
        return image_path.read_bytes()

    def list_baselines(self) -> list[DisplayBaseline]:
        if not self.root.exists():
            return []

        baselines: list[DisplayBaseline] = []
        for metadata_path in sorted(self.root.glob("*/baseline.json")):
            with metadata_path.open("r", encoding="utf-8") as metadata_file:
                baselines.append(baseline_from_dict(json.load(metadata_file)))
        return baselines


def baseline_from_dict(payload: dict[str, object]) -> DisplayBaseline:
    """Convert stored JSON metadata into a DisplayBaseline."""
    regions = [
        ProductRegion(
            region_id=str(region.get("region_id", "")),
            bbox=tuple(region["bbox"]) if region.get("bbox") else None,
            polygon=tuple(tuple(point) for point in region.get("polygon", [])),
            label=str(region.get("label", "")),
            visual_signature=(
                str(region["visual_signature"])
                if region.get("visual_signature") is not None
                else None
            ),
        )
        for region in payload.get("regions", [])
        if isinstance(region, dict)
    ]
    return DisplayBaseline(
        baseline_id=str(payload["baseline_id"]),
        name=str(payload["name"]),
        reference_filename=str(payload["reference_filename"]),
        reference_width=int(payload["reference_width"]),
        reference_height=int(payload["reference_height"]),
        regions=regions,
    )


def _safe_reference_filename(filename: str) -> str:
    name = Path(filename).name or "reference_image"
    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("._")
    return safe_name or "reference_image"
