"""Tests for the Display Compliance baseline scaffold."""

from __future__ import annotations

from pathlib import Path
import struct
import zlib

import pytest

from app.display_compliance.baseline import create_baseline
from app.display_compliance.models import DisplayBaseline
from app.display_compliance.storage import LocalBaselineStorage


def _png_bytes(width: int = 3, height: int = 2) -> bytes:
    def chunk(chunk_type: bytes, data: bytes) -> bytes:
        checksum = zlib.crc32(chunk_type + data) & 0xFFFFFFFF
        return struct.pack(">I", len(data)) + chunk_type + data + struct.pack(">I", checksum)

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    raw_rows = b"".join(b"\x00" + (b"\xff\x00\x00" * width) for _ in range(height))
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(
        b"IDAT", zlib.compress(raw_rows)
    ) + chunk(b"IEND", b"")


def test_create_baseline_from_valid_image_records_domain_metadata() -> None:
    baseline = create_baseline(
        name="Endcap Reference",
        filename="reference.png",
        image_bytes=_png_bytes(7, 5),
    )

    assert isinstance(baseline, DisplayBaseline)
    assert baseline.name == "Endcap Reference"
    assert baseline.reference_filename == "reference.png"
    assert baseline.reference_width == 7
    assert baseline.reference_height == 5
    assert baseline.regions == []
    assert baseline.baseline_id


def test_blank_baseline_name_is_rejected() -> None:
    with pytest.raises(ValueError, match="name is required"):
        create_baseline(name="  ", filename="reference.png", image_bytes=_png_bytes())


@pytest.mark.parametrize("image_bytes", [b"", b"not an image"])
def test_empty_or_invalid_image_bytes_are_rejected(image_bytes: bytes) -> None:
    with pytest.raises(ValueError):
        create_baseline(name="Reference", filename="reference.png", image_bytes=image_bytes)


def test_baseline_metadata_round_trips_through_local_storage(tmp_path: Path) -> None:
    image_bytes = _png_bytes(4, 6)
    baseline = create_baseline(
        name="Baseline A",
        filename="baseline a.png",
        image_bytes=image_bytes,
    )
    storage = LocalBaselineStorage(tmp_path)

    storage.save_baseline(baseline, image_bytes)
    loaded = storage.load_baseline(baseline.baseline_id)

    assert loaded == baseline
    assert storage.list_baselines() == [baseline]
    assert (tmp_path / baseline.baseline_id / "baseline.json").exists()
    assert (tmp_path / baseline.baseline_id / "baseline_a.png").read_bytes() == image_bytes
