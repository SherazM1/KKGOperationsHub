"""Baseline creation logic for Display Compliance."""

from __future__ import annotations

import struct
from uuid import uuid4

from app.display_compliance.models import DisplayBaseline


def create_baseline(
    *,
    name: str,
    filename: str,
    image_bytes: bytes,
) -> DisplayBaseline:
    """Create a reference-display baseline from uploaded image bytes."""
    clean_name = name.strip()
    if not clean_name:
        raise ValueError("Display / baseline name is required.")
    if not image_bytes:
        raise ValueError("A reference image is required.")

    width, height = _read_image_dimensions(image_bytes)
    return DisplayBaseline(
        baseline_id=uuid4().hex,
        name=clean_name,
        reference_filename=filename,
        reference_width=width,
        reference_height=height,
        regions=[],
    )


def _read_image_dimensions(image_bytes: bytes) -> tuple[int, int]:
    if image_bytes.startswith(b"\x89PNG\r\n\x1a\n"):
        return _read_png_dimensions(image_bytes)
    if image_bytes.startswith(b"\xff\xd8"):
        return _read_jpeg_dimensions(image_bytes)
    raise ValueError("Reference image must be a valid PNG or JPEG file.")


def _read_png_dimensions(image_bytes: bytes) -> tuple[int, int]:
    if len(image_bytes) < 24 or image_bytes[12:16] != b"IHDR":
        raise ValueError("Reference image is not a valid PNG file.")
    width, height = struct.unpack(">II", image_bytes[16:24])
    if width <= 0 or height <= 0:
        raise ValueError("Reference image has invalid dimensions.")
    return width, height


def _read_jpeg_dimensions(image_bytes: bytes) -> tuple[int, int]:
    index = 2
    length = len(image_bytes)
    while index < length:
        if image_bytes[index] != 0xFF:
            raise ValueError("Reference image is not a valid JPEG file.")
        while index < length and image_bytes[index] == 0xFF:
            index += 1
        if index >= length:
            break

        marker = image_bytes[index]
        index += 1
        if marker in {0xD8, 0xD9} or 0xD0 <= marker <= 0xD7:
            continue
        if index + 2 > length:
            break

        segment_length = int.from_bytes(image_bytes[index : index + 2], "big")
        if segment_length < 2 or index + segment_length > length:
            break

        if marker in {
            0xC0,
            0xC1,
            0xC2,
            0xC3,
            0xC5,
            0xC6,
            0xC7,
            0xC9,
            0xCA,
            0xCB,
            0xCD,
            0xCE,
            0xCF,
        }:
            if segment_length < 7:
                break
            height = int.from_bytes(image_bytes[index + 3 : index + 5], "big")
            width = int.from_bytes(image_bytes[index + 5 : index + 7], "big")
            if width <= 0 or height <= 0:
                raise ValueError("Reference image has invalid dimensions.")
            return width, height

        index += segment_length

    raise ValueError("Reference image is not a valid JPEG file.")

