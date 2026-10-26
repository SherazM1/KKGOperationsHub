"""Generate SKID tag documents as a single batch or separate DC batches."""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
import re
from zipfile import ZIP_DEFLATED, ZipFile

from app.models.skid_tag import SkidTag
from app.services.docx_generator_skid_tags import generate_skid_tags_docx
from app.services.pdf_generator_skid_tags import generate_skid_tags_pdf


@dataclass(frozen=True)
class SkidTagBatch:
    dc: str | None
    file_stem: str
    tag_count: int
    docx: bytes
    pdf: bytes


def generate_skid_tag_batches(
    tags: list[SkidTag], *, group_by_dc: bool = False
) -> list[SkidTagBatch]:
    """Preserve tag order and pallet numbering within each DC's batch."""
    if not tags:
        raise ValueError("No SKID tags provided for generation.")
    groups: dict[str | None, list[SkidTag]] = {}
    for tag in tags:
        groups.setdefault(tag.dc if group_by_dc else None, []).append(tag)

    batches = []
    used_stems: set[str] = set()
    for dc, dc_tags in groups.items():
        stem = "skid_tags"
        if dc is not None:
            safe_dc = re.sub(r"[^A-Za-z0-9_-]+", "_", dc).strip("_") or "unknown"
            stem += f"_DC_{safe_dc}"
        unique_stem = stem
        suffix = 2
        while unique_stem.casefold() in used_stems:
            unique_stem = f"{stem}_{suffix}"
            suffix += 1
        used_stems.add(unique_stem.casefold())
        batches.append(SkidTagBatch(
            dc=dc,
            file_stem=unique_stem,
            tag_count=len(dc_tags),
            docx=generate_skid_tags_docx(dc_tags),
            pdf=generate_skid_tags_pdf(dc_tags),
        ))
    return batches


def zip_skid_tag_batches(batches: list[SkidTagBatch], *, include_docx: bool = True) -> bytes:
    """Package every DC's PDFs, optionally including the matching Word files."""
    buffer = BytesIO()
    with ZipFile(buffer, "w", compression=ZIP_DEFLATED) as archive:
        for batch in batches:
            if include_docx:
                archive.writestr(f"{batch.file_stem}.docx", batch.docx)
            archive.writestr(f"{batch.file_stem}.pdf", batch.pdf)
    return buffer.getvalue()
