"""PDF merge helpers for generated BOL PDFs."""

from __future__ import annotations

from pathlib import Path
from typing import Sequence

from pypdf import PdfWriter

from app.services.bol_standard_pdf_converter import ConvertedPdfFile


def merge_pdf_files(
    pdf_files: Sequence[ConvertedPdfFile],
    output_path: Path,
) -> Path:
    """Merge existing PDF files in the provided order without re-rendering pages."""

    existing_pdf_paths = [Path(pdf_file.file_path) for pdf_file in pdf_files if Path(pdf_file.file_path).exists()]
    if not existing_pdf_paths:
        raise ValueError("No generated PDF files are available to merge.")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    writer = PdfWriter()

    for pdf_path in existing_pdf_paths:
        writer.append(pdf_path, import_outline=False)

    with output_path.open("wb") as output_file:
        writer.write(output_file)

    return output_path
