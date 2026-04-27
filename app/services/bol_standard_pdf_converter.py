"""PDF conversion service for Standard-mode generated DOCX files."""

from __future__ import annotations

import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path
from tempfile import mkdtemp
from urllib.parse import quote

from app.services.bol_standard_docx_generator import GeneratedDocxFile


LIBREOFFICE_EXECUTABLE_NAMES: tuple[str, ...] = ("soffice", "libreoffice")
LIBREOFFICE_CONVERSION_TIMEOUT_SECONDS = 120


@dataclass(slots=True)
class ConvertedPdfFile:
    bol_number: str
    file_name: str
    file_path: str
    document_type: str = ""
    load_number: str = ""
    stop_number: int | None = None


@dataclass(slots=True)
class FailedPdfConversion:
    bol_number: str
    source_docx: str
    error: str


@dataclass(slots=True)
class StandardPdfConversionResult:
    output_dir: str
    converted_files: list[ConvertedPdfFile]
    failed_conversions: list[FailedPdfConversion]
    converter_name: str | None
    conversion_available: bool
    unavailable_reason: str | None
    converter_path: str | None = None

    @property
    def converted_count(self) -> int:
        return len(self.converted_files)

    @property
    def failed_count(self) -> int:
        return len(self.failed_conversions)


def _find_libreoffice_executable() -> tuple[str | None, str | None]:
    for executable_name in LIBREOFFICE_EXECUTABLE_NAMES:
        executable_path = shutil.which(executable_name)
        if executable_path:
            return executable_name, executable_path
    return None, None


def _libreoffice_profile_uri(profile_dir: Path) -> str:
    return "file:///" + quote(str(profile_dir.resolve()).replace("\\", "/"))


def _convert_with_libreoffice(
    libreoffice_path: str,
    source_docx: Path,
    destination_pdf: Path,
) -> None:
    destination_dir = destination_pdf.parent
    destination_dir.mkdir(parents=True, exist_ok=True)

    if destination_pdf.exists():
        destination_pdf.unlink()

    profile_dir = Path(mkdtemp(prefix="kkg_libreoffice_profile_"))
    try:
        result = subprocess.run(
            [
                libreoffice_path,
                "--headless",
                "--nologo",
                "--nofirststartwizard",
                "--nolockcheck",
                f"-env:UserInstallation={_libreoffice_profile_uri(profile_dir)}",
                "--convert-to",
                "pdf",
                "--outdir",
                str(destination_dir),
                str(source_docx),
            ],
            capture_output=True,
            text=True,
            check=False,
            timeout=LIBREOFFICE_CONVERSION_TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(
            "LibreOffice conversion timed out after "
            f"{LIBREOFFICE_CONVERSION_TIMEOUT_SECONDS} seconds."
        ) from exc
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)

    if result.returncode != 0:
        message = result.stderr.strip() or result.stdout.strip() or "Unknown LibreOffice error."
        raise RuntimeError(message)

    libreoffice_output = destination_dir / f"{source_docx.stem}.pdf"
    if not libreoffice_output.exists():
        output_text = result.stdout.strip() or result.stderr.strip()
        detail = f" LibreOffice output: {output_text}" if output_text else ""
        raise RuntimeError("Conversion completed without creating a PDF output file." + detail)

    if libreoffice_output.resolve() != destination_pdf.resolve():
        if destination_pdf.exists():
            destination_pdf.unlink()
        libreoffice_output.replace(destination_pdf)

    if not destination_pdf.exists():
        raise RuntimeError("Conversion completed but the expected PDF file is missing.")


def _resolve_converter() -> tuple[str | None, str | None, str | None]:
    converter_name, converter_path = _find_libreoffice_executable()
    if converter_name and converter_path:
        return converter_name, converter_path, None

    return (
        None,
        None,
        "PDF conversion is unavailable because LibreOffice was not found. "
        "Install LibreOffice and make either 'soffice' or 'libreoffice' available on PATH.",
    )


def _run_conversion(converter_path: str, source_docx: Path, destination_pdf: Path) -> None:
    _convert_with_libreoffice(converter_path, source_docx, destination_pdf)


def convert_standard_docx_set_to_pdf(
    generated_docx_files: list[GeneratedDocxFile],
    output_dir: Path | None = None,
) -> StandardPdfConversionResult:
    output_root = output_dir or Path(mkdtemp(prefix="kkg_standard_bol_pdf_"))
    output_root.mkdir(parents=True, exist_ok=True)

    if not generated_docx_files:
        raise ValueError("No generated DOCX files were provided for PDF conversion.")

    converter_name, converter_path, unavailable_reason = _resolve_converter()
    if converter_name is None:
        return StandardPdfConversionResult(
            output_dir=str(output_root.resolve()),
            converted_files=[],
            failed_conversions=[],
            converter_name=None,
            conversion_available=False,
            unavailable_reason=unavailable_reason,
            converter_path=None,
        )

    converted_files: list[ConvertedPdfFile] = []
    failed_conversions: list[FailedPdfConversion] = []

    for docx_file in generated_docx_files:
        source_docx = Path(docx_file.file_path)
        destination_pdf = output_root / f"{source_docx.stem}.pdf"

        if not source_docx.exists():
            failed_conversions.append(
                FailedPdfConversion(
                    bol_number=docx_file.bol_number,
                    source_docx=str(source_docx),
                    error="Source DOCX file does not exist.",
                )
            )
            continue

        try:
            if converter_path is None:
                raise RuntimeError("LibreOffice executable path was not resolved.")
            _run_conversion(converter_path, source_docx, destination_pdf)
            converted_files.append(
                ConvertedPdfFile(
                    bol_number=docx_file.bol_number,
                    file_name=destination_pdf.name,
                    file_path=str(destination_pdf.resolve()),
                    document_type=str(getattr(docx_file, "document_type", "") or ""),
                    load_number=str(getattr(docx_file, "load_number", "") or ""),
                    stop_number=getattr(docx_file, "stop_number", None),
                )
            )
        except Exception as exc:
            failed_conversions.append(
                FailedPdfConversion(
                    bol_number=docx_file.bol_number,
                    source_docx=str(source_docx),
                    error=str(exc),
                )
            )

    return StandardPdfConversionResult(
        output_dir=str(output_root.resolve()),
        converted_files=converted_files,
        failed_conversions=failed_conversions,
        converter_name=converter_name,
        conversion_available=True,
        unavailable_reason=None,
        converter_path=converter_path,
    )
