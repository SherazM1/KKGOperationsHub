"""PDF conversion service for Standard-mode generated DOCX files."""

from __future__ import annotations

import platform
import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path
from tempfile import mkdtemp

from app.services.bol_standard_docx_generator import GeneratedDocxFile


@dataclass(slots=True)
class ConvertedPdfFile:
    bol_number: str
    file_name: str
    file_path: str


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

    @property
    def converted_count(self) -> int:
        return len(self.converted_files)

    @property
    def failed_count(self) -> int:
        return len(self.failed_conversions)


def _convert_with_docx2pdf(source_docx: Path, destination_pdf: Path) -> None:
    from docx2pdf import convert as docx2pdf_convert  # type: ignore

    docx2pdf_convert(str(source_docx), str(destination_pdf))


def _convert_with_win32com(source_docx: Path, destination_pdf: Path) -> None:
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    wd_format_pdf = 17
    pythoncom.CoInitialize()
    word_app = win32com.client.DispatchEx("Word.Application")
    word_app.Visible = False
    document = None
    try:
        document = word_app.Documents.Open(str(source_docx.resolve()))
        document.SaveAs(str(destination_pdf.resolve()), FileFormat=wd_format_pdf)
    finally:
        if document is not None:
            document.Close(False)
        word_app.Quit()
        pythoncom.CoUninitialize()


def _convert_with_soffice(source_docx: Path, destination_pdf: Path) -> None:
    soffice_path = shutil.which("soffice")
    if not soffice_path:
        raise RuntimeError("LibreOffice 'soffice' executable was not found.")

    destination_dir = destination_pdf.parent
    destination_dir.mkdir(parents=True, exist_ok=True)

    result = subprocess.run(
        [
            soffice_path,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(destination_dir),
            str(source_docx),
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "Unknown soffice error.")

    soffice_output = destination_dir / f"{source_docx.stem}.pdf"
    if not soffice_output.exists():
        raise RuntimeError("Conversion completed without creating a PDF output file.")
    if soffice_output.resolve() != destination_pdf.resolve():
        soffice_output.replace(destination_pdf)


def _resolve_converter() -> tuple[str | None, str | None]:
    try:
        import docx2pdf  # noqa: F401  # type: ignore

        return "docx2pdf", None
    except Exception:
        pass

    if platform.system().lower().startswith("win"):
        try:
            import pythoncom  # noqa: F401  # type: ignore
            import win32com.client  # noqa: F401  # type: ignore

            return "win32com", None
        except Exception:
            pass

    if shutil.which("soffice"):
        return "soffice", None

    return (
        None,
        "PDF conversion is unavailable in this environment. Install docx2pdf, "
        "or pywin32 with Microsoft Word on Windows, or LibreOffice (soffice).",
    )


def _run_conversion(converter_name: str, source_docx: Path, destination_pdf: Path) -> None:
    if converter_name == "docx2pdf":
        _convert_with_docx2pdf(source_docx, destination_pdf)
        return
    if converter_name == "win32com":
        _convert_with_win32com(source_docx, destination_pdf)
        return
    if converter_name == "soffice":
        _convert_with_soffice(source_docx, destination_pdf)
        return
    raise RuntimeError(f"Unsupported converter: {converter_name}")


def convert_standard_docx_set_to_pdf(
    generated_docx_files: list[GeneratedDocxFile],
    output_dir: Path | None = None,
) -> StandardPdfConversionResult:
    output_root = output_dir or Path(mkdtemp(prefix="kkg_standard_bol_pdf_"))
    output_root.mkdir(parents=True, exist_ok=True)

    if not generated_docx_files:
        raise ValueError("No generated DOCX files were provided for PDF conversion.")

    converter_name, unavailable_reason = _resolve_converter()
    if converter_name is None:
        return StandardPdfConversionResult(
            output_dir=str(output_root.resolve()),
            converted_files=[],
            failed_conversions=[],
            converter_name=None,
            conversion_available=False,
            unavailable_reason=unavailable_reason,
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
            _run_conversion(converter_name, source_docx, destination_pdf)
            converted_files.append(
                ConvertedPdfFile(
                    bol_number=docx_file.bol_number,
                    file_name=destination_pdf.name,
                    file_path=str(destination_pdf.resolve()),
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
    )

