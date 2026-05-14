from __future__ import annotations

from dataclasses import fields, is_dataclass
from pathlib import Path
from tempfile import mkdtemp

from pypdf import PdfReader
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

from app.ui import bol_generator
from app.services.bol_file_bundle_service import create_multistop_bundles, create_standard_bundles
from app.services.bol_multistop_docx_generator import MultistopGeneratedDocxFile
from app.services.bol_pdf_merge_service import merge_pdf_files
from app.services.bol_standard_docx_generator import GeneratedDocxFile
from app.services.bol_standard_pdf_converter import ConvertedPdfFile


def _write_pdf(path: Path, pages: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    canv = canvas.Canvas(str(path), pagesize=letter)
    for page_text in pages:
        canv.drawString(72, 720, page_text)
        canv.showPage()
    canv.save()


def _pdf_text(path: Path) -> str:
    reader = PdfReader(str(path))
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def _converted(path: Path, bol_number: str) -> ConvertedPdfFile:
    return ConvertedPdfFile(
        bol_number=bol_number,
        file_name=path.name,
        file_path=str(path),
    )


def test_merge_pdf_files_combines_pages_and_preserves_input_order(tmp_path: Path) -> None:
    first_pdf = tmp_path / "first.pdf"
    second_pdf = tmp_path / "second.pdf"
    _write_pdf(first_pdf, ["FIRST PAGE"])
    _write_pdf(second_pdf, ["SECOND PAGE", "THIRD PAGE"])

    output_path = merge_pdf_files(
        [_converted(first_pdf, "BOL-1"), _converted(second_pdf, "BOL-2")],
        tmp_path / "combined.pdf",
    )

    reader = PdfReader(str(output_path))
    assert len(reader.pages) == 3
    text = _pdf_text(output_path)
    assert text.index("FIRST PAGE") < text.index("SECOND PAGE") < text.index("THIRD PAGE")


def test_standard_bundle_exposes_combined_pdf_and_preserves_pdf_zip(tmp_path: Path) -> None:
    docx_path = tmp_path / "standard_bol_1.docx"
    docx_path.write_bytes(b"docx")
    first_pdf = tmp_path / "standard_bol_1.pdf"
    second_pdf = tmp_path / "standard_bol_2.pdf"
    _write_pdf(first_pdf, ["BOL ONE"])
    _write_pdf(second_pdf, ["BOL TWO"])

    bundle = create_standard_bundles(
        generated_docx_files=[
            GeneratedDocxFile(
                bol_number="BOL-1",
                file_name=docx_path.name,
                file_path=str(docx_path),
            )
        ],
        converted_pdf_files=[_converted(first_pdf, "BOL-1"), _converted(second_pdf, "BOL-2")],
        output_dir=tmp_path / "bundles",
        bundle_name_prefix="no_recourse_bol",
        batch_name="May 14 Batch",
    )

    assert bundle.pdf_bundle is not None
    assert Path(bundle.pdf_bundle.file_path).exists()
    assert bundle.combined_pdf is not None
    assert bundle.combined_pdf.file_name == "May_14_Batch_no_recourse_bol_combined.pdf"
    assert Path(bundle.combined_pdf.file_path).exists()
    assert not any(field.name in {"data", "bytes", "content"} for field in fields(bundle.combined_pdf))
    assert len(PdfReader(bundle.combined_pdf.file_path).pages) == 2
    assert _pdf_text(Path(bundle.combined_pdf.file_path)).index("BOL ONE") < _pdf_text(
        Path(bundle.combined_pdf.file_path)
    ).index("BOL TWO")


def test_multistop_bundle_uses_same_combined_pdf_merge_behavior(tmp_path: Path) -> None:
    docx_path = tmp_path / "combined_multistop_bol_1.docx"
    docx_path.write_bytes(b"docx")
    first_pdf = tmp_path / "combined_multistop_bol_1.pdf"
    second_pdf = tmp_path / "combined_multistop_bol_2.pdf"
    _write_pdf(first_pdf, ["MULTI ONE"])
    _write_pdf(second_pdf, ["MULTI TWO"])

    bundle = create_multistop_bundles(
        generated_docx_files=[
            MultistopGeneratedDocxFile(
                bol_number="MBOL-1",
                file_name=docx_path.name,
                file_path=str(docx_path),
                document_type="combined",
                load_number="LOAD-1",
                stop_number=None,
            )
        ],
        converted_pdf_files=[
            ConvertedPdfFile(
                bol_number="MBOL-1",
                file_name=first_pdf.name,
                file_path=str(first_pdf),
                document_type="combined",
                load_number="LOAD-1",
            ),
            ConvertedPdfFile(
                bol_number="MBOL-2",
                file_name=second_pdf.name,
                file_path=str(second_pdf),
                document_type="combined",
                load_number="LOAD-2",
            ),
        ],
        output_dir=tmp_path / "bundles",
        batch_name="Multi Batch",
    )

    assert bundle.pdf_bundle is not None
    assert bundle.combined_pdf is not None
    assert bundle.combined_pdf.file_name == "Multi_Batch_multistop_bol_combined.pdf"
    text = _pdf_text(Path(bundle.combined_pdf.file_path))
    assert text.index("MULTI ONE") < text.index("MULTI TWO")


def test_bundle_artifacts_store_paths_not_payload_bytes(tmp_path: Path) -> None:
    docx_path = tmp_path / "standard_bol_1.docx"
    pdf_path = tmp_path / "standard_bol_1.pdf"
    docx_path.write_bytes(b"docx")
    _write_pdf(pdf_path, ["BOL ONE"])

    bundle = create_standard_bundles(
        generated_docx_files=[
            GeneratedDocxFile(
                bol_number="BOL-1",
                file_name=docx_path.name,
                file_path=str(docx_path),
            )
        ],
        converted_pdf_files=[_converted(pdf_path, "BOL-1")],
        output_dir=tmp_path / "bundles",
    )

    for artifact in (
        bundle.docx_bundle,
        bundle.pdf_bundle,
        bundle.all_files_bundle,
        bundle.combined_pdf,
    ):
        assert artifact is not None
        assert is_dataclass(artifact)
        assert isinstance(artifact.file_path, str)
        assert Path(artifact.file_path).exists()
        assert not any(isinstance(getattr(artifact, field.name), bytes) for field in fields(artifact))


def test_clear_generated_artifacts_removes_only_recorded_bol_temp_dirs(tmp_path: Path) -> None:
    safe_temp_dir = Path(mkdtemp(prefix="kkg_standard_bol_bundles_"))
    safe_file = safe_temp_dir / "bundle.zip"
    safe_file.write_bytes(b"zip")
    unrelated_dir = tmp_path / "user_uploads"
    unrelated_dir.mkdir()
    unrelated_file = unrelated_dir / "source.xlsx"
    unrelated_file.write_bytes(b"input")

    bol_generator.st.session_state["bol_generation_output_dirs"] = [
        str(safe_temp_dir),
        str(unrelated_dir),
    ]
    bol_generator.st.session_state["bol_docx_result"] = object()
    bol_generator.st.session_state["bol_pdf_result"] = object()
    bol_generator.st.session_state["bol_bundle_result"] = object()
    bol_generator.st.session_state["bol_bundle_error"] = "stale"
    bol_generator.st.session_state["bol_all_files_bundle_requested"] = True

    bol_generator._clear_generated_artifacts(delete_files=True)

    assert not safe_temp_dir.exists()
    assert unrelated_file.exists()
    assert bol_generator.st.session_state["bol_generation_output_dirs"] == []
    assert bol_generator.st.session_state["bol_docx_result"] is None
    assert bol_generator.st.session_state["bol_pdf_result"] is None
    assert bol_generator.st.session_state["bol_bundle_result"] is None
    assert bol_generator.st.session_state["bol_bundle_error"] is None
    assert bol_generator.st.session_state["bol_all_files_bundle_requested"] is False
