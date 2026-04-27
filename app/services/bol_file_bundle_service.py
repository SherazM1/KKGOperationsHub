"""Bundle packaging service for BOL generated files."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from tempfile import mkdtemp
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile

from app.services.bol_standard_docx_generator import GeneratedDocxFile
from app.services.bol_standard_pdf_converter import ConvertedPdfFile


DEFAULT_BUNDLE_NAME_PREFIX = "standard_bol"


@dataclass(slots=True)
class BundleArtifact:
    bundle_type: str
    file_name: str
    file_path: str
    file_count: int
    group_count: int = 0
    combined_count: int = 0
    stop_count: int = 0
    missing_count: int = 0


@dataclass(slots=True)
class StandardBundleResult:
    output_dir: str
    docx_bundle: BundleArtifact | None
    pdf_bundle: BundleArtifact | None
    all_files_bundle: BundleArtifact | None


def _build_zip(zip_path: Path, files: list[tuple[Path, str]]) -> BundleArtifact:
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    added_count = 0
    used_names: set[str] = set()
    with ZipFile(zip_path, "w", compression=ZIP_DEFLATED) as zip_file:
        for file_path, archive_name in files:
            if not file_path.exists():
                continue
            candidate_name = archive_name
            counter = 2
            while candidate_name in used_names:
                stem = Path(archive_name).stem
                suffix = Path(archive_name).suffix
                candidate_name = f"{stem}_{counter}{suffix}"
                counter += 1

            zip_file.write(file_path, arcname=candidate_name)
            used_names.add(candidate_name)
            added_count += 1

    return BundleArtifact(
        bundle_type=zip_path.stem,
        file_name=zip_path.name,
        file_path=str(zip_path.resolve()),
        file_count=added_count,
    )


def _sanitize_archive_part(value: str) -> str:
    cleaned = "".join(
        char if char.isalnum() or char in ("-", "_") else "_"
        for char in (value or "").strip()
    )
    cleaned = cleaned.strip("_")
    return cleaned or "unknown"


def _safe_archive_name(value: str, expected_suffix: str | None = ".docx") -> str:
    fallback_suffix = expected_suffix or ".docx"
    path = Path(value or f"generated{fallback_suffix}")
    stem = _sanitize_archive_part(path.stem)
    if expected_suffix is None:
        suffix = path.suffix
    else:
        suffix = path.suffix if path.suffix.lower() == expected_suffix else expected_suffix
    return f"{stem}{suffix}"


def _build_multistop_zip(
    zip_path: Path,
    generated_files: list[Any],
    expected_suffix: str | None,
) -> BundleArtifact:
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    added_count = 0
    missing_count = 0
    combined_count = 0
    stop_count = 0
    group_names: set[str] = set()
    used_archive_names: set[str] = set()

    with ZipFile(zip_path, "w", compression=ZIP_DEFLATED) as zip_file:
        for generated_file in generated_files:
            source_path = Path(generated_file.file_path)
            if not source_path.exists():
                missing_count += 1
                continue

            bol_number = _sanitize_archive_part(generated_file.bol_number)
            load_number = _sanitize_archive_part(
                str(getattr(generated_file, "load_number", "") or "")
            )
            folder_name = f"load_{load_number}_bol_{bol_number}"
            archive_file_name = _safe_archive_name(generated_file.file_name, expected_suffix)
            archive_name = f"{folder_name}/{archive_file_name}"

            counter = 2
            while archive_name in used_archive_names:
                stem = Path(archive_file_name).stem
                suffix = Path(archive_file_name).suffix
                archive_name = f"{folder_name}/{stem}_{counter}{suffix}"
                counter += 1

            zip_file.write(source_path, arcname=archive_name)
            used_archive_names.add(archive_name)
            group_names.add(folder_name)
            added_count += 1

            document_type = getattr(generated_file, "document_type", "")
            if document_type == "combined":
                combined_count += 1
            elif document_type == "stop":
                stop_count += 1

    return BundleArtifact(
        bundle_type=zip_path.stem,
        file_name=zip_path.name,
        file_path=str(zip_path.resolve()),
        file_count=added_count,
        group_count=len(group_names),
        combined_count=combined_count,
        stop_count=stop_count,
        missing_count=missing_count,
    )


def create_multistop_bundles(
    generated_docx_files: list[GeneratedDocxFile],
    converted_pdf_files: list[ConvertedPdfFile] | None = None,
    output_dir: Path | None = None,
    bundle_name_prefix: str = "multistop_bol",
    include_all_files_bundle: bool = True,
) -> StandardBundleResult:
    output_root = output_dir or Path(mkdtemp(prefix="kkg_multistop_bol_bundles_"))
    output_root.mkdir(parents=True, exist_ok=True)

    prefix = (bundle_name_prefix or "multistop_bol").strip()
    if not prefix:
        prefix = "multistop_bol"

    converted_pdf_files = converted_pdf_files or []
    existing_docx_files = [
        generated_file
        for generated_file in generated_docx_files
        if Path(generated_file.file_path).exists()
    ]
    existing_pdf_files = [
        converted_file
        for converted_file in converted_pdf_files
        if Path(converted_file.file_path).exists()
    ]
    docx_bundle = (
        _build_multistop_zip(
            output_root / f"{prefix}_docx_bundle.zip",
            generated_docx_files,
            ".docx",
        )
        if existing_docx_files
        else None
    )
    pdf_bundle = (
        _build_multistop_zip(
            output_root / f"{prefix}_pdf_bundle.zip",
            converted_pdf_files,
            ".pdf",
        )
        if existing_pdf_files
        else None
    )
    all_files_bundle = (
        _build_multistop_zip(
            output_root / f"{prefix}_all_files_bundle.zip",
            [*generated_docx_files, *converted_pdf_files],
            None,
        )
        if include_all_files_bundle and (existing_docx_files or existing_pdf_files)
        else None
    )

    return StandardBundleResult(
        output_dir=str(output_root.resolve()),
        docx_bundle=docx_bundle,
        pdf_bundle=pdf_bundle,
        all_files_bundle=all_files_bundle,
    )


def create_multistop_docx_bundle(
    generated_docx_files: list[GeneratedDocxFile],
    output_dir: Path | None = None,
    bundle_name_prefix: str = "multistop_bol",
) -> StandardBundleResult:
    return create_multistop_bundles(
        generated_docx_files=generated_docx_files,
        converted_pdf_files=[],
        output_dir=output_dir,
        bundle_name_prefix=bundle_name_prefix,
        include_all_files_bundle=True,
    )


def create_standard_bundles(
    generated_docx_files: list[GeneratedDocxFile],
    converted_pdf_files: list[ConvertedPdfFile],
    output_dir: Path | None = None,
    bundle_name_prefix: str = DEFAULT_BUNDLE_NAME_PREFIX,
    include_all_files_bundle: bool = True,
) -> StandardBundleResult:
    output_root = output_dir or Path(mkdtemp(prefix="kkg_standard_bol_bundles_"))
    output_root.mkdir(parents=True, exist_ok=True)

    docx_entries: list[tuple[Path, str]] = [
        (Path(file.file_path), file.file_name)
        for file in generated_docx_files
        if Path(file.file_path).exists()
    ]
    pdf_entries: list[tuple[Path, str]] = [
        (Path(file.file_path), file.file_name)
        for file in converted_pdf_files
        if Path(file.file_path).exists()
    ]

    prefix = (bundle_name_prefix or DEFAULT_BUNDLE_NAME_PREFIX).strip()
    if not prefix:
        prefix = DEFAULT_BUNDLE_NAME_PREFIX
    docx_bundle_filename = f"{prefix}_docx_bundle.zip"
    pdf_bundle_filename = f"{prefix}_pdf_bundle.zip"
    all_files_bundle_filename = f"{prefix}_all_files_bundle.zip"

    docx_bundle = (
        _build_zip(output_root / docx_bundle_filename, docx_entries)
        if docx_entries
        else None
    )
    pdf_bundle = (
        _build_zip(output_root / pdf_bundle_filename, pdf_entries)
        if pdf_entries
        else None
    )

    all_entries = docx_entries + pdf_entries
    all_bundle = (
        _build_zip(output_root / all_files_bundle_filename, all_entries)
        if include_all_files_bundle and all_entries
        else None
    )

    return StandardBundleResult(
        output_dir=str(output_root.resolve()),
        docx_bundle=docx_bundle,
        pdf_bundle=pdf_bundle,
        all_files_bundle=all_bundle,
    )
