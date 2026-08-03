"""Streamlit UI for the Spec Sheet Extractor scaffold."""

from __future__ import annotations

from dataclasses import asdict

import streamlit as st

from app.spec_sheet_extractor.extractor import (
    extract_header_fields_from_uploads,
    inventory_pdf_uploads,
)
from app.spec_sheet_extractor.models import (
    PdfFileInventoryResult,
    PdfHeaderExtractionResult,
    PdfPageInventoryRecord,
)


def _clear_inventory_state() -> None:
    st.session_state["spec_sheet_extractor_uploaded_files"] = []
    st.session_state["spec_sheet_extractor_upload_signature"] = None
    st.session_state["spec_sheet_extractor_file_results"] = []
    st.session_state["spec_sheet_extractor_page_inventory"] = []
    _clear_header_extraction_state()


def _clear_header_extraction_state() -> None:
    st.session_state["spec_sheet_extractor_header_results"] = []
    st.session_state["spec_sheet_extractor_header_summary"] = {}
    st.session_state["spec_sheet_extractor_extraction_signature"] = None


def _upload_signature(uploaded_files: list[object]) -> tuple[tuple[object, ...], ...]:
    return tuple(
        (
            getattr(uploaded_file, "file_id", None),
            getattr(uploaded_file, "name", None),
            getattr(uploaded_file, "size", None),
        )
        for uploaded_file in uploaded_files
    )


def _initialize_spec_sheet_extractor_state() -> None:
    st.session_state.setdefault("spec_sheet_extractor_uploaded_files", [])
    st.session_state.setdefault("spec_sheet_extractor_upload_signature", None)
    st.session_state.setdefault("spec_sheet_extractor_file_results", [])
    st.session_state.setdefault("spec_sheet_extractor_page_inventory", [])
    st.session_state.setdefault("spec_sheet_extractor_header_results", [])
    st.session_state.setdefault("spec_sheet_extractor_header_summary", {})
    st.session_state.setdefault("spec_sheet_extractor_extraction_signature", None)


def _file_summary_rows(file_results: list[PdfFileInventoryResult]) -> list[dict[str, object]]:
    return [
        {
            "Filename": result.source_filename,
            "Page Count": result.page_count,
            "Status": result.status,
            "Error": result.error_message or "",
        }
        for result in file_results
    ]


def _page_preview_rows(page_inventory: list[PdfPageInventoryRecord]) -> list[dict[str, object]]:
    return [
        {
            "Source File": record.source_filename,
            "Page Number": record.page_number,
            "Status": record.status,
        }
        for record in page_inventory
    ]


def _header_summary_rows(results: list[PdfHeaderExtractionResult]) -> list[dict[str, object]]:
    return [result.to_preview_row() for result in results]


def _header_summary(results: list[PdfHeaderExtractionResult]) -> dict[str, int]:
    return {
        "total_pages_processed": len(results),
        "pages_extracted_successfully": sum(
            1 for result in results if result.extraction_status == "Extracted"
        ),
        "pages_with_warnings": sum(
            1 for result in results if result.extraction_status == "Extracted with blanks"
        ),
        "pages_failed": sum(1 for result in results if result.extraction_status == "Failed"),
    }


def render_spec_sheet_extractor_view() -> None:
    """Render the initial placeholder page for the Spec Sheet Extractor."""
    _initialize_spec_sheet_extractor_state()

    if st.button("<- Back to Home", key="spec_sheet_extractor_back_button"):
        st.session_state["page"] = "home"
        st.stop()

    st.title("Spec Sheet Extractor")
    st.caption(
        "Upload one or more Kendal King spec-sheet PDFs and export one Excel "
        "row per PDF page."
    )
    st.info("Each PDF page will become one Excel row in the final export.")

    uploaded_files = st.file_uploader(
        "Upload spec-sheet PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        key="spec_sheet_extractor_pdf_uploads",
    )

    if not uploaded_files:
        _clear_inventory_state()
        st.info("Upload one or more PDF files to inventory pages.")
        return

    signature = _upload_signature(uploaded_files)
    if st.session_state.get("spec_sheet_extractor_upload_signature") != signature:
        file_results, page_inventory = inventory_pdf_uploads(uploaded_files)
        st.session_state["spec_sheet_extractor_uploaded_files"] = [
            uploaded_file.name for uploaded_file in uploaded_files
        ]
        st.session_state["spec_sheet_extractor_upload_signature"] = signature
        st.session_state["spec_sheet_extractor_file_results"] = [
            asdict(result) for result in file_results
        ]
        st.session_state["spec_sheet_extractor_page_inventory"] = [
            asdict(record) for record in page_inventory
        ]
        _clear_header_extraction_state()

    file_results = [
        PdfFileInventoryResult(**result)
        for result in st.session_state["spec_sheet_extractor_file_results"]
    ]
    page_inventory = [
        PdfPageInventoryRecord(**record)
        for record in st.session_state["spec_sheet_extractor_page_inventory"]
    ]

    uploaded_count = len(uploaded_files)
    valid_count = sum(1 for result in file_results if result.status == "Ready")
    failed_count = sum(1 for result in file_results if result.status == "Failed")
    total_pages = sum(result.page_count for result in file_results if result.status == "Ready")

    summary_columns = st.columns(4)
    summary_columns[0].metric("Uploaded Files", uploaded_count)
    summary_columns[1].metric("Valid PDFs", valid_count)
    summary_columns[2].metric("Total Pages", total_pages)
    summary_columns[3].metric("Failed PDFs", failed_count)

    if failed_count:
        st.warning(f"{failed_count} PDF file(s) could not be read.")

    st.subheader("File Inventory")
    st.dataframe(_file_summary_rows(file_results), use_container_width=True, hide_index=True)

    st.subheader("Page Preview")
    if page_inventory:
        st.dataframe(_page_preview_rows(page_inventory), use_container_width=True, hide_index=True)
    else:
        st.info("No readable PDF pages were found.")

    valid_file_indexes = {
        result.source_file_index for result in file_results if result.status == "Ready"
    }
    extract_disabled = not valid_file_indexes
    if st.button(
        "Extract Header Fields",
        key="spec_sheet_extractor_extract_header_fields",
        type="primary",
        disabled=extract_disabled,
    ):
        progress = st.progress(0, text="Extracting header fields...")
        with st.spinner("Extracting header fields from valid PDF pages..."):
            header_results = extract_header_fields_from_uploads(
                uploaded_files,
                source_file_indexes=valid_file_indexes,
            )
        progress.progress(100, text="Header extraction complete.")
        st.session_state["spec_sheet_extractor_header_results"] = [
            asdict(result) for result in header_results
        ]
        st.session_state["spec_sheet_extractor_header_summary"] = _header_summary(header_results)
        st.session_state["spec_sheet_extractor_extraction_signature"] = signature

    header_results = [
        PdfHeaderExtractionResult(**result)
        for result in st.session_state["spec_sheet_extractor_header_results"]
    ]
    if header_results:
        summary = st.session_state["spec_sheet_extractor_header_summary"]
        st.subheader("Header Extraction Summary")
        header_columns = st.columns(4)
        header_columns[0].metric(
            "Pages Processed",
            summary.get("total_pages_processed", 0),
        )
        header_columns[1].metric(
            "Extracted",
            summary.get("pages_extracted_successfully", 0),
        )
        header_columns[2].metric("Warnings", summary.get("pages_with_warnings", 0))
        header_columns[3].metric("Failed", summary.get("pages_failed", 0))

        st.subheader("Header Field Preview")
        st.dataframe(_header_summary_rows(header_results), use_container_width=True, hide_index=True)
