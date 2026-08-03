"""Streamlit UI for the Spec Sheet Extractor scaffold."""

from __future__ import annotations

import streamlit as st


def render_spec_sheet_extractor_view() -> None:
    """Render the initial placeholder page for the Spec Sheet Extractor."""
    if st.button("<- Back to Home", key="spec_sheet_extractor_back_button"):
        st.session_state["page"] = "home"
        st.stop()

    st.title("Spec Sheet Extractor")
    st.caption(
        "Upload one or more Kendal King spec-sheet PDFs and export one Excel "
        "row per PDF page."
    )

    st.file_uploader(
        "Upload spec-sheet PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        disabled=True,
        key="spec_sheet_extractor_pdf_uploader",
    )
    st.info("PDF extraction and Excel export will be added in a future step.")
