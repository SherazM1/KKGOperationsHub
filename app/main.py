"""Streamlit entry point for the EOTF label maker application."""

from __future__ import annotations

import streamlit as st

from app.services.excel_reader import read_excel
from app.services.pdf_generator import generate_label_pdf


def main() -> None:
    """Run the Streamlit user interface."""
    st.set_page_config(page_title="Kendal King Label Maker", layout="centered")

    st.title("Kendal King Shipping Label Maker")
    st.write(
        "Upload an Excel workbook to generate print-ready, letter-sized barcodes and labels."
    )

    uploaded_file = st.file_uploader(
        "Upload Excel input",
        type=["xlsx", "xlsm", "xls"],
        help="Required columns: Supplier, Store, PO, Description, SAP",
    )

    if uploaded_file is None:
        st.info("Upload an Excel file to begin.")
        return

    st.success(f"Loaded file: {uploaded_file.name}")

    if st.button("Generate PDF", type="primary"):
        try:
            labels = read_excel(uploaded_file)
            st.success(f"Parsed {len(labels)} label rows.")

            pdf_bytes = generate_label_pdf(labels)

        except ValueError as exc:
            st.error(f"Validation error: {exc}")
            return

        except Exception as exc:
            st.error(f"Unexpected error: {exc}")
            return

        st.download_button(
            label="Download Labels PDF",
            data=pdf_bytes,
            file_name="kkg_labels.pdf",
            mime="application/pdf",
        )


if __name__ == "__main__":
    main()
