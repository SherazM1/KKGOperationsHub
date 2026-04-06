"""Streamlit entry point for the EOTF label maker application."""

from __future__ import annotations

from pathlib import Path

import streamlit as st

from app.services.excel_reader import read_excel
from app.services.excel_reader_albertsons import read_excel_albertsons
from app.services.excel_reader_sams import read_excel_sams
from app.services.pdf_generator_albertsons import generate_albertsons_pdf
from app.services.pdf_generator import generate_label_pdf
from app.services.pdf_generator_sams import generate_sams_pdf


def _apply_theme_styles() -> None:
    st.markdown(
        """
        <style>
        [data-testid="stAppViewContainer"] {
            background: #ffffff;
        }
        [data-testid="stHeader"] {
            background: rgba(0, 0, 0, 0);
        }
        .stApp, [data-testid="stMarkdownContainer"], [data-testid="stText"] {
            color: #1f2937;
        }
        .kkg-module-card {
            background: #f7f9fc;
            border: 1px solid #d6dee8;
            border-radius: 12px;
            padding: 1rem 1rem 0.35rem 1rem;
            margin-top: 0.5rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _resolve_logo_path() -> Path | None:
    candidate = Path.cwd() / "assets" / "KKG-Logo-02.png"
    if candidate.exists():
        return candidate
    return None


def render_hub_header() -> None:
    logo_col, title_col = st.columns([1, 6], vertical_alignment="center")

    logo_path = _resolve_logo_path()
    if logo_path is not None:
        with logo_col:
            st.image(str(logo_path), width=110)

    with title_col:
        st.title("Kendal King Operations Hub")
        st.caption("Internal tools for shipping, labels, and operations workflows.")


def render_mode_selector() -> str | None:
    st.subheader("Label Maker")
    st.write("Select a label workflow.")

    if "label_mode" not in st.session_state:
        st.session_state["label_mode"] = None

    left_col, middle_col, right_col = st.columns(3)

    with left_col:
        if st.button("Walmart Labels", use_container_width=True):
            st.session_state["label_mode"] = "walmart"

    with middle_col:
        if st.button("Sam's Warehouse Labels", use_container_width=True):
            st.session_state["label_mode"] = "sams"

    with right_col:
        if st.button("Albertsons Carton Labels", use_container_width=True):
            st.session_state["label_mode"] = "albertsons"

    return st.session_state["label_mode"]


def render_walmart_mode() -> None:
    try:
        st.write("Upload Excel workbook to generate EOTF labels.")

        uploaded_file = st.file_uploader(
            "Upload Excel input",
            type=["xlsx", "xlsm", "xls"],
            help="Required columns: Supplier, Store, PO, Description, SAP",
            key="walmart_file_uploader",
        )

        if uploaded_file is None:
            st.info("Upload an Excel file to begin.")
            return

        labels = read_excel(uploaded_file)
        page_count = len(labels)
        st.success(f"Parsed {len(labels)} rows. This will generate {page_count} pages.")

        if st.button("Generate Walmart PDF", type="primary", key="generate_walmart_pdf"):
            pdf_bytes = generate_label_pdf(labels)
            st.download_button(
                label="Download Walmart Labels PDF",
                data=pdf_bytes,
                file_name="walmart_labels.pdf",
                mime="application/pdf",
                key="download_walmart_pdf",
            )

    except ValueError as exc:
        st.error(f"Validation error: {exc}")
    except Exception as exc:
        st.error(f"Unexpected error: {exc}")


def render_sams_mode() -> None:
    try:
        st.write("Upload Excel workbook to generate Sam's warehouse labels.")

        uploaded_file = st.file_uploader(
            "Upload Excel input",
            type=["xlsx", "xlsm", "xls"],
            key="sams_file_uploader",
        )

        if uploaded_file is None:
            st.info("Upload an Excel file to begin.")
            return

        labels = read_excel_sams(uploaded_file)
        page_count = len(labels) * 2
        st.success(f"Parsed {len(labels)} rows. This will generate {page_count} pages.")

        if st.button("Generate Sam's PDF", type="primary", key="generate_sams_pdf"):
            pdf_bytes = generate_sams_pdf(labels)
            st.download_button(
                label="Download Sam's Labels PDF",
                data=pdf_bytes,
                file_name="sams_labels.pdf",
                mime="application/pdf",
                key="download_sams_pdf",
            )

    except ValueError as exc:
        st.error(f"Validation error: {exc}")
    except Exception as exc:
        st.error(f"Unexpected error: {exc}")


def render_albertsons_mode() -> None:
    try:
        st.title("Albertsons Carton Label Generator")
        st.write("Upload Excel to generate carton labels.")

        uploaded_file = st.file_uploader(
            "Upload Excel input",
            type=["xlsx", "xlsm", "xls"],
            key="albertsons_file_uploader",
        )

        if uploaded_file is None:
            st.info("Upload an Excel file to begin.")
            return

        labels = read_excel_albertsons(uploaded_file)
        st.success(f"Parsed {len(labels)} label rows.")

        if st.button("Generate Albertsons PDF", type="primary", key="generate_albertsons_pdf"):
            pdf_bytes = generate_albertsons_pdf(labels)
            st.download_button(
                label="Download Albertsons Labels",
                data=pdf_bytes,
                file_name="albertsons_labels.pdf",
                mime="application/pdf",
                key="download_albertsons_pdf",
            )

    except ValueError as exc:
        st.error(f"Validation error: {exc}")
    except Exception as exc:
        st.error(f"Unexpected error: {exc}")


def render_home() -> None:
    logo_path = Path.cwd() / "assets" / "KKG-Logo-02.png"

    col1, col2 = st.columns([1, 5])

    with col1:
        if logo_path.exists():
            st.image(str(logo_path), width=100)
        else:
            st.warning(f"Logo not found at {logo_path}")

    with col2:
        st.title("Kendal King Operations Hub")
        st.caption("Internal tools for shipping, labels, and operations workflows.")

    st.markdown("---")

    st.subheader("Tools")

    if st.button("Label Maker", use_container_width=True):
        st.session_state["page"] = "label_maker"

    st.markdown("---")

    st.subheader("Coming Soon")

    st.button("Shipping Tools (Coming Soon)", disabled=True, use_container_width=True)
    st.button("Inventory (Coming Soon)", disabled=True, use_container_width=True)


def render_label_maker() -> None:
    if st.button("← Back to Home"):
        st.session_state["page"] = "home"
        st.stop()

    _apply_theme_styles()
    render_hub_header()
    st.markdown('<div class="kkg-module-card">', unsafe_allow_html=True)
    render_mode_selector()
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")

    if st.session_state["label_mode"] == "walmart":
        render_walmart_mode()
    elif st.session_state["label_mode"] == "sams":
        render_sams_mode()
    elif st.session_state["label_mode"] == "albertsons":
        render_albertsons_mode()
    else:
        st.info("Select a label mode to begin.")


def main() -> None:
    """Run the Streamlit user interface."""
    st.set_page_config(page_title="Kendal King Operations Hub", layout="centered")

    if "page" not in st.session_state:
        st.session_state["page"] = "home"

    if st.session_state["page"] == "home":
        render_home()
    elif st.session_state["page"] == "label_maker":
        render_label_maker()


if __name__ == "__main__":
    main()
