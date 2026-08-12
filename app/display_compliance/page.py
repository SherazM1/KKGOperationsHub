"""Streamlit UI for the Display Compliance module."""

from __future__ import annotations

from dataclasses import asdict

import streamlit as st

from app.display_compliance.baseline import create_baseline
from app.display_compliance.models import DisplayBaseline
from app.display_compliance.storage import LocalBaselineStorage, baseline_from_dict


def _initialize_display_compliance_state() -> None:
    st.session_state.setdefault("display_compliance_created_baseline", None)
    st.session_state.setdefault("display_compliance_status_message", None)


def _created_baseline() -> DisplayBaseline | None:
    payload = st.session_state.get("display_compliance_created_baseline")
    if not payload:
        return None
    return baseline_from_dict(payload)


def _baseline_metadata_rows(baseline: DisplayBaseline) -> list[dict[str, object]]:
    return [
        {"Field": "Baseline Name", "Value": baseline.name},
        {"Field": "Source Filename", "Value": baseline.reference_filename},
        {
            "Field": "Reference Dimensions",
            "Value": f"{baseline.reference_width} x {baseline.reference_height}",
        },
        {"Field": "Baseline ID", "Value": baseline.baseline_id},
        {"Field": "Detected Region Count", "Value": len(baseline.regions)},
    ]


def render_display_compliance_view() -> None:
    """Render the initial Display Compliance baseline workflow."""
    _initialize_display_compliance_state()
    storage = LocalBaselineStorage()

    if st.button("<- Back to Home", key="display_compliance_back_button"):
        st.session_state["page"] = "home"
        st.stop()

    st.title("Displays: Product Quality Control ")
    st.caption("Create a known-perfect display baseline for future product-placement QC.")

    st.subheader("Create Baseline")
    baseline_name = st.text_input(
        "Display Name",
        key="display_compliance_baseline_name",
    )
    uploaded_file = st.file_uploader(
        "Perfect Reference Image",
        type=["png", "jpg", "jpeg"],
        help=(
            "Upload the correct display image. This baseline will become the "
            "source of truth for future placement checks."
        ),
        key="display_compliance_reference_image",
    )

    image_bytes: bytes | None = None
    if uploaded_file is not None:
        image_bytes = uploaded_file.getvalue()
        st.image(image_bytes, caption="Perfect reference image", use_container_width=True)

    if st.button(
        "Create Baseline",
        type="primary",
        key="display_compliance_create_baseline",
    ):
        try:
            if uploaded_file is None or image_bytes is None:
                raise ValueError("A perfect reference image is required.")
            baseline = create_baseline(
                name=baseline_name,
                filename=uploaded_file.name,
                image_bytes=image_bytes,
            )
            storage.save_baseline(baseline, image_bytes)
            st.session_state["display_compliance_created_baseline"] = asdict(baseline)
            st.session_state["display_compliance_status_message"] = (
                "Baseline saved. Automatic product-region mapping will be added "
                "in the next vision pass."
            )
        except ValueError as exc:
            st.error(str(exc))

    if st.session_state.get("display_compliance_status_message"):
        st.success(st.session_state["display_compliance_status_message"])

    baseline = _created_baseline()
    if baseline is not None:
        st.subheader("Baseline Data")
        st.dataframe(_baseline_metadata_rows(baseline), use_container_width=True, hide_index=True)
        st.info("Automatic region detection is not active yet. Detected region count is 0.")

    saved_baselines = storage.list_baselines()
    if saved_baselines:
        st.subheader("Saved Baselines")
        st.dataframe(
            [
                {
                    "Name": saved_baseline.name,
                    "Source Filename": saved_baseline.reference_filename,
                    "Dimensions": (
                        f"{saved_baseline.reference_width} x "
                        f"{saved_baseline.reference_height}"
                    ),
                    "Regions": len(saved_baseline.regions),
                    "Baseline ID": saved_baseline.baseline_id,
                }
                for saved_baseline in saved_baselines
            ],
            use_container_width=True,
            hide_index=True,
        )
