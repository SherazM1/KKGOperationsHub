"""Streamlit UI for the Display Compliance module."""

from __future__ import annotations

from dataclasses import asdict
from dataclasses import replace

import streamlit as st

from app.display_compliance.baseline import create_baseline
from app.display_compliance.models import DisplayBaseline
from app.display_compliance.segmentation import (
    DisplayComplianceSegmentationError,
    detect_baseline_regions,
    render_annotated_preview,
)
from app.display_compliance.storage import LocalBaselineStorage, baseline_from_dict


def _initialize_display_compliance_state() -> None:
    st.session_state.setdefault("display_compliance_created_baseline", None)
    st.session_state.setdefault("display_compliance_status_message", None)
    st.session_state.setdefault("display_compliance_selected_baseline_id", None)
    st.session_state.setdefault("display_compliance_annotated_preview_bytes", None)


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
        {"Field": "Detected Region Count", "Value": str(len(baseline.regions))},
    ]


def _selected_or_created_baseline(
    *,
    storage: LocalBaselineStorage,
    saved_baselines: list[DisplayBaseline],
) -> DisplayBaseline | None:
    selected_baseline_id = st.session_state.get("display_compliance_selected_baseline_id")
    if selected_baseline_id:
        try:
            return storage.load_baseline(selected_baseline_id)
        except FileNotFoundError:
            st.session_state["display_compliance_selected_baseline_id"] = None

    created = _created_baseline()
    if created is not None:
        return created
    return saved_baselines[0] if saved_baselines else None


def _render_candidate_detection(
    *,
    storage: LocalBaselineStorage,
    baseline: DisplayBaseline,
) -> None:
    st.subheader("Candidate Product Regions")
    st.caption(
        "Experimental local detection proposes candidate product regions for review. "
        "These are not guaranteed final product boundaries."
    )

    if st.button(
        "Detect Product Regions",
        type="primary",
        key="display_compliance_detect_product_regions",
    ):
        try:
            image_bytes = storage.load_reference_image_bytes(baseline)
            regions = detect_baseline_regions(baseline=baseline, image_bytes=image_bytes)
            updated_baseline = replace(baseline, regions=regions)
            storage.save_baseline_metadata(updated_baseline)
            st.session_state["display_compliance_created_baseline"] = asdict(updated_baseline)
            st.session_state["display_compliance_selected_baseline_id"] = (
                updated_baseline.baseline_id
            )
            st.session_state["display_compliance_annotated_preview_bytes"] = (
                render_annotated_preview(image_bytes=image_bytes, regions=regions)
            )
            st.session_state["display_compliance_status_message"] = (
                f"Detected {len(regions)} candidate product region(s). "
                "Rerun detection any time to replace these candidates."
            )
        except (DisplayComplianceSegmentationError, FileNotFoundError, ValueError) as exc:
            st.error(f"Detection error: {exc}")
        except Exception as exc:
            st.error(f"Unexpected detection error: {exc}")

    current_baseline = storage.load_baseline(baseline.baseline_id)
    region_count = len(current_baseline.regions)
    st.metric("Candidate Region Count", region_count)
    if region_count == 0:
        st.warning("No candidate product regions are currently saved for this baseline.")

    preview_bytes = st.session_state.get("display_compliance_annotated_preview_bytes")
    if preview_bytes is None and current_baseline.regions:
        try:
            preview_bytes = render_annotated_preview(
                image_bytes=storage.load_reference_image_bytes(current_baseline),
                regions=current_baseline.regions,
            )
            st.session_state["display_compliance_annotated_preview_bytes"] = preview_bytes
        except (DisplayComplianceSegmentationError, FileNotFoundError, ValueError):
            preview_bytes = None
    if preview_bytes:
        st.image(
            preview_bytes,
            caption="Annotated candidate product regions",
            use_container_width=True,
        )


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
        baseline_options = {baseline.baseline_id: baseline for baseline in saved_baselines}
        selected_baseline_id = st.selectbox(
            "Select saved baseline",
            options=list(baseline_options.keys()),
            format_func=lambda baseline_id: baseline_options[baseline_id].name,
            key="display_compliance_selected_baseline_id",
        )
        selected_baseline = baseline_options[selected_baseline_id]
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

    active_baseline = _selected_or_created_baseline(
        storage=storage,
        saved_baselines=saved_baselines,
    )
    if active_baseline is not None:
        _render_candidate_detection(storage=storage, baseline=active_baseline)
