"""Streamlit UI for the Display Compliance module."""

from __future__ import annotations

from dataclasses import asdict
from dataclasses import replace

import streamlit as st

from app.display_compliance.baseline import create_baseline
from app.display_compliance.models import DisplayBaseline
from app.display_compliance.proposal_backends import (
    RegionProposalBackendUnavailable,
    analyze_learned_candidate_regions,
)
from app.display_compliance.segmentation import (
    DisplayComplianceSegmentationError,
    analyze_candidate_regions,
    render_annotated_preview,
)
from app.display_compliance.storage import LocalBaselineStorage, baseline_from_dict

_SELECTED_BASELINE_KEY = "display_compliance_selected_baseline_id"
_ANNOTATED_PREVIEW_BYTES_KEY = "display_compliance_annotated_preview_bytes"
_ANNOTATED_PREVIEW_BASELINE_KEY = "display_compliance_annotated_preview_baseline_id"
_DETECTION_DIAGNOSTICS_KEY = "display_compliance_detection_diagnostics"
_DETECTION_DIAGNOSTIC_IMAGES_KEY = "display_compliance_detection_diagnostic_images"
_DETECTION_DIAGNOSTIC_SAMPLE_KEY = "display_compliance_detection_diagnostic_sample"
_DETECTION_DIAGNOSTIC_BASELINE_KEY = "display_compliance_detection_diagnostic_baseline_id"
_DETECTION_METHOD_KEY = "display_compliance_detection_method"
_CLASSICAL_METHOD = "Classical CV"
_LEARNED_METHOD = "Learned Segmentation - Experimental"


def _initialize_display_compliance_state() -> None:
    st.session_state.setdefault("display_compliance_created_baseline", None)
    st.session_state.setdefault("display_compliance_status_message", None)
    st.session_state.setdefault(_SELECTED_BASELINE_KEY, None)
    st.session_state.setdefault(_ANNOTATED_PREVIEW_BYTES_KEY, None)
    st.session_state.setdefault(_ANNOTATED_PREVIEW_BASELINE_KEY, None)
    st.session_state.setdefault(_DETECTION_DIAGNOSTICS_KEY, None)
    st.session_state.setdefault(_DETECTION_DIAGNOSTIC_IMAGES_KEY, None)
    st.session_state.setdefault(_DETECTION_DIAGNOSTIC_SAMPLE_KEY, None)
    st.session_state.setdefault(_DETECTION_DIAGNOSTIC_BASELINE_KEY, None)
    st.session_state.setdefault(_DETECTION_METHOD_KEY, _CLASSICAL_METHOD)


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


def _sanitize_selected_baseline_state(baseline_options: dict[str, DisplayBaseline]) -> None:
    selected_baseline_id = st.session_state.get(_SELECTED_BASELINE_KEY)
    if selected_baseline_id is not None and selected_baseline_id not in baseline_options:
        st.session_state.pop(_SELECTED_BASELINE_KEY, None)


def _clear_stale_preview_state(active_baseline_id: str | None) -> None:
    preview_baseline_id = st.session_state.get(_ANNOTATED_PREVIEW_BASELINE_KEY)
    if preview_baseline_id is not None and preview_baseline_id != active_baseline_id:
        st.session_state.pop(_ANNOTATED_PREVIEW_BYTES_KEY, None)
        st.session_state.pop(_ANNOTATED_PREVIEW_BASELINE_KEY, None)

    diagnostics_baseline_id = st.session_state.get(_DETECTION_DIAGNOSTIC_BASELINE_KEY)
    if diagnostics_baseline_id is not None and diagnostics_baseline_id != active_baseline_id:
        st.session_state.pop(_DETECTION_DIAGNOSTICS_KEY, None)
        st.session_state.pop(_DETECTION_DIAGNOSTIC_IMAGES_KEY, None)
        st.session_state.pop(_DETECTION_DIAGNOSTIC_SAMPLE_KEY, None)
        st.session_state.pop(_DETECTION_DIAGNOSTIC_BASELINE_KEY, None)


def _format_baseline_option(
    baseline_id: str | None,
    baseline_options: dict[str, DisplayBaseline],
) -> str:
    if baseline_id is None:
        return "Choose an option"
    baseline = baseline_options.get(baseline_id)
    return baseline.name if baseline is not None else "Unavailable baseline"


def _render_detection_diagnostics(baseline_id: str) -> None:
    if st.session_state.get(_DETECTION_DIAGNOSTIC_BASELINE_KEY) != baseline_id:
        return
    diagnostics = st.session_state.get(_DETECTION_DIAGNOSTICS_KEY)
    if not diagnostics:
        return

    st.subheader("Detection Diagnostics")
    if diagnostics.get("backend") == "sam2":
        rows = [
            {"Metric": "Original width", "Value": diagnostics["original_width"]},
            {"Metric": "Original height", "Value": diagnostics["original_height"]},
            {"Metric": "Working width", "Value": diagnostics["working_width"]},
            {"Metric": "Working height", "Value": diagnostics["working_height"]},
            {"Metric": "Backend", "Value": "SAM 2"},
            {"Metric": "Device", "Value": diagnostics["device"]},
            {"Metric": "Model load seconds", "Value": diagnostics["model_load_seconds"]},
            {"Metric": "Inference seconds", "Value": diagnostics["inference_seconds"]},
            {"Metric": "Total seconds", "Value": diagnostics["total_seconds"]},
            {"Metric": "Raw learned masks", "Value": diagnostics["raw_mask_count"]},
            {"Metric": "Rejected too small", "Value": diagnostics["rejected_too_small"]},
            {"Metric": "Rejected too large", "Value": diagnostics["rejected_too_large"]},
            {"Metric": "Rejected thin/degenerate", "Value": diagnostics["rejected_degenerate"]},
            {"Metric": "Rejected aspect ratio", "Value": diagnostics["rejected_aspect_ratio"]},
            {"Metric": "Rejected low confidence", "Value": diagnostics["rejected_low_confidence"]},
            {"Metric": "Rejected low solidity", "Value": diagnostics["rejected_low_solidity"]},
            {
                "Metric": "After basic filtering",
                "Value": diagnostics["proposals_after_basic_filtering"],
            },
            {
                "Metric": "Removed by IoU dedup",
                "Value": diagnostics["removed_by_iou_deduplication"],
            },
            {
                "Metric": "Removed as nested duplicates",
                "Value": diagnostics["removed_by_nested_deduplication"],
            },
            {
                "Metric": "After duplicate cleanup",
                "Value": diagnostics["proposals_after_duplicate_cleanup"],
            },
            {"Metric": "Repeated-size clusters", "Value": diagnostics["size_cluster_count"]},
            {
                "Metric": "Repeated-size proposals",
                "Value": diagnostics["repeated_size_member_count"],
            },
            {
                "Metric": "Alignment-supported proposals",
                "Value": diagnostics["alignment_supported_count"],
            },
            {"Metric": "Final candidate regions", "Value": diagnostics["final_region_count"]},
        ]
    else:
        rows = [
            {"Metric": "Original width", "Value": diagnostics["original_width"]},
            {"Metric": "Original height", "Value": diagnostics["original_height"]},
            {"Metric": "Working width", "Value": diagnostics["working_width"]},
            {"Metric": "Working height", "Value": diagnostics["working_height"]},
            {"Metric": "Strategy A raw proposals", "Value": diagnostics["strategy_a_raw_proposal_count"]},
            {"Metric": "Strategy A after filtering", "Value": diagnostics["strategy_a_proposals_after_geometry_filter"]},
            {"Metric": "Strategy B raw proposals", "Value": diagnostics["strategy_b_raw_proposal_count"]},
            {"Metric": "Strategy B after filtering", "Value": diagnostics["strategy_b_proposals_after_geometry_filter"]},
            {"Metric": "Strategy C raw proposals", "Value": diagnostics["strategy_c_raw_proposal_count"]},
            {"Metric": "Strategy C after filtering", "Value": diagnostics["strategy_c_proposals_after_geometry_filter"]},
            {"Metric": "Merged pool before dedup", "Value": diagnostics["merged_pool_count_before_dedup"]},
            {
                "Metric": "Removed by IoU dedup",
                "Value": diagnostics["removed_by_iou_deduplication"],
            },
            {
                "Metric": "Removed by coverage dedup",
                "Value": diagnostics["removed_by_coverage_deduplication"],
            },
            {
                "Metric": "Removed during dedup",
                "Value": diagnostics["removed_by_deduplication"],
            },
            {"Metric": "Repeated-size clusters", "Value": diagnostics["size_cluster_count"]},
            {
                "Metric": "Repeated-size proposals",
                "Value": diagnostics["repeated_size_member_count"],
            },
            {
                "Metric": "Alignment-boosted proposals",
                "Value": diagnostics["alignment_boosted_count"],
            },
            {
                "Metric": "Multi-strategy proposals",
                "Value": diagnostics["multi_strategy_supported_count"],
            },
            {"Metric": "Final candidate regions", "Value": diagnostics["final_region_count"]},
        ]
    st.dataframe(
        rows,
        use_container_width=True,
        hide_index=True,
    )

    diagnostic_images = st.session_state.get(_DETECTION_DIAGNOSTIC_IMAGES_KEY) or {}
    proposal_sample = st.session_state.get(_DETECTION_DIAGNOSTIC_SAMPLE_KEY) or []
    if not diagnostic_images and not proposal_sample:
        return

    with st.expander("View Detection Diagnostics"):
        for key, label in [
            ("normalized", "Normalized"),
            ("edges", "Edge Map"),
            ("threshold", "Adaptive Threshold"),
            ("morphology", "Morphology"),
            ("strategy_a_proposals", "Strategy A: Morphology Contours"),
            ("strategy_b_proposals", "Strategy B: Cleaned Edges"),
            ("strategy_c_proposals", "Strategy C: Structural Rectangles"),
            ("merged_proposals", "Merged Proposals Before Dedup"),
            ("final_proposals", "Final Proposals"),
            ("learned_raw_masks", "Learned: Raw Masks"),
            ("learned_raw_bboxes", "Learned: Raw Bounding Boxes"),
            ("learned_after_basic_filtering", "Learned: After Basic Filtering"),
            ("learned_after_duplicate_cleanup", "Learned: After Duplicate Cleanup"),
            ("learned_structurally_supported", "Learned: Structurally Supported"),
            ("learned_final_regions", "Learned: Final Regions"),
        ]:
            image_bytes = diagnostic_images.get(key)
            if image_bytes:
                st.image(image_bytes, caption=label, use_container_width=True)
        if proposal_sample:
            st.dataframe(proposal_sample, use_container_width=True, hide_index=True)


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
    detection_method = st.selectbox(
        "Region Proposal Method",
        options=[_CLASSICAL_METHOD, _LEARNED_METHOD],
        key=_DETECTION_METHOD_KEY,
    )
    if detection_method == _LEARNED_METHOD:
        st.caption(
            "Learned segmentation is experimental. The first run may download and cache "
            "a small local model before analyzing the image."
        )

    if st.button(
        "Detect Regions for Different Product Placements",
        type="primary",
        key="display_compliance_detect_product_regions",
        disabled=baseline is None,
    ):
        if baseline is None:
            st.info("Select a saved baseline to detect candidate product regions.")
            return
        try:
            image_bytes = storage.load_reference_image_bytes(baseline)
            if detection_method == _LEARNED_METHOD:
                with st.spinner(
                    "Preparing learned vision model, downloading the model if needed, "
                    "and analyzing the baseline image..."
                ):
                    diagnostic_result = analyze_learned_candidate_regions(image_bytes=image_bytes)
            else:
                diagnostic_result = analyze_candidate_regions(image_bytes=image_bytes)
            regions = diagnostic_result.regions
            updated_baseline = replace(baseline, regions=regions)
            storage.save_baseline_metadata(updated_baseline)
            st.session_state["display_compliance_created_baseline"] = asdict(updated_baseline)
            st.session_state[_ANNOTATED_PREVIEW_BYTES_KEY] = (
                render_annotated_preview(image_bytes=image_bytes, regions=regions)
                if regions
                else image_bytes
            )
            st.session_state[_ANNOTATED_PREVIEW_BASELINE_KEY] = updated_baseline.baseline_id
            st.session_state[_DETECTION_DIAGNOSTICS_KEY] = asdict(
                diagnostic_result.diagnostics
            )
            st.session_state[_DETECTION_DIAGNOSTIC_IMAGES_KEY] = (
                diagnostic_result.diagnostic_images
            )
            proposal_sample = getattr(diagnostic_result, "proposal_sample", [])
            st.session_state[_DETECTION_DIAGNOSTIC_SAMPLE_KEY] = [
                asdict(row) for row in proposal_sample
            ]
            st.session_state[_DETECTION_DIAGNOSTIC_BASELINE_KEY] = (
                updated_baseline.baseline_id
            )
            st.session_state["display_compliance_status_message"] = (
                f"Detected {len(regions)} candidate product region(s). "
                "Rerun detection any time to replace these candidates."
            )
        except (
            DisplayComplianceSegmentationError,
            FileNotFoundError,
            RegionProposalBackendUnavailable,
            ValueError,
        ) as exc:
            if detection_method == _LEARNED_METHOD:
                st.error(
                    "Learned segmentation is temporarily unavailable. "
                    "Classical CV remains available."
                )
                with st.expander("Technical details"):
                    st.write(str(exc))
            else:
                st.error(f"Detection error: {exc}")
        except Exception as exc:
            st.error(f"Unexpected detection error: {exc}")

    try:
        current_baseline = storage.load_baseline(baseline.baseline_id)
    except (FileNotFoundError, ValueError) as exc:
        st.info(f"Selected baseline is no longer available: {exc}")
        return

    region_count = len(current_baseline.regions)
    st.metric("Candidate Region Count", region_count)
    if region_count == 0:
        st.warning("No candidate product regions are currently saved for this baseline.")

    preview_bytes = None
    if (
        st.session_state.get(_ANNOTATED_PREVIEW_BASELINE_KEY) == current_baseline.baseline_id
    ):
        preview_bytes = st.session_state.get(_ANNOTATED_PREVIEW_BYTES_KEY)
    if preview_bytes is None and current_baseline.regions:
        try:
            preview_bytes = render_annotated_preview(
                image_bytes=storage.load_reference_image_bytes(current_baseline),
                regions=current_baseline.regions,
            )
            st.session_state[_ANNOTATED_PREVIEW_BYTES_KEY] = preview_bytes
            st.session_state[_ANNOTATED_PREVIEW_BASELINE_KEY] = current_baseline.baseline_id
        except (DisplayComplianceSegmentationError, FileNotFoundError, ValueError):
            preview_bytes = None
    if preview_bytes:
        caption = (
            "Annotated candidate product regions"
            if region_count > 0
            else "No candidate regions detected. Reference image shown below."
        )
        st.image(
            preview_bytes,
            caption=caption,
            use_container_width=True,
        )
    _render_detection_diagnostics(current_baseline.baseline_id)


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
    baseline_options = {baseline.baseline_id: baseline for baseline in saved_baselines}
    _sanitize_selected_baseline_state(baseline_options)
    selected_baseline = None
    if saved_baselines:
        st.subheader("Saved Baselines")
        selected_baseline_id = st.selectbox(
            "Select saved baseline",
            options=[None, *baseline_options.keys()],
            format_func=lambda baseline_id: _format_baseline_option(
                baseline_id,
                baseline_options,
            ),
            key=_SELECTED_BASELINE_KEY,
        )
        selected_baseline = (
            baseline_options.get(selected_baseline_id)
            if selected_baseline_id is not None
            else None
        )
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
    else:
        st.info("No saved baselines yet. Create a baseline above to begin.")

    active_baseline = selected_baseline
    active_baseline_id = active_baseline.baseline_id if active_baseline is not None else None
    _clear_stale_preview_state(active_baseline_id)
    if active_baseline is not None:
        _render_candidate_detection(storage=storage, baseline=active_baseline)
    elif saved_baselines:
        st.info("Select a saved baseline to detect candidate product regions.")
