"""Streamlit UI for the item-based Truck Inventory planner."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from app.services.truck_inventory_export import (
    export_required_columns_excel,
)
from app.services.truck_inventory_item_setup import (
    apply_item_setup,
    build_default_item_setup,
    merge_item_setup,
    preset_to_setup_values,
    validate_item_setup,
)
from app.services.truck_inventory_load_summary import get_load_summary
from app.services.truck_inventory_normalizer import normalize_rows
from app.services.truck_inventory_parser import parse_combined_load_sheet, parse_excel_file
from app.services.truck_inventory_truck_assigner import assign_to_trucks, get_truck_summary_stats
from app.services.truck_inventory_validator import get_validation_summary, validate_records
from app.services.truck_inventory_visualizer import render_truck_visualization
from app.utils.truck_presets import ITEM_PRESETS, TRUCK_PRESETS


def _initialize_truck_state() -> None:
    """Initialize session state variables for truck inventory."""
    defaults = {
        "truck_pure_file": None,
        "truck_cdw_file": None,
        "truck_combined_file": None,
        "truck_raw_records": [],
        "truck_normalized_records": [],
        "truck_validated_records": [],
        "truck_item_setup": [],
        "truck_trucks": [],
        "truck_load_summary": {},
        "truck_last_validation_summary": "",
        "truck_build_message": "",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def render_inputs_tab() -> None:
    """Render file upload and parsing controls."""
    st.subheader("File Uploads")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**PURE File** (optional)")
        pure_file = st.file_uploader(
            "Upload PURE file",
            type=["xlsx", "xlsm", "xls"],
            key="pure_uploader",
            label_visibility="collapsed",
        )
        if pure_file:
            st.session_state.truck_pure_file = pure_file

    with col2:
        st.markdown("**CDW File** (optional)")
        cdw_file = st.file_uploader(
            "Upload CDW file",
            type=["xlsx", "xlsm", "xls"],
            key="cdw_uploader",
            label_visibility="collapsed",
        )
        if cdw_file:
            st.session_state.truck_cdw_file = cdw_file

    st.divider()

    st.markdown("**Combined Load Sheet** (optional)")
    combined_file = st.file_uploader(
        "Upload combined load sheet",
        type=["xlsx", "xlsm", "xls"],
        key="combined_uploader",
        label_visibility="collapsed",
    )
    if combined_file:
        st.session_state.truck_combined_file = combined_file

    st.divider()

    has_upload = any([
        st.session_state.truck_pure_file,
        st.session_state.truck_cdw_file,
        st.session_state.truck_combined_file,
    ])
    if not has_upload:
        st.info("Upload at least one Excel file to start the Truck Inventory workflow.")

    if st.button("Parse Uploaded Data", use_container_width=True, key="truck_process_btn", disabled=not has_upload):
        with st.spinner("Parsing input..."):
            _process_truck_files()

    if st.session_state.truck_last_validation_summary:
        st.info(st.session_state.truck_last_validation_summary)

    if st.session_state.truck_normalized_records:
        st.success("Input parsed. Review rows, complete Item Setup, then build the truck plan.")


def _process_truck_files() -> None:
    """Parse uploaded files and prepare item setup rows keyed by Item #."""
    records = []

    if st.session_state.truck_pure_file:
        result = parse_excel_file(st.session_state.truck_pure_file)
        if result.success:
            pure_records = normalize_rows(result.rows, st.session_state.truck_pure_file.name, "pure")
            records.extend(pure_records)
            st.toast(f"Loaded {len(pure_records)} records from PURE file")
        else:
            st.error(f"Failed to parse PURE file: {result.error_message}")

    if st.session_state.truck_cdw_file:
        result = parse_excel_file(st.session_state.truck_cdw_file)
        if result.success:
            cdw_records = normalize_rows(result.rows, st.session_state.truck_cdw_file.name, "cdw")
            records.extend(cdw_records)
            st.toast(f"Loaded {len(cdw_records)} records from CDW file")
        else:
            st.error(f"Failed to parse CDW file: {result.error_message}")

    if st.session_state.truck_combined_file:
        result = parse_combined_load_sheet(st.session_state.truck_combined_file)
        if result.success:
            combined_records = normalize_rows(
                result.rows,
                st.session_state.truck_combined_file.name,
                "combined_load_sheet",
            )
            records.extend(combined_records)
            st.toast(f"Loaded {len(combined_records)} records from combined sheet")
        else:
            st.error(f"Failed to parse combined sheet: {result.error_message}")

    if not records:
        st.error("No data could be loaded from uploaded files.")
        return

    validated, val_result = validate_records(records)
    st.session_state.truck_raw_records = records
    st.session_state.truck_validated_records = validated
    st.session_state.truck_normalized_records = validated
    st.session_state.truck_item_setup = merge_item_setup(st.session_state.truck_item_setup, validated)
    if not st.session_state.truck_item_setup:
        st.session_state.truck_item_setup = build_default_item_setup(validated)
    st.session_state.truck_trucks = []
    st.session_state.truck_load_summary = get_load_summary(validated)
    st.session_state.truck_last_validation_summary = f"Validation: {get_validation_summary(val_result)}"
    st.session_state.truck_build_message = ""


def render_normalized_data_tab() -> None:
    """Render parsed business rows and normalized preview."""
    st.subheader("Normalized Data Preview")

    if not st.session_state.truck_normalized_records:
        st.info("Step 1 is not complete. Upload and parse Excel input in the Inputs tab first.")
        return

    records = st.session_state.truck_normalized_records
    df = pd.DataFrame([r.to_dict() for r in records])

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Rows", len(records))
    with col2:
        st.metric("Source Types", df["source_type"].nunique())
    with col3:
        st.metric("KKG Loads", df["kkg_load_number"].nunique())
    with col4:
        st.metric("Total Qty", f"{df['qty'].sum():.0f}")

    preview_columns = [
        "kkg_load_number",
        "retailer_po_number",
        "item_number",
        "qty",
        "source_type",
        "validation_status",
        "validation_notes",
    ]
    available_columns = [column for column in preview_columns if column in df.columns]
    st.dataframe(df[available_columns], use_container_width=True, height=400)


def render_load_summary_tab() -> None:
    """Render load-level summary grouped by KKG Load #."""
    st.subheader("Load Summary by KKG Load #")

    if not st.session_state.truck_load_summary:
        st.info("Load summary will appear after Excel input is parsed.")
        return

    summary = st.session_state.truck_load_summary
    rows = [
        {
            "KKG Load #": load_number,
            "Rows": data["rows"],
            "Unique Item # Count": data["item_count"],
            "Total Qty": data["total_qty"],
            "Total Weight": data["total_weight"],
            "Validation Status": data["validation_status"],
            "Validation Notes": data["validation_notes"],
        }
        for load_number, data in sorted(summary.items())
    ]
    st.dataframe(pd.DataFrame(rows), use_container_width=True)


def render_truck_builder_tab() -> None:
    """Render item setup, truck selection, and build/evaluate controls."""
    st.subheader("Truck Builder")

    if not st.session_state.truck_normalized_records:
        st.info("Upload and parse Excel input before completing Item Setup and building a truck plan.")
        return

    _render_item_setup_section()
    st.divider()
    _render_truck_selection_and_evaluate()
    st.divider()
    _render_truck_results()


def _render_item_setup_section() -> None:
    st.markdown("**Item Setup**")
    st.caption("Item setup is keyed by Item #. Presets can fill defaults, and every value remains editable.")

    preset_names = [preset.name for preset in ITEM_PRESETS.values()]
    setup_df = pd.DataFrame(st.session_state.truck_item_setup)
    previous_setup = [dict(row) for row in st.session_state.truck_item_setup]
    edited_df = st.data_editor(
        setup_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "Item #": st.column_config.TextColumn("Item #", disabled=True),
            "Preset": st.column_config.SelectboxColumn("Preset", options=preset_names),
            "Length": st.column_config.NumberColumn("Length", min_value=0.0, step=0.1),
            "Width": st.column_config.NumberColumn("Width", min_value=0.0, step=0.1),
            "Height": st.column_config.NumberColumn("Height", min_value=0.0, step=0.1),
            "Weight": st.column_config.NumberColumn("Weight", min_value=0.0, step=0.1),
            "Is Stackable?": st.column_config.SelectboxColumn("Is Stackable?", options=["No", "Yes"]),
            "Stack Qty": st.column_config.NumberColumn("Stack Qty", min_value=1, step=1),
            "Color": st.column_config.TextColumn("Color"),
        },
        key="truck_item_setup_editor",
    )
    edited_setup = edited_df.to_dict(orient="records")
    if edited_setup != previous_setup:
        st.session_state.truck_item_setup = edited_setup
        st.session_state.truck_trucks = []
        st.session_state.truck_build_message = "Item Setup changed. Rebuild the truck plan to refresh fit results."
    else:
        st.session_state.truck_item_setup = edited_setup

    setup_issues = validate_item_setup(
        st.session_state.truck_normalized_records,
        st.session_state.truck_item_setup,
    )
    if setup_issues:
        st.warning("Complete Item Setup before building the truck plan.")
        st.markdown("\n".join(f"- {issue}" for issue in setup_issues))
    else:
        st.success("Item Setup is complete.")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Apply Selected Presets", use_container_width=True):
            st.session_state.truck_item_setup = _apply_selected_item_presets(st.session_state.truck_item_setup)
            st.rerun()
    with col2:
        st.caption("Use hex colors such as #4ECDC4, or keep the auto-assigned values.")


def _apply_selected_item_presets(setup_rows: list[dict]) -> list[dict]:
    updated_rows = []
    for row in setup_rows:
        updated = dict(row)
        preset_values = preset_to_setup_values(str(updated.get("Preset", "")))
        if preset_values:
            updated.update(preset_values)
        updated_rows.append(updated)
    return updated_rows


def _render_truck_selection_and_evaluate() -> None:
    st.markdown("**Truck Type**")
    truck_preset = st.selectbox(
        "Truck Type",
        options=list(TRUCK_PRESETS.keys()),
        format_func=lambda x: TRUCK_PRESETS[x].name,
        key="truck_preset",
    )
    preset = TRUCK_PRESETS[truck_preset]
    threshold = st.number_input(
        "Operational Weight Threshold (lbs)",
        min_value=1.0,
        max_value=float(preset.max_weight_lbs),
        value=float(preset.operational_weight_threshold_lbs),
        step=100.0,
        key="truck_operational_weight_threshold",
    )
    st.caption(
        f"{preset.length_in:g} x {preset.width_in:g} x {preset.height_in:g} inches; "
        f"legal max {preset.max_weight_lbs:,.0f} lbs"
    )

    setup_issues = validate_item_setup(
        st.session_state.truck_normalized_records,
        st.session_state.truck_item_setup,
    )
    if st.session_state.truck_build_message:
        st.info(st.session_state.truck_build_message)

    if st.button(
        "Build / Evaluate Truck Plan",
        use_container_width=True,
        disabled=bool(setup_issues),
    ):
        with st.spinner("Evaluating fit..."):
            _evaluate_truck_plan(truck_preset, threshold)


def _evaluate_truck_plan(truck_preset: str, threshold: float) -> None:
    records, setup_issues = apply_item_setup(
        st.session_state.truck_normalized_records,
        st.session_state.truck_item_setup,
    )
    if setup_issues:
        st.error("Item setup needs attention before fit can be evaluated.")
        st.markdown("\n".join(f"- {issue}" for issue in setup_issues))
        return

    trucks = assign_to_trucks(
        records,
        preset_key=truck_preset,
        grouping_rule="kkg_load_number",
        operational_weight_threshold_lbs=threshold,
    )
    st.session_state.truck_normalized_records = records
    st.session_state.truck_validated_records = records
    st.session_state.truck_trucks = trucks
    st.session_state.truck_load_summary = get_load_summary(records)
    failed_count = sum(1 for truck in trucks if truck.validation_status == "error")
    if failed_count:
        st.session_state.truck_build_message = f"Build complete: {failed_count} load(s) need attention."
    else:
        st.session_state.truck_build_message = "Build complete: all evaluated loads fit."
    st.success("Truck plan evaluated.")


def _render_truck_results() -> None:
    if not st.session_state.truck_trucks:
        st.info("No evaluated truck plan yet. Complete Item Setup, select a truck type, then build the plan.")
        return

    trucks = st.session_state.truck_trucks
    stats = get_truck_summary_stats(trucks)
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Truck Plans", stats["total_trucks"])
    with col2:
        st.metric("Used Floor Area", f"{stats['total_used_pallets']:.0f}")
    with col3:
        st.metric("Total Weight", f"{stats['total_weight']:.0f} lb")
    with col4:
        st.metric("Failed Loads", stats["failed_loads"])

    for truck in trucks:
        label = (
            f"Truck #{truck.truck_id} - {truck.truck_preset} | "
            f"KKG Load {truck.kkg_load_number} | {truck.validation_status.title()}"
        )
        with st.expander(label, expanded=truck.validation_status == "error"):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Item Qty", truck.items_count)
            with col2:
                st.metric("Weight", f"{truck.total_weight:.0f} lb")
            with col3:
                st.metric("Floor Use", truck.utilization_display)

            st.caption(
                f"Truck dimensions: {truck.truck_length:.0f} x {truck.truck_width:.0f} x "
                f"{truck.truck_height:.0f}; threshold {truck.truck_max_weight:,.0f} lb"
            )
            if truck.validation_notes:
                st.warning("Load does not meet all fit rules.")
                st.markdown("\n".join(f"- {note}" for note in truck.validation_notes))


def render_visualization_tab() -> None:
    """Render truck item layouts and color legend."""
    st.subheader("Truck Visualization")

    if not st.session_state.truck_trucks:
        st.info("Visualization appears after the truck plan is built.")
        return

    render_truck_visualization(st.session_state.truck_trucks)


def render_export_tab() -> None:
    """Render final required export."""
    st.subheader("Export Data")

    if not st.session_state.truck_normalized_records:
        st.info("No parsed rows available for export. Upload and parse Excel input first.")
        return

    if not st.session_state.truck_trucks:
        st.warning("Truck plan has not been evaluated yet. Build the plan before final export.")
        return

    st.markdown("**Required Excel Export**")
    st.caption("Export contains only KKG Load #, Retailer PO #, Item #, and Qty.")
    st.download_button(
        label="Download Required Excel",
        data=export_required_columns_excel(st.session_state.truck_normalized_records),
        file_name="truck_inventory_required_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


def render_truck_inventory_view() -> None:
    """Main render function for Truck Inventory module."""
    if st.button("Back to Home"):
        st.session_state["page"] = "home"
        st.stop()

    _initialize_truck_state()

    st.markdown("### Truck Inventory Module")
    st.markdown("Upload Excel load rows, set up item dimensions by Item #, then evaluate truck fit.")
    _render_workflow_status()
    st.divider()

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "Inputs",
        "Normalized Data",
        "Load Summary",
        "Truck Builder",
        "Visualization",
        "Export",
    ])

    with tab1:
        render_inputs_tab()
    with tab2:
        render_normalized_data_tab()
    with tab3:
        render_load_summary_tab()
    with tab4:
        render_truck_builder_tab()
    with tab5:
        render_visualization_tab()
    with tab6:
        render_export_tab()


def _render_workflow_status() -> None:
    parsed = bool(st.session_state.truck_normalized_records)
    setup_issues = (
        validate_item_setup(st.session_state.truck_normalized_records, st.session_state.truck_item_setup)
        if parsed
        else ["Input not parsed"]
    )
    setup_complete = parsed and not setup_issues
    built = bool(st.session_state.truck_trucks)
    export_ready = parsed and built

    steps = [
        ("1. Upload", parsed),
        ("2. Item Setup", setup_complete),
        ("3. Evaluate", built),
        ("4. Export", export_ready),
    ]
    cols = st.columns(len(steps))
    for col, (label, complete) in zip(cols, steps):
        with col:
            st.metric(label, "Done" if complete else "Pending")
