"""Streamlit UI for the item-based Truck Inventory planner."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from app.services.truck_inventory_export import (
    export_load_summary_csv,
    export_normalized_data_csv,
    export_required_columns_csv,
    export_truck_assignments_csv,
    export_truck_boxes_csv,
)
from app.services.truck_inventory_load_summary import get_load_summary
from app.services.truck_inventory_normalizer import normalize_rows
from app.services.truck_inventory_parser import parse_combined_load_sheet, parse_excel_file
from app.services.truck_inventory_truck_assigner import assign_to_trucks, get_truck_summary_stats
from app.services.truck_inventory_validator import get_validation_summary, validate_records
from app.services.truck_inventory_visualizer import render_truck_visualization
from app.utils.truck_presets import TRUCK_PRESETS


def _initialize_truck_state() -> None:
    """Initialize session state variables for truck inventory."""
    defaults = {
        "truck_pure_file": None,
        "truck_cdw_file": None,
        "truck_combined_file": None,
        "truck_raw_records": [],
        "truck_normalized_records": [],
        "truck_validated_records": [],
        "truck_trucks": [],
        "truck_load_summary": {},
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def render_inputs_tab() -> None:
    """Render file upload and planning configuration controls."""
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
    st.subheader("Configuration")

    col1, col2 = st.columns(2)
    with col1:
        grouping_rule = st.selectbox(
            "Group By",
            options=["kkg_load_number", "po_number", "none"],
            format_func=lambda x: {
                "kkg_load_number": "KKG Load #",
                "po_number": "Retailer PO #",
                "none": "Each Item Row",
            }.get(x, x),
            key="truck_grouping_rule",
        )

    with col2:
        st.info("MVP placement uses one visual box per item unit and one layer high.")

    truck_preset = st.selectbox(
        "Truck Preset",
        options=list(TRUCK_PRESETS.keys()),
        format_func=lambda x: TRUCK_PRESETS[x].name,
        key="truck_preset",
    )
    preset = TRUCK_PRESETS[truck_preset]
    st.caption(
        f"Preset capacity: {preset.length_ft} ft x {preset.width_ft} ft x "
        f"{preset.height_ft} ft, {preset.max_weight_lbs:,.0f} lb max. "
        "Source truck dimensions override this when provided."
    )

    st.divider()

    if st.button("Process Files", use_container_width=True, key="truck_process_btn"):
        with st.spinner("Processing files..."):
            _process_truck_files(grouping_rule, truck_preset)


def _process_truck_files(grouping_rule: str, truck_preset: str) -> None:
    """Parse, normalize, validate, assign, and summarize uploaded files."""
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

    st.session_state.truck_raw_records = records
    validated, val_result = validate_records(records)
    trucks = assign_to_trucks(validated, preset_key=truck_preset, grouping_rule=grouping_rule)

    st.session_state.truck_validated_records = validated
    st.session_state.truck_normalized_records = validated
    st.session_state.truck_trucks = trucks
    st.session_state.truck_load_summary = get_load_summary(validated)

    st.info(f"Validation: {get_validation_summary(val_result)}")
    st.success("Processing complete. Review results in the other tabs.")


def render_normalized_data_tab() -> None:
    """Render normalized rows and validation status."""
    st.subheader("Normalized Data Preview")

    if not st.session_state.truck_normalized_records:
        st.info("No data loaded. Upload and process files in the Inputs tab first.")
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

    st.dataframe(df, use_container_width=True, height=400)

    st.download_button(
        label="Download Normalized Data (CSV)",
        data=export_normalized_data_csv(records),
        file_name="truck_inventory_normalized.csv",
        mime="text/csv",
    )


def render_load_summary_tab() -> None:
    """Render load-level summary grouped by KKG Load #."""
    st.subheader("Load Summary by KKG Load #")

    if not st.session_state.truck_load_summary:
        st.info("No load summary available. Process files first.")
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

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Total KKG Loads", len(summary))
    with col2:
        st.metric("Total Qty", f"{sum(s['total_qty'] for s in summary.values()):.0f}")

    st.download_button(
        label="Download Load Summary (CSV)",
        data=export_load_summary_csv(summary),
        file_name="truck_inventory_load_summary.csv",
        mime="text/csv",
    )


def render_truck_builder_tab() -> None:
    """Render item-based truck plans and assignment details."""
    st.subheader("Truck Builder & Assignments")

    if not st.session_state.truck_trucks:
        st.info("No truck plans available. Process files and configure inputs first.")
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

    st.divider()

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
                f"{truck.truck_height:.0f}; max weight {truck.truck_max_weight:,.0f} lb"
            )

            if truck.validation_notes:
                st.warning("; ".join(truck.validation_notes))

            item_rows = {}
            for box in truck.boxes:
                key = (box.retailer_po_number, box.item_number, box.row_qty)
                item_rows[key] = item_rows.get(key, 0) + 1

            st.markdown("**Item Rows:**")
            if item_rows:
                for (po_number, item_number, row_qty), placed_qty in sorted(item_rows.items()):
                    st.caption(f"- PO {po_number} | Item {item_number} | Qty {row_qty} | Placed {placed_qty}")
            else:
                st.caption("(No placed item boxes)")

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="Truck Summary (CSV)",
            data=export_truck_assignments_csv(trucks),
            file_name="truck_inventory_assignments.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with col2:
        st.download_button(
            label="Item Box Details (CSV)",
            data=export_truck_boxes_csv(trucks),
            file_name="truck_inventory_boxes.csv",
            mime="text/csv",
            use_container_width=True,
        )


def render_visualization_tab() -> None:
    """Render truck item layouts and color legend."""
    st.subheader("Truck Visualization")

    if not st.session_state.truck_trucks:
        st.info("No trucks to visualize. Process files first.")
        return

    render_truck_visualization(st.session_state.truck_trucks)


def render_export_tab() -> None:
    """Render all CSV export options."""
    st.subheader("Export Data")
    st.markdown("Download normalized inputs, required operations columns, load summaries, and truck plan details.")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Normalized & Validation**")
        if st.session_state.truck_normalized_records:
            st.download_button(
                label="Normalized Data",
                data=export_normalized_data_csv(st.session_state.truck_normalized_records),
                file_name="truck_inventory_normalized.csv",
                mime="text/csv",
                use_container_width=True,
            )
        else:
            st.button("Normalized Data (disabled)", disabled=True, use_container_width=True)

    with col2:
        st.markdown("**Required Columns**")
        if st.session_state.truck_normalized_records:
            st.download_button(
                label="Required Export",
                data=export_required_columns_csv(st.session_state.truck_normalized_records),
                file_name="truck_inventory_required_export.csv",
                mime="text/csv",
                use_container_width=True,
            )
        else:
            st.button("Required Export (disabled)", disabled=True, use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Load Summary**")
        if st.session_state.truck_load_summary:
            st.download_button(
                label="Load Summary",
                data=export_load_summary_csv(st.session_state.truck_load_summary),
                file_name="truck_inventory_load_summary.csv",
                mime="text/csv",
                use_container_width=True,
            )
        else:
            st.button("Load Summary (disabled)", disabled=True, use_container_width=True)

    with col2:
        st.markdown("**Truck Plan & Boxes**")
        if st.session_state.truck_trucks:
            st.download_button(
                label="Truck Summary",
                data=export_truck_assignments_csv(st.session_state.truck_trucks),
                file_name="truck_inventory_assignments.csv",
                mime="text/csv",
                use_container_width=True,
            )
            st.download_button(
                label="Item Box Details",
                data=export_truck_boxes_csv(st.session_state.truck_trucks),
                file_name="truck_inventory_boxes.csv",
                mime="text/csv",
                use_container_width=True,
            )
        else:
            st.button("Truck Plan (disabled)", disabled=True, use_container_width=True)


def render_truck_inventory_view() -> None:
    """Main render function for Truck Inventory module."""
    if st.button("Back to Home"):
        st.session_state["page"] = "home"
        st.stop()

    _initialize_truck_state()

    st.markdown("### Truck Inventory Module")
    st.markdown("Plan item-based truck loads from Excel input grouped by KKG Load #.")
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
