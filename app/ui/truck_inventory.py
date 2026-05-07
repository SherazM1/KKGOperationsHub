"""Streamlit UI for Truck Inventory module."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from app.models.truck_inventory_record import TruckInventoryRecord
from app.services.truck_inventory_parser import parse_excel_file, parse_combined_load_sheet
from app.services.truck_inventory_normalizer import normalize_rows
from app.services.truck_inventory_validator import validate_records, get_validation_summary
from app.services.truck_inventory_pallet_calculator import (
    PALLET_CALC_MODES,
    DEFAULT_CASES_PER_PALLET,
    calculate_pallets,
    get_pallet_summary,
)
from app.services.truck_inventory_truck_assigner import assign_to_trucks, get_truck_summary_stats
from app.services.truck_inventory_visualizer import render_truck_visualization
from app.services.truck_inventory_export import (
    export_normalized_data_csv,
    export_pallet_summary_csv,
    export_truck_assignments_csv,
    export_truck_boxes_csv,
)
from app.utils.truck_presets import TRUCK_PRESETS


def _initialize_truck_state() -> None:
    """Initialize session state variables for truck inventory."""
    if "truck_pure_file" not in st.session_state:
        st.session_state.truck_pure_file = None
    if "truck_cdw_file" not in st.session_state:
        st.session_state.truck_cdw_file = None
    if "truck_combined_file" not in st.session_state:
        st.session_state.truck_combined_file = None
    if "truck_raw_records" not in st.session_state:
        st.session_state.truck_raw_records = []
    if "truck_normalized_records" not in st.session_state:
        st.session_state.truck_normalized_records = []
    if "truck_validated_records" not in st.session_state:
        st.session_state.truck_validated_records = []
    if "truck_trucks" not in st.session_state:
        st.session_state.truck_trucks = []
    if "truck_pallet_summary" not in st.session_state:
        st.session_state.truck_pallet_summary = {}


def render_inputs_tab() -> None:
    """Render the Inputs tab with file uploaders."""
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
    
    # Configuration section
    st.subheader("Configuration")
    
    col1, col2 = st.columns(2)
    
    with col1:
        grouping_rule = st.selectbox(
            "Group By",
            options=["load_group", "po_number", "none"],
            format_func=lambda x: {
                "load_group": "Load Group",
                "po_number": "PO Number",
                "none": "None (Sequential)",
            }.get(x, x),
            key="truck_grouping_rule",
        )
    
    with col2:
        pallet_calc_mode = st.selectbox(
            "Pallet Calculation",
            options=list(PALLET_CALC_MODES.keys()),
            format_func=lambda x: PALLET_CALC_MODES.get(x, x),
            key="truck_pallet_mode",
        )
    
    col1, col2 = st.columns(2)
    
    with col1:
        truck_preset = st.selectbox(
            "Truck Preset",
            options=list(TRUCK_PRESETS.keys()),
            format_func=lambda x: TRUCK_PRESETS[x].name,
            key="truck_preset",
        )
    
    with col2:
        cases_per_pallet = st.number_input(
            "Cases Per Pallet",
            value=DEFAULT_CASES_PER_PALLET,
            min_value=1,
            step=1,
            key="truck_cases_per_pallet",
        )
    
    st.divider()
    
    # Process button
    if st.button("🔄 Process Files", use_container_width=True, key="truck_process_btn"):
        with st.spinner("Processing files..."):
            _process_truck_files(
                grouping_rule,
                pallet_calc_mode,
                truck_preset,
                cases_per_pallet,
            )


def _process_truck_files(
    grouping_rule: str,
    pallet_calc_mode: str,
    truck_preset: str,
    cases_per_pallet: int,
) -> None:
    """Process uploaded files and perform all calculations."""
    records = []
    
    # Parse PURE file if uploaded
    if st.session_state.truck_pure_file:
        result = parse_excel_file(st.session_state.truck_pure_file)
        if result.success:
            pure_records = normalize_rows(
                result.rows,
                st.session_state.truck_pure_file.name,
                "pure",
            )
            records.extend(pure_records)
            st.toast(f"✓ Loaded {len(pure_records)} records from PURE file")
        else:
            st.error(f"Failed to parse PURE file: {result.error_message}")
    
    # Parse CDW file if uploaded
    if st.session_state.truck_cdw_file:
        result = parse_excel_file(st.session_state.truck_cdw_file)
        if result.success:
            cdw_records = normalize_rows(
                result.rows,
                st.session_state.truck_cdw_file.name,
                "cdw",
            )
            records.extend(cdw_records)
            st.toast(f"✓ Loaded {len(cdw_records)} records from CDW file")
        else:
            st.error(f"Failed to parse CDW file: {result.error_message}")
    
    # Parse combined file if uploaded
    if st.session_state.truck_combined_file:
        result = parse_combined_load_sheet(st.session_state.truck_combined_file)
        if result.success:
            combined_records = normalize_rows(
                result.rows,
                st.session_state.truck_combined_file.name,
                "combined_load_sheet",
            )
            records.extend(combined_records)
            st.toast(f"✓ Loaded {len(combined_records)} records from combined sheet")
        else:
            st.error(f"Failed to parse combined sheet: {result.error_message}")
    
    if not records:
        st.error("No data could be loaded from uploaded files.")
        return
    
    st.session_state.truck_raw_records = records
    st.session_state.truck_normalized_records = records
    
    # Validate
    validated, val_result = validate_records(records)
    st.session_state.truck_validated_records = validated
    
    st.info(f"Validation: {get_validation_summary(val_result)}")
    
    # Calculate pallets
    palletted = calculate_pallets(
        validated,
        mode=pallet_calc_mode,
        cases_per_pallet=cases_per_pallet,
    )
    st.session_state.truck_normalized_records = palletted
    
    # Get pallet summary
    pallet_summary = get_pallet_summary(palletted)
    st.session_state.truck_pallet_summary = pallet_summary
    
    # Assign to trucks
    trucks = assign_to_trucks(palletted, preset_key=truck_preset, grouping_rule=grouping_rule)
    st.session_state.truck_trucks = trucks
    
    st.success("✓ Processing complete! View results in other tabs.")


def render_normalized_data_tab() -> None:
    """Render the Normalized Data tab with preview and export."""
    st.subheader("Normalized Data Preview")
    
    if not st.session_state.truck_normalized_records:
        st.info("No data loaded. Upload and process files in the Inputs tab first.")
        return
    
    records = st.session_state.truck_normalized_records
    
    # Create DataFrame
    data = [r.to_dict() for r in records]
    df = pd.DataFrame(data)
    
    # Display stats
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Records", len(records))
    with col2:
        st.metric("Source Types", df["source_type"].nunique())
    with col3:
        st.metric("Load Groups", df["load_group"].nunique())
    with col4:
        st.metric("Total Pallets", f"{df['estimated_pallets'].sum():.1f}")
    
    # Display table
    st.dataframe(df, use_container_width=True, height=400)
    
    # Export button
    csv_data = export_normalized_data_csv(records)
    st.download_button(
        label="📥 Download Normalized Data (CSV)",
        data=csv_data,
        file_name="truck_inventory_normalized.csv",
        mime="text/csv",
    )


def render_pallet_summary_tab() -> None:
    """Render the Pallet Summary tab."""
    st.subheader("Pallet Summary by Load Group")
    
    if not st.session_state.truck_pallet_summary:
        st.info("No pallet summary available. Process files first.")
        return
    
    summary = st.session_state.truck_pallet_summary
    
    # Create DataFrame
    data = [
        {
            "Load Group": group,
            "Item Count": s["items"],
            "Total Pallets": s["total_pallets"],
            "Total Quantity": s["total_qty"],
        }
        for group, s in sorted(summary.items())
    ]
    df = pd.DataFrame(data)
    
    st.dataframe(df, use_container_width=True)
    
    # Overall stats
    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Total Load Groups", len(summary))
    with col2:
        st.metric("Total Pallets", f"{sum(s['total_pallets'] for s in summary.values()):.1f}")
    
    # Export button
    csv_data = export_pallet_summary_csv(summary)
    st.download_button(
        label="📥 Download Pallet Summary (CSV)",
        data=csv_data,
        file_name="truck_inventory_pallet_summary.csv",
        mime="text/csv",
    )


def render_truck_builder_tab() -> None:
    """Render the Truck Builder tab with assignment details."""
    st.subheader("Truck Builder & Assignments")
    
    if not st.session_state.truck_trucks:
        st.info("No trucks assigned. Process files and configure in Inputs tab.")
        return
    
    trucks = st.session_state.truck_trucks
    
    # Overall stats
    stats = get_truck_summary_stats(trucks)
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Trucks", stats["total_trucks"])
    with col2:
        st.metric("Total Used", f"{stats['total_used_pallets']:.0f}")
    with col3:
        st.metric("Total Capacity", f"{stats['total_capacity']:.0f}")
    with col4:
        st.metric("Avg Utilization", f"{stats['average_utilization']:.1f}%")
    
    st.divider()
    
    # Truck assignments
    st.markdown("**Truck Assignments:**")
    for truck in trucks:
        with st.expander(
            f"**Truck #{truck.truck_id}** - {truck.truck_preset} "
            f"({truck.utilization_display} utilized)",
            expanded=False,
        ):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Used Pallets", f"{truck.total_used_pallets:.0f}")
            with col2:
                st.metric("Capacity", f"{truck.total_capacity_pallets:.0f}")
            with col3:
                st.metric("Remaining", f"{truck.remaining_capacity:.0f}")
            
            st.markdown("**Load Groups:**")
            if truck.load_groups:
                for group in truck.load_groups:
                    st.caption(f"• {group}")
            else:
                st.caption("(No load groups)")
    
    st.divider()
    
    # Export buttons
    col1, col2, col3 = st.columns(3)
    with col1:
        csv_data = export_truck_assignments_csv(trucks)
        st.download_button(
            label="📥 Truck Summary (CSV)",
            data=csv_data,
            file_name="truck_inventory_assignments.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with col2:
        csv_data = export_truck_boxes_csv(trucks)
        st.download_button(
            label="📥 Box Details (CSV)",
            data=csv_data,
            file_name="truck_inventory_boxes.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with col3:
        st.write("")  # Spacer


def render_visualization_tab() -> None:
    """Render the Visualization tab with truck layouts and legend."""
    st.subheader("Truck Visualization")
    
    if not st.session_state.truck_trucks:
        st.info("No trucks to visualize. Process files first.")
        return
    
    trucks = st.session_state.truck_trucks
    render_truck_visualization(trucks)


def render_export_tab() -> None:
    """Render the Export tab with all download options."""
    st.subheader("Export Data")
    
    st.markdown("""
    Download your truck inventory data in multiple formats for further analysis or reporting.
    """)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**Normalized & Validation**")
        if st.session_state.truck_normalized_records:
            csv_data = export_normalized_data_csv(st.session_state.truck_normalized_records)
            st.download_button(
                label="📥 Normalized Data",
                data=csv_data,
                file_name="truck_inventory_normalized.csv",
                mime="text/csv",
                use_container_width=True,
            )
        else:
            st.button("Normalized Data (disabled)", disabled=True, use_container_width=True)
    
    with col2:
        st.markdown("**Pallet Summary**")
        if st.session_state.truck_pallet_summary:
            csv_data = export_pallet_summary_csv(st.session_state.truck_pallet_summary)
            st.download_button(
                label="📥 Pallet Summary",
                data=csv_data,
                file_name="truck_inventory_pallet_summary.csv",
                mime="text/csv",
                use_container_width=True,
            )
        else:
            st.button("Pallet Summary (disabled)", disabled=True, use_container_width=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**Truck Assignments**")
        if st.session_state.truck_trucks:
            csv_data = export_truck_assignments_csv(st.session_state.truck_trucks)
            st.download_button(
                label="📥 Truck Summary",
                data=csv_data,
                file_name="truck_inventory_assignments.csv",
                mime="text/csv",
                use_container_width=True,
            )
        else:
            st.button("Truck Summary (disabled)", disabled=True, use_container_width=True)
    
    with col2:
        st.markdown("**Box Placements**")
        if st.session_state.truck_trucks:
            csv_data = export_truck_boxes_csv(st.session_state.truck_trucks)
            st.download_button(
                label="📥 Box Details",
                data=csv_data,
                file_name="truck_inventory_boxes.csv",
                mime="text/csv",
                use_container_width=True,
            )
        else:
            st.button("Box Details (disabled)", disabled=True, use_container_width=True)


def render_truck_inventory_view() -> None:
    """Main render function for Truck Inventory module."""
    if st.button("← Back to Home"):
        st.session_state["page"] = "home"
        st.stop()
    
    _initialize_truck_state()
    
    st.markdown("### Truck Inventory Module")
    st.markdown("""
    Plan and visualize truck loads from PURE and CDW order files.
    """)
    
    st.divider()
    
    # Tabs
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📥 Inputs",
        "📋 Normalized Data",
        "📦 Pallet Summary",
        "🚚 Truck Builder",
        "📊 Visualization",
        "💾 Export",
    ])
    
    with tab1:
        render_inputs_tab()
    
    with tab2:
        render_normalized_data_tab()
    
    with tab3:
        render_pallet_summary_tab()
    
    with tab4:
        render_truck_builder_tab()
    
    with tab5:
        render_visualization_tab()
    
    with tab6:
        render_export_tab()
