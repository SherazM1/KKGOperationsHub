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
from app.services.truck_inventory_parser import parse_combined_load_sheet
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
        "truck_item_setup_by_load": {},
        "truck_selected_load_number": None,
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
    st.subheader("Main Input")
    st.markdown("**Combined Load Sheet**")
    st.caption("Required for truck planning. This workbook drives KKG Load # grouping, Item Setup, visualization, and export.")
    combined_file = st.file_uploader(
        "Upload combined load sheet",
        type=["xlsx", "xlsm", "xls"],
        key="combined_uploader",
        label_visibility="collapsed",
    )
    if combined_file:
        st.session_state.truck_combined_file = combined_file

    st.divider()
    st.subheader("Optional Reference Files")
    st.caption("PURE/CDW PO files are optional references only in this MVP and are not required to parse or build the truck plan.")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**PURE PO File**")
        pure_file = st.file_uploader(
            "Upload PURE PO file",
            type=["xlsx", "xlsm", "xls"],
            key="pure_uploader",
            label_visibility="collapsed",
        )
        if pure_file:
            st.session_state.truck_pure_file = pure_file

    with col2:
        st.markdown("**CDW PO File**")
        cdw_file = st.file_uploader(
            "Upload CDW PO file",
            type=["xlsx", "xlsm", "xls"],
            key="cdw_uploader",
            label_visibility="collapsed",
        )
        if cdw_file:
            st.session_state.truck_cdw_file = cdw_file

    st.divider()
    has_primary_input = bool(st.session_state.truck_combined_file)
    if not has_primary_input:
        st.info("Upload the Combined Load Sheet to start truck planning.")

    if st.button(
        "Parse Combined Load Sheet",
        use_container_width=True,
        key="truck_process_btn",
        disabled=not has_primary_input,
    ):
        with st.spinner("Parsing input..."):
            _process_truck_files()

    if st.session_state.truck_last_validation_summary:
        st.info(st.session_state.truck_last_validation_summary)

    if st.session_state.truck_normalized_records:
        st.success("Input parsed. Select a KKG Load #, complete Item Setup for that load, then build the truck plan.")


def _process_truck_files() -> None:
    """Parse uploaded files and prepare item setup rows keyed by Item #."""
    records = []

    if st.session_state.truck_combined_file:
        result = parse_combined_load_sheet(st.session_state.truck_combined_file)
        if result.success:
            combined_records = normalize_rows(
                result.rows,
                st.session_state.truck_combined_file.name,
                "combined_load_sheet",
            )
            records.extend(combined_records)
            st.toast(f"Loaded {len(combined_records)} planning rows from combined load sheet")
            if result.file_type:
                st.caption(f"Parser selected: {result.file_type}")
        else:
            st.error(f"Failed to parse combined sheet: {result.error_message}")
            return
    else:
        st.error("Combined Load Sheet is required for truck planning.")
        return

    if not records:
        st.error("No planning rows could be loaded from the Combined Load Sheet.")
        return

    validated, val_result = validate_records(records)
    st.session_state.truck_raw_records = records
    st.session_state.truck_validated_records = validated
    st.session_state.truck_normalized_records = validated
    load_numbers = _get_load_numbers(validated)
    if not load_numbers:
        st.error("No KKG Load # values were found in the Combined Load Sheet.")
        return
    if st.session_state.truck_selected_load_number not in load_numbers:
        st.session_state.truck_selected_load_number = load_numbers[0]
    _ensure_item_setup_for_selected_load()
    st.session_state.truck_trucks = []
    st.session_state.truck_load_summary = get_load_summary(validated)
    optional_refs = []
    if st.session_state.truck_pure_file:
        optional_refs.append("PURE PO reference attached")
    if st.session_state.truck_cdw_file:
        optional_refs.append("CDW PO reference attached")
    ref_note = f" ({'; '.join(optional_refs)})" if optional_refs else ""
    st.session_state.truck_last_validation_summary = f"Combined Load Sheet validation: {get_validation_summary(val_result)}{ref_note}"
    st.session_state.truck_build_message = ""


def render_normalized_data_tab() -> None:
    """Render parsed business rows and normalized preview."""
    st.subheader("Normalized Data Preview")

    if not st.session_state.truck_normalized_records:
        st.info("Step 1 is not complete. Upload and parse Excel input in the Inputs tab first.")
        return

    records = _get_selected_records()
    if not records:
        st.warning("Select a KKG Load # to review its parsed rows.")
        return
    df = pd.DataFrame([r.to_dict() for r in records])

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Rows", len(records))
    with col2:
        st.metric("Source Types", df["source_type"].nunique())
    with col3:
        st.metric("Selected Load", st.session_state.truck_selected_load_number or "-")
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
    if st.session_state.truck_selected_load_number:
        selected = st.session_state.truck_selected_load_number
        if selected in summary:
            data = summary[selected]
            st.info(
                f"Selected load {selected}: {data['rows']} row(s), "
                f"{data['item_count']} unique item(s), {data['total_qty']} total qty."
            )


def render_truck_builder_tab() -> None:
    """Render item setup, truck selection, and build/evaluate controls."""
    st.subheader("Truck Builder")

    if not st.session_state.truck_normalized_records:
        st.info("Upload and parse Excel input before completing Item Setup and building a truck plan.")
        return

    if not _get_selected_records():
        st.warning("Select a KKG Load # before configuring items.")
        return

    _render_item_setup_section()
    st.divider()
    _render_truck_selection_and_evaluate()
    st.divider()
    _render_truck_results()


def _render_item_setup_section() -> None:
    st.markdown("**Item Setup**")
    selected_load = st.session_state.truck_selected_load_number
    selected_records = _get_selected_records()
    _ensure_item_setup_for_selected_load()
    st.caption(
        f"Generated from unique Item # values in selected KKG Load # {selected_load}. "
        "Known sample items are auto-filled; unknown items stay Custom and need manual setup. "
        "Item # is parsed from the load sheet and is read-only."
    )

    preset_names = [preset.name for preset in ITEM_PRESETS.values()]
    current_setup = st.session_state.truck_item_setup_by_load.get(selected_load, [])
    setup_df = pd.DataFrame(current_setup)
    previous_setup = [dict(row) for row in current_setup]
    edited_df = st.data_editor(
        setup_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "Item #": st.column_config.TextColumn("Item #", disabled=True),
            "Preset": st.column_config.SelectboxColumn("Preset / Type", options=preset_names),
            "Length": st.column_config.NumberColumn("Length", min_value=0.0, step=0.1),
            "Width": st.column_config.NumberColumn("Width", min_value=0.0, step=0.1),
            "Height": st.column_config.NumberColumn("Height", min_value=0.0, step=0.1),
            "Weight": st.column_config.NumberColumn("Weight", min_value=0.0, step=0.1),
            "Is Stackable?": st.column_config.SelectboxColumn("Is Stackable?", options=["No", "Yes"]),
            "Stack Qty": st.column_config.NumberColumn("Stack Qty", min_value=1, step=1),
            "Color": st.column_config.TextColumn("Color"),
        },
        key=f"truck_item_setup_editor_{selected_load}",
    )
    edited_setup = edited_df.to_dict(orient="records")
    if edited_setup != previous_setup:
        st.session_state.truck_item_setup_by_load[selected_load] = edited_setup
        st.session_state.truck_item_setup = edited_setup
        st.session_state.truck_trucks = []
        st.session_state.truck_build_message = "Item Setup changed. Rebuild the truck plan to refresh fit results."
    else:
        st.session_state.truck_item_setup_by_load[selected_load] = edited_setup
        st.session_state.truck_item_setup = edited_setup

    setup_issues = validate_item_setup(
        selected_records,
        st.session_state.truck_item_setup_by_load.get(selected_load, []),
    )
    if setup_issues:
        st.warning("Complete Item Setup before building the truck plan.")
        st.markdown("\n".join(f"- {issue}" for issue in setup_issues))
    else:
        st.success("Item Setup is complete.")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Apply Selected Presets", use_container_width=True):
            updated_setup = _apply_selected_item_presets(
                st.session_state.truck_item_setup_by_load.get(selected_load, [])
            )
            st.session_state.truck_item_setup_by_load[selected_load] = updated_setup
            st.session_state.truck_item_setup = updated_setup
            st.session_state.truck_trucks = []
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


def _render_load_selection() -> None:
    load_numbers = _get_load_numbers(st.session_state.truck_normalized_records)
    if not load_numbers:
        st.warning("No KKG Load # values were found in the parsed Combined Load Sheet.")
        return

    previous_selected_load = st.session_state.truck_selected_load_number
    current_load = previous_selected_load
    if current_load not in load_numbers:
        current_load = load_numbers[0]
        st.session_state.truck_selected_load_number = current_load

    selector_key = "truck_load_selector"
    selector_value = st.session_state.get(selector_key)
    if selector_value in load_numbers:
        current_load = selector_value
        st.session_state.truck_selected_load_number = current_load
        if current_load != previous_selected_load:
            _ensure_item_setup_for_selected_load()
            st.session_state.truck_trucks = []
            st.session_state.truck_build_message = "Selected load changed. Build the truck plan for this load."
    else:
        st.session_state[selector_key] = current_load

    selected_load = st.selectbox(
        "KKG Load #",
        options=load_numbers,
        index=load_numbers.index(current_load),
        key=selector_key,
    )

    if selected_load != st.session_state.truck_selected_load_number:
        st.session_state.truck_selected_load_number = selected_load
        _ensure_item_setup_for_selected_load()
        st.session_state.truck_trucks = []
        st.session_state.truck_build_message = "Selected load changed. Build the truck plan for this load."

    selected_records = _get_selected_records()
    unique_items = len({record.item_number for record in selected_records if record.item_number})
    total_qty = sum(record.qty or 0 for record in selected_records)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Rows", len(selected_records))
    with col2:
        st.metric("Unique Items", unique_items)
    with col3:
        st.metric("Total Qty", total_qty)


def _ensure_item_setup_for_selected_load() -> None:
    selected_load = st.session_state.truck_selected_load_number
    if not selected_load:
        return

    selected_records = _get_selected_records()
    setup_by_load = st.session_state.truck_item_setup_by_load
    existing_setup = setup_by_load.get(selected_load, [])
    setup_by_load[selected_load] = (
        merge_item_setup(existing_setup, selected_records)
        if existing_setup
        else build_default_item_setup(selected_records)
    )
    st.session_state.truck_item_setup = setup_by_load[selected_load]


def _get_load_numbers(records) -> list[str]:
    return sorted({
        record.kkg_load_number
        for record in records
        if record.kkg_load_number
    })


def _get_selected_records():
    selected_load = st.session_state.get("truck_selected_load_number")
    if not selected_load:
        return []
    return [
        record for record in st.session_state.truck_normalized_records
        if record.kkg_load_number == selected_load
    ]


def _replace_selected_records(updated_records) -> None:
    selected_load = st.session_state.truck_selected_load_number
    updated_by_key = {
        (record.kkg_load_number, record.retailer_po_number, record.item_number, index): record
        for index, record in enumerate(updated_records)
    }
    updated_iter = iter(updated_by_key.values())
    merged = []
    for record in st.session_state.truck_normalized_records:
        if record.kkg_load_number == selected_load:
            merged.append(next(updated_iter))
        else:
            merged.append(record)
    st.session_state.truck_normalized_records = merged


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
        _get_selected_records(),
        st.session_state.truck_item_setup_by_load.get(st.session_state.truck_selected_load_number, []),
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
    selected_records = _get_selected_records()
    selected_setup = st.session_state.truck_item_setup_by_load.get(st.session_state.truck_selected_load_number, [])
    records, setup_issues = apply_item_setup(
        selected_records,
        selected_setup,
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
    _replace_selected_records(records)
    st.session_state.truck_validated_records = st.session_state.truck_normalized_records
    st.session_state.truck_trucks = trucks
    st.session_state.truck_load_summary = get_load_summary(st.session_state.truck_normalized_records)
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

    if not st.session_state.truck_selected_load_number:
        st.warning("Select a KKG Load # before exporting.")
        return

    if not st.session_state.truck_trucks:
        st.warning("Truck plan has not been evaluated yet. Build the plan before final export.")
        return

    st.markdown("**Required Excel Export**")
    st.caption(
        f"Export contains only selected KKG Load # {st.session_state.truck_selected_load_number} "
        "with KKG Load #, Retailer PO #, Item #, and Qty."
    )
    st.download_button(
        label="Download Required Excel",
        data=export_required_columns_excel(_get_selected_records()),
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
    if st.session_state.truck_normalized_records:
        _render_load_selection()
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
    selected_records = _get_selected_records() if parsed else []
    setup_issues = (
        validate_item_setup(
            selected_records,
            st.session_state.truck_item_setup_by_load.get(st.session_state.truck_selected_load_number, []),
        )
        if parsed and selected_records
        else ["Input not parsed"]
    )
    setup_complete = parsed and bool(selected_records) and not setup_issues
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
