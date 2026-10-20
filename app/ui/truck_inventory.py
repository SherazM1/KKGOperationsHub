"""Streamlit UI for the item-based Truck Inventory planner."""

from __future__ import annotations

import pandas as pd
import streamlit as st
from app.services.truck_inventory_manual import assignment_rows, apply_manual_assignments

from app.services.truck_inventory_export import (
    export_required_columns_excel,
)
from app.services.truck_inventory_item_setup import (
    apply_item_setup,
    build_default_item_setup,
    merge_item_setup,
    normalize_editor_rows as _normalize_editor_rows,
    preset_to_setup_values,
    validate_item_setup,
)
from app.services.truck_inventory_load_summary import get_load_summary
from app.services.truck_inventory_normalizer import normalize_rows
from app.services.truck_inventory_parser import parse_combined_load_sheet
from app.services.truck_inventory_truck_assigner import assign_to_trucks, get_truck_summary_stats
from app.services.truck_inventory_validator import get_validation_summary, validate_records
from app.services.truck_inventory_visualizer import render_truck_visualization
from app.utils.truck_presets import ITEM_PRESETS, ITEM_SETUP_COLOR_OPTIONS, TRUCK_PRESETS


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
        "truck_item_setup_draft_by_load": {},
        "truck_item_setup_confirmed_by_load": {},
        "truck_selected_load_number": None,
        "truck_trucks": [],
        "truck_load_summary": {},
        "truck_last_validation_summary": "",
        "truck_build_message": "",
        "truck_item_setup_editor_revision": {},
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value
    # Keep choices when their widgets are on another workspace.
    for key in ("truck_item_selector", "truck_loads_selector", "truck_preset", "truck_operational_weight_threshold"):
        if key in st.session_state:
            st.session_state[key] = st.session_state[key]


def render_inputs_tab() -> None:
    """Render file upload and parsing controls."""
    st.markdown("**Upload your Excel file**")
    st.caption("Load sheets and store inventory are supported. Upload once, then choose DCs and items below.")
    with st.expander("Supported columns and units"):
        st.write("Store inventory: Store #, DC #, Item #, Quantity, Pallet Length, Pallet Width, Pallet Height, and Weight. Dimensions are inches; weight is pounds per pallet. Files without load numbers are grouped by DC.")
    combined_file = st.file_uploader(
        "Upload combined load sheet",
        type=["xlsx", "xlsm", "xls"],
        key="combined_uploader",
        label_visibility="collapsed",
    )
    if combined_file:
        st.session_state.truck_combined_file = combined_file

    with st.expander("Optional PO reference files"):
        st.caption("Reference attachments only; these do not change the truck plan.")
        _render_reference_uploads()

    has_primary_input = bool(st.session_state.truck_combined_file)
    if st.button(
        "Load Excel file",
        use_container_width=True,
        key="truck_process_btn",
        type="primary",
        disabled=not has_primary_input,
    ):
        with st.spinner("Reading your file..."):
            _process_truck_files()
        if st.session_state.truck_normalized_records:
            st.rerun()

    if st.session_state.truck_last_validation_summary:
        with st.expander("Import details"):
            st.write(st.session_state.truck_last_validation_summary)


def _render_reference_uploads() -> None:
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
    st.session_state.pop("truck_item_selector", None)
    st.session_state.pop("truck_item_selection_context", None)
    # A newly parsed file may change dimensions for an existing item or DC.
    # Discard setup derived from the previous parse so imported values refresh.
    for key in ("truck_item_setup_by_load", "truck_item_setup_draft_by_load", "truck_item_setup_confirmed_by_load"):
        st.session_state[key] = {}
    st.session_state.truck_item_setup_editor_revision = {
        key: value + 1 for key, value in st.session_state.truck_item_setup_editor_revision.items()
    }
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
        st.metric("Selected Loads", _selection_label() or "-")
    with col4:
        st.metric("Total Qty", f"{df['qty'].sum():.0f}")

    preview_columns = [
        "kkg_load_number",
        "store_number",
        "dc_number",
        "retailer_po_number",
        "item_number",
        "qty",
        "item_length",
        "item_width",
        "item_height",
        "item_weight",
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

    selected_load = st.session_state.truck_selected_load_number
    st.caption(
        f"Selected KKG loads: {_selection_label()}. Item Setup, evaluation, visualization, and export use these loads."
    )
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
        f"Generated from unique Item # values in selected KKG loads: {_selection_label()}. "
        "Imported dimensions and weight prefill the table. Missing values need manual setup. If an item has conflicting dimensions or weights, the affected field is zero for review. "
        "Item # is parsed from the load sheet and is read-only."
    )

    preset_names = [preset.name for preset in ITEM_PRESETS.values()]
    saved_setup = st.session_state.truck_item_setup_by_load.get(selected_load, [])
    draft_setup = st.session_state.truck_item_setup_draft_by_load.get(selected_load, [])
    setup_df = pd.DataFrame(draft_setup)
    setup_df["Double Stack"] = [str(row.get("Is Stackable?", "No")).lower() == "yes" for row in draft_setup]
    setup_df = setup_df.drop(columns=["Is Stackable?", "Stack Qty", "Color"], errors="ignore")
    editor_revision = st.session_state.truck_item_setup_editor_revision.get(selected_load, 0)
    editor_key = f"truck_item_setup_editor_{selected_load}_{editor_revision}"

    with st.form(key=f"truck_item_setup_form_{selected_load}_{editor_revision}"):
        edited_df = st.data_editor(
            setup_df,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            column_config={
                "Item #": st.column_config.TextColumn("Item #", disabled=True),
                "Preset": st.column_config.SelectboxColumn("Preset / Type", options=preset_names),
                "Length": st.column_config.NumberColumn("Length (in)", min_value=0.0, step=0.1),
                "Width": st.column_config.NumberColumn("Width (in)", min_value=0.0, step=0.1),
                "Height": st.column_config.NumberColumn("Height (in)", min_value=0.0, step=0.1),
                "Weight": st.column_config.NumberColumn("Weight (lb)", min_value=0.0, step=0.1),
                "Double Stack": st.column_config.CheckboxColumn("Double Stack", help="Stack two units. Half pallets cannot be stacked. Total stack height plus 6 inches must fit below the roof."),
            },
            key=editor_key,
        )
        col1, col2 = st.columns(2)
        with col1:
            save_setup = st.form_submit_button("Save Item Setup", use_container_width=True)
        with col2:
            apply_presets = st.form_submit_button("Apply Selected Presets", use_container_width=True)

    edited_setup = _normalize_editor_rows(edited_df.to_dict(orient="records"))
    edited_setup, _ = _apply_preset_changes(draft_setup, edited_setup)
    if apply_presets:
        updated_setup = _apply_selected_item_presets(edited_setup)
        st.session_state.truck_item_setup_draft_by_load[selected_load] = updated_setup
        st.session_state.truck_item_setup_confirmed_by_load[selected_load] = False
        st.session_state.truck_trucks = []
        st.session_state.truck_build_message = "Item Setup draft changed. Save Item Setup before rebuilding the truck plan."
        st.session_state.truck_item_setup_editor_revision[selected_load] = editor_revision + 1
        st.rerun()

    if save_setup:
        st.session_state.truck_item_setup_draft_by_load[selected_load] = edited_setup
        st.session_state.truck_item_setup_by_load[selected_load] = edited_setup
        st.session_state.truck_item_setup = edited_setup
        st.session_state.truck_item_setup_confirmed_by_load[selected_load] = True
        st.session_state.truck_trucks = []
        saved_setup = edited_setup
        setup_issues = validate_item_setup(selected_records, edited_setup)
        if setup_issues:
            st.session_state.truck_build_message = "Saved Item Setup has validation issues."
        else:
            st.session_state.truck_build_message = f"Item Setup saved for KKG loads: {_selection_label()}."
        st.session_state.truck_item_setup_editor_revision[selected_load] = editor_revision + 1
        st.rerun()

    setup_confirmed = st.session_state.truck_item_setup_confirmed_by_load.get(selected_load, False)
    saved_issues = validate_item_setup(selected_records, saved_setup) if setup_confirmed else []
    if not setup_confirmed:
        st.warning("Item Setup is not saved yet. Complete the table, then click Save Item Setup.")
    elif saved_issues:
        st.warning("Saved Item Setup needs attention before building the truck plan.")
        st.markdown("\n".join(f"- {issue}" for issue in saved_issues))
    else:
        st.success(f"Item Setup saved for KKG loads: {_selection_label()}.")
    st.caption("Double Stack places two units on top of each other within a load/PO row. Half pallets cannot be stacked. The top of a double stack must be at least 6 inches below the truck roof. Colors are assigned by item number. Save Item Setup to apply edits.")


def _apply_selected_item_presets(setup_rows: list[dict]) -> list[dict]:
    updated_rows = []
    for row in setup_rows:
        updated = dict(row)
        preset_values = preset_to_setup_values(str(updated.get("Preset", "")))
        if preset_values:
            updated.update(preset_values)
        updated_rows.append(updated)
    return updated_rows


def _apply_preset_changes(previous_rows: list[dict], edited_rows: list[dict]) -> tuple[list[dict], bool]:
    previous_preset_by_item = {
        str(row.get("Item #", "")): str(row.get("Preset", ""))
        for row in previous_rows
        if row.get("Item #")
    }
    updated_rows = []
    preset_changed = False
    for row in edited_rows:
        updated = dict(row)
        item_number = str(updated.get("Item #", ""))
        previous_preset = previous_preset_by_item.get(item_number)
        current_preset = str(updated.get("Preset", ""))
        if previous_preset is not None and current_preset != previous_preset:
            preset_values = preset_to_setup_values(current_preset)
            if preset_values:
                updated.update(preset_values)
            preset_changed = True
        updated_rows.append(updated)
    return updated_rows, preset_changed


def _render_load_selection() -> None:
    load_numbers = _get_load_numbers(st.session_state.truck_normalized_records)
    if not load_numbers:
        st.warning("No KKG Load # values were found in the parsed Combined Load Sheet.")
        return

    st.markdown("**Loads / DC groups to combine in one truck**")
    st.caption("Select one or more groups. Store inventory uses DC groups when no load number is supplied. Original store, DC, item, load, and PO numbers are preserved.")
    selector_key = "truck_loads_selector"
    if selector_key not in st.session_state:
        st.session_state[selector_key] = [load for load in _selected_load_numbers() if load in load_numbers] or [load_numbers[0]]
    else:
        st.session_state[selector_key] = [load for load in st.session_state[selector_key] if load in load_numbers]
    selected = sorted(st.multiselect("Select loads / DC groups", options=load_numbers, key=selector_key))
    selected_load = selected[0] if len(selected) == 1 else tuple(selected) if selected else None

    if selected_load != st.session_state.truck_selected_load_number:
        st.session_state.truck_selected_load_number = selected_load
        _ensure_item_setup_for_selected_load()
        st.session_state.truck_trucks = []
        st.session_state.truck_build_message = "Load selection changed. Save Item Setup and build the combined truck plan."

    available_items = sorted({r.item_number for r in st.session_state.truck_normalized_records
                              if r.kkg_load_number in selected and r.item_number})
    item_key = "truck_item_selector"
    context = tuple(selected)
    if st.session_state.get("truck_item_selection_context") != context or item_key not in st.session_state:
        st.session_state[item_key] = available_items
        st.session_state.truck_item_selection_context = context
    previous_items = st.session_state.get("truck_active_items")
    items = st.multiselect("Items to include in this load", options=available_items, key=item_key)
    st.caption("All items are selected initially. Remove items to plan only the products you need; the uploaded file stays parsed.")
    active_items = tuple(items)
    if previous_items != active_items:
        st.session_state.truck_active_items = active_items
        _ensure_item_setup_for_selected_load()
        st.session_state.truck_trucks = []
        st.session_state.truck_build_message = "Item selection changed. Save Item Setup and rebuild the plan."

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
    draft_by_load = st.session_state.truck_item_setup_draft_by_load
    confirmed_by_load = st.session_state.truck_item_setup_confirmed_by_load
    existing_setup = setup_by_load.get(selected_load, [])
    saved_setup = (
        merge_item_setup(existing_setup, selected_records)
        if existing_setup
        else build_default_item_setup(selected_records)
    )
    if saved_setup != existing_setup:
        confirmed_by_load[selected_load] = False
    setup_by_load[selected_load] = saved_setup

    existing_draft = draft_by_load.get(selected_load, [])
    draft_by_load[selected_load] = (
        merge_item_setup(existing_draft, selected_records)
        if existing_draft
        else [dict(row) for row in saved_setup]
    )
    st.session_state.truck_item_setup = setup_by_load[selected_load]


def _get_load_numbers(records) -> list[str]:
    return sorted({
        record.kkg_load_number
        for record in records
        if record.kkg_load_number
    })


def _selected_load_numbers() -> list[str]:
    selected = st.session_state.get("truck_selected_load_number")
    return list(selected) if isinstance(selected, tuple) else [selected] if selected else []


def _selection_label() -> str:
    return ", ".join(_selected_load_numbers())


def _get_selected_records():
    selected_loads = _selected_load_numbers()
    items = st.session_state.get("truck_item_selector")
    return [
        record for record in st.session_state.truck_normalized_records
        if record.kkg_load_number in selected_loads
        and (items is None or record.item_number in items)
    ]


def _replace_selected_records(updated_records) -> None:
    selected_ids = {id(record) for record in _get_selected_records()}
    updated_iter = iter(updated_records)
    merged = []
    for record in st.session_state.truck_normalized_records:
        if id(record) in selected_ids:
            merged.append(next(updated_iter))
        else:
            merged.append(record)
    st.session_state.truck_normalized_records = merged


def _invalidate_truck_plan() -> None:
    st.session_state.truck_trucks = []
    st.session_state.truck_build_message = "Truck settings changed. Build the plan to update results."


def _render_truck_selection_and_evaluate() -> None:
    st.markdown("**Truck Type**")
    truck_preset = st.selectbox(
        "Truck Type",
        options=list(TRUCK_PRESETS.keys()),
        format_func=lambda x: TRUCK_PRESETS[x].name,
        key="truck_preset",
        on_change=_invalidate_truck_plan,
    )
    preset = TRUCK_PRESETS[truck_preset]
    if st.session_state.get("truck_weight_preset") != truck_preset:
        st.session_state.truck_operational_weight_threshold = float(preset.operational_weight_threshold_lbs)
        st.session_state.truck_weight_preset = truck_preset
    threshold = st.number_input(
        "Operational Weight Threshold (lbs)",
        min_value=1.0,
        max_value=float(preset.max_weight_lbs),
        step=100.0,
        key="truck_operational_weight_threshold",
        on_change=_invalidate_truck_plan,
    )
    st.caption(
        f"{preset.length_in:g} x {preset.width_in:g} x {preset.height_in:g} inches; "
        f"legal max {preset.max_weight_lbs:,.0f} lbs"
    )
    st.caption("Additional trucks of this type are added automatically when space or weight capacity is reached. Row quantities may be split; store, DC, and PO details stay attached to each pallet.")

    selected_load = st.session_state.truck_selected_load_number
    setup_confirmed = st.session_state.truck_item_setup_confirmed_by_load.get(selected_load, False)
    setup_issues = validate_item_setup(
        _get_selected_records(),
        st.session_state.truck_item_setup_by_load.get(selected_load, []),
    )
    if st.session_state.truck_build_message:
        st.info(st.session_state.truck_build_message)
    if not setup_confirmed:
        st.warning("Save Item Setup before building the truck plan.")

    if st.button(
        "Build / Evaluate Truck Plan",
        type="primary",
        use_container_width=True,
        disabled=(not setup_confirmed or bool(setup_issues)),
    ):
        with st.spinner("Evaluating fit..."):
            _evaluate_truck_plan(truck_preset, threshold)


def _evaluate_truck_plan(truck_preset: str, threshold: float) -> None:
    selected_records = _get_selected_records()
    selected_load = st.session_state.truck_selected_load_number
    if not st.session_state.truck_item_setup_confirmed_by_load.get(selected_load, False):
        st.error("Save Item Setup before fit can be evaluated.")
        return
    selected_setup = st.session_state.truck_item_setup_by_load.get(selected_load, [])
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
        grouping_rule="combined",
        operational_weight_threshold_lbs=threshold,
    )
    _replace_selected_records(records)
    st.session_state.truck_validated_records = st.session_state.truck_normalized_records
    st.session_state.truck_trucks = trucks
    st.session_state.truck_manual_revision = st.session_state.get("truck_manual_revision", 0) + 1
    st.session_state.truck_load_summary = get_load_summary(st.session_state.truck_normalized_records)
    failed_count = sum(1 for truck in trucks if truck.validation_status == "error")
    if failed_count:
        st.session_state.truck_build_message = f"Build complete: {len(trucks) - failed_count} truck(s) planned; {failed_count} group(s) cannot fit and need review."
    else:
        st.session_state.truck_build_message = f"Build complete: all items assigned across {len(trucks)} truck(s)."
    st.success("Truck plan evaluated. Loads are split across additional trucks when needed.")


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
                st.metric("Load Capacity", truck.load_capacity_display)

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
    _render_manual_editor()


def _render_manual_editor() -> None:
    trucks = st.session_state.truck_trucks
    with st.expander("Edit truck assignments and pallet order"):
        if any(truck.validation_status != "valid" for truck in trucks):
            st.info("Resolve unassigned pallets before editing the plan.")
            return
        st.caption("Change Truck # to move a pallet or double stack. Enter a new truck number to add a truck. Change Order to adjust placement order within each truck. Double stacks move together. Save validates space, weight, and stacking; rebuilding replaces manual edits.")
        revision = st.session_state.get("truck_manual_revision", 0)
        with st.form(f"truck_manual_form_{revision}"):
            edited = st.data_editor(
                pd.DataFrame(assignment_rows(trucks)), hide_index=True,
                use_container_width=True, num_rows="fixed",
                disabled=["Pallet ID", "Item #", "Qty", "Store #", "DC #", "Load", "PO"],
                column_config={
                    "Truck #": st.column_config.NumberColumn("Truck #", min_value=1, step=1, required=True),
                    "Order": st.column_config.NumberColumn("Order", min_value=1, step=1, required=True),
                }, key=f"truck_manual_editor_{revision}")
            save = st.form_submit_button("Save Manual Truck Plan")
        if save:
            try:
                updated = apply_manual_assignments(trucks, edited.to_dict(orient="records"))
            except ValueError as exc:
                st.error(str(exc))
                st.caption("The previous plan remains active. Correct the assignments and save again.")
            else:
                st.session_state.truck_trucks = updated
                st.session_state.truck_manual_revision = revision + 1
                st.session_state.truck_build_message = "Manual truck plan saved. Visualization and export use your assignments."
                st.rerun()


def render_export_tab() -> None:
    """Render final required export."""
    st.subheader("Your truck plan")

    if not st.session_state.truck_normalized_records:
        st.info("No parsed rows available for export. Upload and parse Excel input first.")
        return

    if not st.session_state.truck_selected_load_number:
        st.warning("Select a KKG Load # before exporting.")
        return

    if not st.session_state.truck_trucks:
        st.warning("Truck plan has not been evaluated yet. Build the plan before final export.")
        return

    st.caption(
        f"Export contains selected KKG loads: {_selection_label()} "
        "with load, PO, item, and quantity. Store inventory also includes Store #, DC #, pallet dimensions, type, and weight."
    )
    st.download_button(
        label="Download truck plan (.xlsx)",
        data=export_required_columns_excel(_get_selected_records(), st.session_state.truck_trucks),
        file_name="truck_inventory_required_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


def _open_truck_review() -> None:
    st.session_state.truck_workspace = "Review & export"


def render_truck_inventory_view() -> None:
    """Main render function for Truck Inventory module."""
    if st.button("Back to Home"):
        st.session_state["page"] = "home"
        st.stop()

    _initialize_truck_state()

    st.title("Truck Inventory")
    st.caption("Choose your freight, build a plan, then review and adjust your trucks.")
    if not st.session_state.truck_normalized_records:
        render_inputs_tab()
        return

    source = st.session_state.truck_combined_file
    if source:
        st.caption(f"{source.name} · {len(st.session_state.truck_normalized_records):,} imported rows")
    screen = st.radio("Workspace", ["Plan", "Review & export", "Source data"],
                      horizontal=True, key="truck_workspace", label_visibility="collapsed")
    if screen == "Plan":
        with st.expander("Replace or reload Excel file"):
            render_inputs_tab()
        st.subheader("1. Choose freight")
        _render_load_selection()
        if _get_selected_records():
            st.divider()
            st.subheader("2. Confirm pallet details")
            _render_item_setup_section()
            st.divider()
            st.subheader("3. Build your truck plan")
            _render_truck_selection_and_evaluate()
            if st.session_state.truck_trucks:
                st.button("Review truck plan", type="primary", use_container_width=True,
                          on_click=_open_truck_review)
    elif screen == "Review & export":
        if not st.session_state.truck_trucks:
            st.info("Build a truck plan in Plan to review trucks and download assignments.")
        else:
            render_export_tab()
            st.divider()
            render_visualization_tab()
    else:
        with st.expander("Imported rows", expanded=True):
            render_normalized_data_tab()
        with st.expander("Load / DC summary"):
            render_load_summary_tab()


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
    load_selected = parsed and bool(st.session_state.truck_selected_load_number) and bool(selected_records)
    setup_confirmed = st.session_state.truck_item_setup_confirmed_by_load.get(
        st.session_state.truck_selected_load_number,
        False,
    )
    setup_complete = load_selected and setup_confirmed and not setup_issues
    built = bool(st.session_state.truck_trucks)
    export_ready = parsed and built

    steps = [
        ("1. Parse", parsed),
        ("2. Select Load", load_selected),
        ("3. Item Setup", setup_complete),
        ("4. Evaluate", built),
        ("5. Export", export_ready),
    ]
    cols = st.columns(len(steps))
    for col, (label, complete) in zip(cols, steps):
        with col:
            st.metric(label, "Done" if complete else "Pending")
