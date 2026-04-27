"""UI-only BOL generator page."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from app.models.bol_standard_record import BolStandardRecord
from app.models.bol_standard_row import BolStandardRow
from app.services.bol_multistop_docx_generator import (
    MULTISTOP_TEMPLATE_PATH,
    generate_multistop_docx_set,
)
from app.services.bol_multistop_mapper import map_multistop_rows_to_records
from app.services.bol_multistop_parser import parse_multistop_bol_excel
from app.services.bol_file_bundle_service import StandardBundleResult, create_standard_bundles
from app.services.bol_standard_docx_generator import (
    StandardDocxGenerationResult,
    generate_standard_docx_set,
    resolve_output_filename_prefix_for_mode,
    resolve_template_path_for_mode,
)
from app.services.bol_standard_mapper import map_standard_rows_to_records
from app.services.bol_standard_pdf_converter import (
    StandardPdfConversionResult,
    convert_standard_docx_set_to_pdf,
)
from app.services.bol_standard_parser import parse_standard_bol_excel
from app.utils.bol_facilities import BOL_FACILITY_LOOKUP, BOL_FACILITY_OPTIONS, BolFacilityRecord


def _initialize_bol_state() -> None:
    default_facility_label = BOL_FACILITY_OPTIONS[0] if BOL_FACILITY_OPTIONS else None
    default_facility = (
        BOL_FACILITY_LOOKUP.get(default_facility_label).copy()
        if default_facility_label in BOL_FACILITY_LOOKUP
        else None
    )

    if "bol_mode" not in st.session_state:
        st.session_state["bol_mode"] = "Standard"
    if "bol_uploaded_filename" not in st.session_state:
        st.session_state["bol_uploaded_filename"] = None
    if "bol_parse_requested" not in st.session_state:
        st.session_state["bol_parse_requested"] = False
    if "bol_parse_error" not in st.session_state:
        st.session_state["bol_parse_error"] = None
    if "bol_parsed_rows" not in st.session_state:
        st.session_state["bol_parsed_rows"] = []
    if "bol_grouped_records" not in st.session_state:
        st.session_state["bol_grouped_records"] = []
    if "bol_record_comments" not in st.session_state:
        st.session_state["bol_record_comments"] = {}
    if "bol_record_selection" not in st.session_state:
        st.session_state["bol_record_selection"] = {}
    if "bol_generation_status" not in st.session_state:
        st.session_state["bol_generation_status"] = "Waiting for generation action."
    if "bol_docx_result" not in st.session_state:
        st.session_state["bol_docx_result"] = None
    if "bol_pdf_result" not in st.session_state:
        st.session_state["bol_pdf_result"] = None
    if "bol_bundle_result" not in st.session_state:
        st.session_state["bol_bundle_result"] = None
    if "bol_bundle_error" not in st.session_state:
        st.session_state["bol_bundle_error"] = None
    if "bol_selected_facility_label" not in st.session_state:
        st.session_state["bol_selected_facility_label"] = default_facility_label
    if "bol_selected_facility" not in st.session_state:
        st.session_state["bol_selected_facility"] = default_facility
    if "bol_batch_comment" not in st.session_state:
        st.session_state["bol_batch_comment"] = ""
    if "bol_batch_comment_textarea" not in st.session_state:
        st.session_state["bol_batch_comment_textarea"] = ""


def _clear_review_state() -> None:
    st.session_state["bol_record_comments"] = {}
    st.session_state["bol_record_selection"] = {}
    st.session_state["bol_docx_result"] = None
    st.session_state["bol_pdf_result"] = None
    st.session_state["bol_bundle_result"] = None
    st.session_state["bol_bundle_error"] = None
    st.session_state["bol_generation_status"] = "Waiting for generation action."


def _refresh_bundles() -> StandardBundleResult | None:
    docx_result = st.session_state["bol_docx_result"]
    pdf_result = st.session_state["bol_pdf_result"]
    mode = st.session_state.get("bol_mode", "Standard")

    if not isinstance(docx_result, StandardDocxGenerationResult):
        st.session_state["bol_bundle_result"] = None
        st.session_state["bol_bundle_error"] = None
        return None

    if mode == "Multistop":
        bundle_name_prefix = "multistop_bol"
    else:
        bundle_name_prefix = resolve_output_filename_prefix_for_mode(mode)

    try:
        bundle_result = create_standard_bundles(
            generated_docx_files=docx_result.generated_files,
            converted_pdf_files=(
                pdf_result.converted_files
                if isinstance(pdf_result, StandardPdfConversionResult)
                else []
            ),
            bundle_name_prefix=bundle_name_prefix,
        )
        st.session_state["bol_bundle_result"] = bundle_result
        if docx_result.generated_count > 0 and bundle_result.docx_bundle is None:
            st.session_state["bol_bundle_error"] = (
                "DOCX bundle creation failed: generated DOCX files were not found on disk."
            )
        else:
            st.session_state["bol_bundle_error"] = None
        return bundle_result
    except Exception as exc:
        st.session_state["bol_bundle_result"] = None
        st.session_state["bol_bundle_error"] = f"Bundle creation failed: {exc}"
        return None


def _clear_generation_state() -> None:
    st.session_state["bol_docx_result"] = None
    st.session_state["bol_pdf_result"] = None
    st.session_state["bol_bundle_result"] = None
    st.session_state["bol_bundle_error"] = None
    st.session_state["bol_generation_status"] = "Waiting for generation action."


def _set_selected_facility(facility_label: str | None) -> None:
    if facility_label is None or facility_label not in BOL_FACILITY_LOOKUP:
        st.session_state["bol_selected_facility_label"] = None
        st.session_state["bol_selected_facility"] = None
        return

    st.session_state["bol_selected_facility_label"] = facility_label
    st.session_state["bol_selected_facility"] = BOL_FACILITY_LOOKUP[facility_label].copy()


def _resolve_generation_context() -> tuple[str, Path]:
    mode = st.session_state["bol_mode"]
    template_path = resolve_template_path_for_mode(mode)
    return mode, template_path


def _read_file_bytes(path: str) -> bytes | None:
    file_path = Path(path)
    if not file_path.exists():
        return None
    return file_path.read_bytes()


def _format_total_skids(value: float) -> int | float:
    if float(value).is_integer():
        return int(value)
    return value


def _record_key(record: Any, index: int) -> str:
    group_key = str(getattr(record, "group_key", "")).strip()
    if group_key:
        return group_key

    bol_number = str(getattr(record, "bol_number", "")).strip()
    load_number = str(getattr(record, "load_number", "")).strip()
    kk_load_number = str(getattr(record, "kk_load_number", "")).strip()

    if bol_number and load_number:
        return f"{bol_number}::{load_number}"
    if bol_number and kk_load_number:
        return f"{bol_number}::{kk_load_number}"
    if bol_number:
        return bol_number
    return f"__MISSING_BOL__{index}"


def _widget_safe_key(raw_key: str) -> str:
    return "".join(char if char.isalnum() else "_" for char in raw_key)


def _sync_review_state(records: list[Any]) -> None:
    comments_state: dict[str, str] = st.session_state["bol_record_comments"]
    selection_state: dict[str, bool] = st.session_state["bol_record_selection"]

    current_keys: set[str] = set()
    for index, record in enumerate(records):
        key = _record_key(record, index)
        current_keys.add(key)

        if key not in comments_state:
            comments_state[key] = record.comments

        if key not in selection_state:
            selection_state[key] = record.is_ready

        record.comments = comments_state[key]
        record.selected_for_generation = selection_state[key]

    stale_comment_keys = [key for key in comments_state if key not in current_keys]
    for key in stale_comment_keys:
        del comments_state[key]

    stale_selection_keys = [key for key in selection_state if key not in current_keys]
    for key in stale_selection_keys:
        del selection_state[key]


def _build_stop_summary(record: Any) -> str:
    stops = getattr(record, "stops", None)
    if not isinstance(stops, list) or not stops:
        return "N/A"

    parts: list[str] = []
    for stop in stops:
        stop_no = getattr(stop, "stop_number", "")
        delivery_dc = str(getattr(stop, "delivery_dc", "")).strip()
        dc_number = str(getattr(stop, "dc_number", "")).strip()

        if delivery_dc:
            summary_value = delivery_dc
        elif dc_number:
            summary_value = dc_number
        else:
            summary_value = "(missing DC)"

        parts.append(f"Stop {stop_no}: {summary_value}")

    return " | ".join(parts) if parts else "N/A"


def _records_to_review_records(records: list[Any], mode: str) -> pd.DataFrame:
    if not records:
        return pd.DataFrame(
            columns=[
                "BOL number",
                "Load number",
                "PO number",
                "Ship date",
                "Carrier",
                "Stop count",
                "Stop summary",
                "Ship from",
                "Ship to",
                "Total cases",
                "Total pallets",
                "Total weight",
                "Item line count",
                "Mode",
                "Status",
            ]
        )

    review_rows: list[dict[str, str | int | float]] = []
    for record in records:
        ship_from = (
            f"{record.ship_from.company}, "
            f"{record.ship_from.street}, "
            f"{record.ship_from.city_state_zip}"
        )
        ship_to_parts = [
            part
            for part in [
                record.consignee_company,
                record.consignee_street,
                record.consignee_city_state_zip,
            ]
            if part
        ]
        ship_to = " - ".join(ship_to_parts)

        review_rows.append(
            {
                "BOL number": record.bol_number,
                "Load number": getattr(record, "load_number", record.kk_load_number),
                "PO number": record.po_number,
                "Ship date": record.ship_date,
                "Carrier": record.carrier,
                "Stop count": getattr(record, "stop_count", "N/A"),
                "Stop summary": _build_stop_summary(record),
                "Ship from": ship_from,
                "Ship to": ship_to,
                "Total cases": _format_total_skids(record.total_skids),
                "Total pallets": getattr(record, "total_pallet", "N/A"),
                "Total weight": (
                    getattr(record, "total_ship_weight", "N/A")
                    if getattr(record, "total_ship_weight", None) not in (None, "")
                    else "N/A"
                ),
                "Item line count": len(record.item_lines),
                "Mode": mode,
                "Status": record.status,
            }
        )

    return pd.DataFrame(review_rows)


def _multistop_skip_breakdown(result: StandardDocxGenerationResult) -> dict[str, int]:
    excluded = 0
    validation = 0
    other = 0

    for skipped in result.skipped_records:
        reason = skipped.reason.strip().lower()
        if reason == "record excluded in review.":
            excluded += 1
        elif (
            "unsupported stop count" in reason
            or "missing required data" in reason
            or "not ready" in reason
            or "missing stop" in reason
            or "malformed stop" in reason
            or "duplicate stop" in reason
        ):
            validation += 1
        else:
            other += 1

    return {
        "excluded_in_review": excluded,
        "validation_skipped": validation,
        "other_skipped": other,
    }


def render_bol_generator_view() -> None:
    _initialize_bol_state()

    if st.button("Back to Home"):
        st.session_state["page"] = "home"
        st.stop()

    st.markdown("---")

    st.title("BOL Generator")
    st.caption("Batch workflow for reviewing records and preparing BOL output sets.")

    st.subheader("Mode Selection")
    st.session_state["bol_mode"] = st.radio(
        "Select BOL mode",
        options=["Standard", "No Recourse", "Multistop"],
        horizontal=True,
        key="bol_mode_radio",
        index=["Standard", "No Recourse", "Multistop"].index(st.session_state["bol_mode"]),
    )

    st.markdown("---")

    st.subheader("Upload Excel")
    st.caption("Accepted file types: .xlsx, .xlsm, .xls")
    uploaded_file = st.file_uploader(
        "Upload Excel input",
        type=["xlsx", "xlsm", "xls"],
        key="bol_excel_uploader",
    )

    previous_filename = st.session_state["bol_uploaded_filename"]
    if uploaded_file is None:
        st.info("No Excel file uploaded yet.")
        st.session_state["bol_uploaded_filename"] = None
        st.session_state["bol_parsed_rows"] = []
        st.session_state["bol_grouped_records"] = []
        st.session_state["bol_batch_comment"] = ""
        st.session_state["bol_batch_comment_textarea"] = ""
        _set_selected_facility(None)
        _clear_review_state()
        st.session_state["bol_parse_error"] = None
        st.session_state["bol_parse_requested"] = False
    else:
        st.session_state["bol_uploaded_filename"] = uploaded_file.name
        st.success(f"Uploaded file: {uploaded_file.name}")
        if previous_filename != uploaded_file.name:
            st.session_state["bol_parsed_rows"] = []
            st.session_state["bol_grouped_records"] = []
            st.session_state["bol_batch_comment"] = ""
            st.session_state["bol_batch_comment_textarea"] = ""
            default_facility_label = BOL_FACILITY_OPTIONS[0] if BOL_FACILITY_OPTIONS else None
            _set_selected_facility(default_facility_label)
            _clear_review_state()
            st.session_state["bol_parse_error"] = None
            st.session_state["bol_parse_requested"] = False

    st.markdown("---")

    st.subheader("Facility Selection")
    st.caption("Choose one ship-from facility for the current uploaded batch.")

    if uploaded_file is None:
        st.info("Upload an Excel file to select a facility for this batch.")
    else:
        current_label = st.session_state.get("bol_selected_facility_label")
        if current_label not in BOL_FACILITY_LOOKUP:
            current_label = BOL_FACILITY_OPTIONS[0] if BOL_FACILITY_OPTIONS else None

        selected_label = st.selectbox(
            "Select ship-from facility (batch-level)",
            options=list(BOL_FACILITY_OPTIONS),
            index=BOL_FACILITY_OPTIONS.index(current_label) if current_label else 0,
            key="bol_batch_facility_selectbox",
        )
        _set_selected_facility(selected_label)

        selected_facility: BolFacilityRecord | None = st.session_state["bol_selected_facility"]
        if selected_facility:
            st.caption(
                "Selected facility details: "
                f"{selected_facility['facility_name']} | "
                f"{selected_facility['location']} | "
                f"{selected_facility['address']}"
            )

    st.markdown("---")

    st.subheader("Batch Comment")
    st.caption("Optional comment for entire BOL program.")
    batch_comment_input = st.text_area(
        "Batch-level comment (optional)",
        key="bol_batch_comment_textarea",
        placeholder="Optional comment for the entire generated set.",
    )
    st.session_state["bol_batch_comment"] = batch_comment_input

    st.markdown("---")

    st.subheader("Parse")
    parse_disabled = st.session_state["bol_uploaded_filename"] is None
    if st.button("Parse Excel", type="primary", disabled=parse_disabled):
        st.session_state["bol_parse_requested"] = True
        st.session_state["bol_parse_error"] = None
        _clear_generation_state()

        selected_mode = st.session_state["bol_mode"]
        try:
            if selected_mode == "Multistop":
                parsed_rows = parse_multistop_bol_excel(uploaded_file)
                grouped_records = map_multistop_rows_to_records(parsed_rows)
            else:
                parsed_rows = parse_standard_bol_excel(uploaded_file)
                grouped_records = map_standard_rows_to_records(parsed_rows)

            if not grouped_records:
                raise ValueError("No grouped BOL records were created from parsed rows.")
            st.session_state["bol_parsed_rows"] = parsed_rows
            st.session_state["bol_grouped_records"] = grouped_records
            _sync_review_state(st.session_state["bol_grouped_records"])
        except ValueError as exc:
            st.session_state["bol_parsed_rows"] = []
            st.session_state["bol_grouped_records"] = []
            _clear_review_state()
            st.session_state["bol_parse_error"] = str(exc)
        except Exception as exc:
            st.session_state["bol_parsed_rows"] = []
            st.session_state["bol_grouped_records"] = []
            _clear_review_state()
            st.session_state["bol_parse_error"] = f"Unexpected parse error: {exc}"

    if parse_disabled:
        st.info("Upload an Excel file to enable parsing.")
    elif st.session_state["bol_parse_error"]:
        st.error(st.session_state["bol_parse_error"])
    elif st.session_state["bol_parse_requested"] and st.session_state["bol_grouped_records"]:
        parsed_rows: list[BolStandardRow] = st.session_state["bol_parsed_rows"]
        grouped_records: list[BolStandardRecord] = st.session_state["bol_grouped_records"]
        unique_bol_count = len({row.bol_number for row in parsed_rows if row.bol_number})
        ready_count = sum(1 for record in grouped_records if record.is_ready)
        issue_count = len(grouped_records) - ready_count
        st.success("Parse complete.")
        st.write(
            {
                "source_file": st.session_state["bol_uploaded_filename"],
                "mode": st.session_state["bol_mode"],
                "worksheet": "MAIN LOAD SHEET",
                "rows_parsed": len(parsed_rows),
                "unique_bol_numbers_found": unique_bol_count,
                "grouped_records": len(grouped_records),
                "ready_records": ready_count,
                "records_with_issues": issue_count,
                "selected_facility": st.session_state["bol_selected_facility"],
            }
        )
    else:
        st.info("Parse summary will appear here after clicking Parse Excel.")

    st.markdown("---")

    with st.expander("Review Records", expanded=False):
        grouped_records: list[Any] = st.session_state["bol_grouped_records"]
        _sync_review_state(grouped_records)

        comments_state: dict[str, str] = st.session_state["bol_record_comments"]
        selection_state: dict[str, bool] = st.session_state["bol_record_selection"]

        for index, record in enumerate(grouped_records):
            key = _record_key(record, index)
            safe_key = _widget_safe_key(key)
            label = record.bol_number or "(missing BOL #)"

            include_value = st.checkbox(
                f"Include BOL {label}",
                value=selection_state.get(key, record.is_ready),
                key=f"bol_include_{index}_{safe_key}",
            )
            selection_state[key] = include_value
            record.selected_for_generation = include_value

            comments_value = st.text_area(
                f"Comments for BOL {label}",
                value=comments_state.get(key, record.comments),
                key=f"bol_comments_{index}_{safe_key}",
                placeholder="Optional notes for this BOL record.",
            )
            comments_state[key] = comments_value
            record.comments = comments_value

            if record.missing_required_fields:
                st.caption("Missing required: " + ", ".join(record.missing_required_fields))
            if record.status == "Unsupported Stop Count":
                st.caption("Unsupported: stop count exceeds the current maximum of 3.")
            if record.warnings:
                st.caption("Warnings: " + " | ".join(record.warnings))
            if record.issues and not (record.missing_required_fields or record.warnings):
                st.caption("Issues: " + " | ".join(record.issues))

        total_records = len(grouped_records)
        ready_records = sum(1 for record in grouped_records if record.is_ready)
        issue_records = total_records - ready_records
        selected_records = sum(
            1
            for record in grouped_records
            if record.selected_for_generation and record.is_ready
        )

        metric_cols = st.columns(4)
        metric_cols[0].metric("Total records found", total_records)
        metric_cols[1].metric("Ready records", ready_records)
        metric_cols[2].metric("Records with issues", issue_records)
        metric_cols[3].metric("Records selected for generation", selected_records)

        st.dataframe(
            _records_to_review_records(grouped_records, st.session_state["bol_mode"]),
            use_container_width=True,
            hide_index=True,
        )

    st.markdown("---")

    st.subheader("Generate")
    selected_records_total = sum(1 for record in grouped_records if record.selected_for_generation)
    selected_ready_records = sum(
        1
        for record in grouped_records
        if record.selected_for_generation and record.is_ready
    )

    if grouped_records and selected_records_total == 0:
        st.warning("No records are selected for generation.")
    elif grouped_records and selected_records_total > 0 and selected_ready_records == 0:
        st.warning("Selected records exist, but none are ready for generation. Resolve missing data first.")

    docx_generation_mode_supported = st.session_state["bol_mode"] in (
        "Standard",
        "No Recourse",
        "Multistop",
    )
    pdf_generation_mode_supported = st.session_state["bol_mode"] in ("Standard", "No Recourse")
    generate_all_mode_supported = st.session_state["bol_mode"] in ("Standard", "No Recourse")

    generate_docx_disabled = (not docx_generation_mode_supported) or not any(
        record.selected_for_generation and record.is_ready for record in grouped_records
    )
    if st.button("Generate DOCX Set", disabled=generate_docx_disabled, use_container_width=True):
        try:
            mode = st.session_state["bol_mode"]
            if mode == "Multistop":
                result = generate_multistop_docx_set(
                    grouped_records,
                    selected_facility=st.session_state["bol_selected_facility"],
                    batch_comment=st.session_state.get("bol_batch_comment_textarea", ""),
                    template_path=MULTISTOP_TEMPLATE_PATH,
                    file_name_prefix="multistop_bol",
                )
            else:
                _, template_path = _resolve_generation_context()
                result = generate_standard_docx_set(
                    grouped_records,
                    selected_facility=st.session_state["bol_selected_facility"],
                    batch_comment=st.session_state.get("bol_batch_comment_textarea", ""),
                    template_path=template_path,
                    file_name_prefix=resolve_output_filename_prefix_for_mode(mode),
                )
            st.session_state["bol_docx_result"] = result
            st.session_state["bol_pdf_result"] = None
            _refresh_bundles()
            if mode == "Multistop":
                skip_breakdown = _multistop_skip_breakdown(result)
                st.session_state["bol_generation_status"] = (
                    f"{mode} DOCX generation complete. Generated {result.generated_count}, "
                    f"skipped {result.skipped_count} "
                    f"(validation {skip_breakdown['validation_skipped']}, "
                    f"excluded {skip_breakdown['excluded_in_review']}, "
                    f"other {skip_breakdown['other_skipped']}), "
                    f"failed {result.failed_count}."
                )
            else:
                st.session_state["bol_generation_status"] = (
                    f"{mode} DOCX generation complete. Generated {result.generated_count}, "
                    f"skipped {result.skipped_count}, failed {result.failed_count}."
                )
        except FileNotFoundError as exc:
            mode = st.session_state["bol_mode"]
            st.session_state["bol_docx_result"] = None
            st.session_state["bol_pdf_result"] = None
            st.session_state["bol_bundle_result"] = None
            st.session_state["bol_bundle_error"] = None
            st.session_state["bol_generation_status"] = (
                f"{mode} DOCX generation failed: selected template file was not found ({exc})."
            )
        except ValueError as exc:
            mode = st.session_state["bol_mode"]
            st.session_state["bol_docx_result"] = None
            st.session_state["bol_pdf_result"] = None
            st.session_state["bol_bundle_result"] = None
            st.session_state["bol_bundle_error"] = None
            st.session_state["bol_generation_status"] = f"{mode} DOCX generation failed: {exc}"
        except Exception as exc:
            mode = st.session_state["bol_mode"]
            st.session_state["bol_docx_result"] = None
            st.session_state["bol_pdf_result"] = None
            st.session_state["bol_bundle_result"] = None
            st.session_state["bol_bundle_error"] = None
            st.session_state["bol_generation_status"] = (
                f"Unexpected {mode} DOCX generation error: {exc}"
            )

    docx_result = st.session_state["bol_docx_result"]
    generate_pdf_disabled = not (
        isinstance(docx_result, StandardDocxGenerationResult)
        and docx_result.generated_count > 0
    )
    if st.button(
        "Generate PDF Set",
        disabled=(not pdf_generation_mode_supported) or generate_pdf_disabled,
        use_container_width=True,
    ):
        try:
            mode = st.session_state["bol_mode"]
            if not isinstance(docx_result, StandardDocxGenerationResult):
                raise ValueError("Generate DOCX Set first.")

            pdf_result = convert_standard_docx_set_to_pdf(docx_result.generated_files)
            st.session_state["bol_pdf_result"] = pdf_result
            _refresh_bundles()

            if not pdf_result.conversion_available:
                st.session_state["bol_generation_status"] = (
                    f"{mode} PDF conversion unavailable. "
                    f"{pdf_result.unavailable_reason}"
                )
            else:
                st.session_state["bol_generation_status"] = (
                    f"{mode} PDF conversion complete. Converted {pdf_result.converted_count}, "
                    f"failed {pdf_result.failed_count}."
                )
        except Exception as exc:
            mode = st.session_state["bol_mode"]
            st.session_state["bol_pdf_result"] = None
            _refresh_bundles()
            st.session_state["bol_generation_status"] = f"{mode} PDF generation failed: {exc}"

    generate_all_disabled = (not generate_all_mode_supported) or generate_docx_disabled
    if st.button("Generate All", disabled=generate_all_disabled, use_container_width=True):
        try:
            mode, template_path = _resolve_generation_context()
            docx_result_all = generate_standard_docx_set(
                grouped_records,
                selected_facility=st.session_state["bol_selected_facility"],
                batch_comment=st.session_state.get("bol_batch_comment_textarea", ""),
                template_path=template_path,
                file_name_prefix=resolve_output_filename_prefix_for_mode(mode),
            )
            st.session_state["bol_docx_result"] = docx_result_all

            pdf_result_all = convert_standard_docx_set_to_pdf(docx_result_all.generated_files)
            st.session_state["bol_pdf_result"] = pdf_result_all
            _refresh_bundles()

            if not pdf_result_all.conversion_available:
                st.session_state["bol_generation_status"] = (
                    f"{mode} Generate All: DOCX generated {docx_result_all.generated_count}. "
                    "PDF conversion unavailable. "
                    f"{pdf_result_all.unavailable_reason}"
                )
            else:
                st.session_state["bol_generation_status"] = (
                    f"{mode} Generate All complete. DOCX {docx_result_all.generated_count}, "
                    f"PDF {pdf_result_all.converted_count}, "
                    f"PDF failures {pdf_result_all.failed_count}."
                )
        except Exception as exc:
            mode = st.session_state["bol_mode"]
            st.session_state["bol_generation_status"] = f"{mode} Generate All failed: {exc}"

    if pdf_generation_mode_supported:
        st.caption("DOCX and PDF generation are enabled for Standard and No Recourse modes.")
    else:
        st.caption(
            "Multistop supports DOCX generation and DOCX bundle download only in this phase. "
            "PDF conversion remains unavailable."
        )

    st.markdown("---")

    st.subheader("Download")
    bundle_result = st.session_state["bol_bundle_result"]
    docx_bundle_bytes = None
    pdf_bundle_bytes = None
    all_bundle_bytes = None

    if isinstance(bundle_result, StandardBundleResult):
        if bundle_result.docx_bundle:
            docx_bundle_bytes = _read_file_bytes(bundle_result.docx_bundle.file_path)
        if bundle_result.pdf_bundle:
            pdf_bundle_bytes = _read_file_bytes(bundle_result.pdf_bundle.file_path)
        if bundle_result.all_files_bundle:
            all_bundle_bytes = _read_file_bytes(bundle_result.all_files_bundle.file_path)

    bundle_read_errors: list[str] = []
    if isinstance(bundle_result, StandardBundleResult):
        if bundle_result.docx_bundle and docx_bundle_bytes is None:
            bundle_read_errors.append(
                f"DOCX bundle file is missing on disk: {bundle_result.docx_bundle.file_path}"
            )
        if bundle_result.pdf_bundle and pdf_bundle_bytes is None:
            bundle_read_errors.append(
                f"PDF bundle file is missing on disk: {bundle_result.pdf_bundle.file_path}"
            )
        if bundle_result.all_files_bundle and all_bundle_bytes is None:
            bundle_read_errors.append(
                f"Combined bundle file is missing on disk: {bundle_result.all_files_bundle.file_path}"
            )

    if bundle_read_errors:
        st.session_state["bol_bundle_error"] = " | ".join(bundle_read_errors)

    if st.session_state["bol_mode"] == "Multistop":
        st.caption("Multistop: DOCX bundle download is available after DOCX generation. PDF downloads are disabled.")

    st.download_button(
        "Download DOCX Bundle",
        data=docx_bundle_bytes or b"",
        file_name=(
            bundle_result.docx_bundle.file_name
            if isinstance(bundle_result, StandardBundleResult) and bundle_result.docx_bundle
            else f"{st.session_state['bol_mode'].lower().replace(' ', '_')}_bol_docx_bundle.zip"
        ),
        mime="application/zip",
        disabled=docx_bundle_bytes is None,
        use_container_width=True,
    )
    st.download_button(
        "Download PDF Bundle",
        data=pdf_bundle_bytes or b"",
        file_name=(
            bundle_result.pdf_bundle.file_name
            if isinstance(bundle_result, StandardBundleResult) and bundle_result.pdf_bundle
            else f"{st.session_state['bol_mode'].lower().replace(' ', '_')}_bol_pdf_bundle.zip"
        ),
        mime="application/zip",
        disabled=(st.session_state["bol_mode"] == "Multistop") or (pdf_bundle_bytes is None),
        use_container_width=True,
    )
    st.download_button(
        "Download All Files",
        data=all_bundle_bytes or b"",
        file_name=(
            bundle_result.all_files_bundle.file_name
            if isinstance(bundle_result, StandardBundleResult) and bundle_result.all_files_bundle
            else f"{st.session_state['bol_mode'].lower().replace(' ', '_')}_bol_all_files_bundle.zip"
        ),
        mime="application/zip",
        disabled=(st.session_state["bol_mode"] == "Multistop") or (all_bundle_bytes is None),
        use_container_width=True,
    )

    st.markdown("---")

    st.subheader("Status & Results")
    st.info(f"Generation status: {st.session_state['bol_generation_status']}")
    if st.session_state["bol_bundle_error"]:
        st.error(st.session_state["bol_bundle_error"])

    docx_result = st.session_state["bol_docx_result"]
    pdf_result = st.session_state["bol_pdf_result"]
    if isinstance(docx_result, StandardDocxGenerationResult):
        selected_mode = st.session_state["bol_mode"]
        selected_template: str | None = None
        if selected_mode == "Multistop":
            selected_template = str(MULTISTOP_TEMPLATE_PATH)
        else:
            try:
                selected_template = str(resolve_template_path_for_mode(selected_mode))
            except ValueError:
                selected_template = None

        st.write(
            {
                "mode": selected_mode,
                "template_path": selected_template,
                "output_directory": docx_result.output_dir,
                "docx_generated": docx_result.generated_count,
                "records_skipped": docx_result.skipped_count,
                "docx_generation_failures": docx_result.failed_count,
                "docx_generation_notices": len(docx_result.notices),
                "selected_records_total": selected_records_total,
                "selected_ready_records": selected_ready_records,
                "selected_facility": st.session_state["bol_selected_facility"],
            }
        )

        if selected_mode == "Multistop":
            skip_breakdown = _multistop_skip_breakdown(docx_result)
            st.write(
                {
                    "validation_skipped": skip_breakdown["validation_skipped"],
                    "excluded_in_review": skip_breakdown["excluded_in_review"],
                    "other_skipped": skip_breakdown["other_skipped"],
                    "generation_failures": docx_result.failed_count,
                }
            )

        if isinstance(pdf_result, StandardPdfConversionResult):
            st.write(
                {
                    "pdf_generated": pdf_result.converted_count,
                    "pdf_conversion_failures": pdf_result.failed_count,
                    "pdf_converter": pdf_result.converter_name,
                    "pdf_conversion_available": pdf_result.conversion_available,
                    "pdf_unavailable_reason": pdf_result.unavailable_reason,
                }
            )

        if isinstance(bundle_result, StandardBundleResult):
            st.write(
                {
                    "docx_bundle_ready": bundle_result.docx_bundle is not None and docx_bundle_bytes is not None,
                    "pdf_bundle_ready": bundle_result.pdf_bundle is not None and pdf_bundle_bytes is not None,
                    "combined_bundle_ready": bundle_result.all_files_bundle is not None and all_bundle_bytes is not None,
                }
            )

        if isinstance(bundle_result, StandardBundleResult):
            st.write(
                {
                    "docx_bundle_file_count": (
                        bundle_result.docx_bundle.file_count if bundle_result.docx_bundle else 0
                    ),
                    "pdf_bundle_file_count": (
                        bundle_result.pdf_bundle.file_count if bundle_result.pdf_bundle else 0
                    ),
                    "combined_bundle_file_count": (
                        bundle_result.all_files_bundle.file_count if bundle_result.all_files_bundle else 0
                    ),
                }
            )

        if docx_result.generated_files:
            st.caption("Generated DOCX files:")
            for generated in docx_result.generated_files:
                st.write(f"- {generated.file_name} ({generated.bol_number})")

        if docx_result.skipped_records:
            st.caption("Skipped records:")
            for skipped in docx_result.skipped_records:
                st.write(f"- {skipped.bol_number}: {skipped.reason}")

        if docx_result.failed_records:
            st.caption("Failed records:")
            for failed in docx_result.failed_records:
                st.write(f"- {failed.bol_number}: {failed.error}")

        if docx_result.notices:
            st.caption("Generation notices:")
            for notice in docx_result.notices:
                st.write(f"- {notice.bol_number}: {notice.message}")

        if isinstance(pdf_result, StandardPdfConversionResult) and pdf_result.failed_conversions:
            st.caption("PDF conversion failures:")
            for failed_pdf in pdf_result.failed_conversions:
                st.write(f"- {failed_pdf.bol_number}: {failed_pdf.error}")
