"""UI-only BOL generator page."""

from __future__ import annotations

import shutil
import tempfile
from pathlib import Path
from time import perf_counter
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
from app.services.bol_file_bundle_service import (
    StandardBundleResult,
    create_multistop_bundles,
    create_standard_bundles,
)
from app.services.bol_doc_upload_parser import (
    BolDocUploadParseResult,
    parse_bol_doc_upload,
)
from app.services.bol_standard_docx_generator import (
    StandardDocxGenerationResult,
    generate_standard_docx_set,
    resolve_output_filename_prefix_for_mode,
    resolve_template_path_for_mode,
)
from app.services.bol_standard_mapper import map_standard_rows_to_records
from app.services.bol_standard_pdf_converter import StandardPdfConversionResult
from app.services.bol_pdf_template_stamper import stamp_bol_pdf_set
from app.services.bol_standard_parser import get_excel_sheet_names, parse_standard_bol_excel
from app.utils.bol_facilities import BOL_FACILITY_LOOKUP, BOL_FACILITY_OPTIONS, BolFacilityRecord


BOL_TEMP_OUTPUT_PREFIXES = (
    "kkg_standard_bol_docx_",
    "kkg_multistop_bol_docx_",
    "kkg_bol_template_pdf_",
    "kkg_standard_bol_pdf_",
    "kkg_standard_bol_bundles_",
    "kkg_multistop_bol_bundles_",
)


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
    if "bol_uploaded_file_signature" not in st.session_state:
        st.session_state["bol_uploaded_file_signature"] = None
    if "bol_selected_worksheet" not in st.session_state:
        st.session_state["bol_selected_worksheet"] = None
    if "bol_parsed_worksheet" not in st.session_state:
        st.session_state["bol_parsed_worksheet"] = None
    if "bol_uploaded_doc_filename" not in st.session_state:
        st.session_state["bol_uploaded_doc_filename"] = None
    if "bol_doc_upload_parse_result" not in st.session_state:
        st.session_state["bol_doc_upload_parse_result"] = None
    if "bol_input_source" not in st.session_state:
        st.session_state["bol_input_source"] = "Excel upload"
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
    if "bol_pdf_source_signature" not in st.session_state:
        st.session_state["bol_pdf_source_signature"] = None
    if "bol_bundle_result" not in st.session_state:
        st.session_state["bol_bundle_result"] = None
    if "bol_bundle_error" not in st.session_state:
        st.session_state["bol_bundle_error"] = None
    if "bol_generation_output_dirs" not in st.session_state:
        st.session_state["bol_generation_output_dirs"] = []
    if "bol_all_files_bundle_requested" not in st.session_state:
        st.session_state["bol_all_files_bundle_requested"] = False
    if "bol_selected_facility_label" not in st.session_state:
        st.session_state["bol_selected_facility_label"] = default_facility_label
    if "bol_selected_facility" not in st.session_state:
        st.session_state["bol_selected_facility"] = default_facility
    if "bol_batch_comment" not in st.session_state:
        st.session_state["bol_batch_comment"] = ""
    if "bol_batch_comment_textarea" not in st.session_state:
        st.session_state["bol_batch_comment_textarea"] = ""
    if "bol_type_selector" not in st.session_state:
        st.session_state["bol_type_selector"] = "PLT"
    if "bol_qty_type_selector" not in st.session_state:
        st.session_state["bol_qty_type_selector"] = "PLT"
    if "bol_batch_name" not in st.session_state:
        st.session_state["bol_batch_name"] = ""
    if "bol_multistop_individual_template_mode" not in st.session_state:
        st.session_state["bol_multistop_individual_template_mode"] = "Standard"


def _is_safe_bol_temp_output_dir(path: Path) -> bool:
    try:
        resolved = path.resolve()
        temp_root = Path(tempfile.gettempdir()).resolve()
    except OSError:
        return False

    return (
        resolved.parent == temp_root
        and resolved.is_dir()
        and any(resolved.name.startswith(prefix) for prefix in BOL_TEMP_OUTPUT_PREFIXES)
    )


def _cleanup_generation_output_dirs() -> None:
    output_dirs = st.session_state.get("bol_generation_output_dirs", [])
    if not isinstance(output_dirs, list):
        st.session_state["bol_generation_output_dirs"] = []
        return

    for output_dir in output_dirs:
        path = Path(str(output_dir))
        if _is_safe_bol_temp_output_dir(path):
            shutil.rmtree(path, ignore_errors=True)
    st.session_state["bol_generation_output_dirs"] = []


def _remember_generation_output_dir(output_dir: str | None) -> None:
    if not output_dir:
        return
    output_dirs = st.session_state.setdefault("bol_generation_output_dirs", [])
    if not isinstance(output_dirs, list):
        output_dirs = []
        st.session_state["bol_generation_output_dirs"] = output_dirs

    resolved = str(Path(output_dir).resolve())
    if resolved not in output_dirs:
        output_dirs.append(resolved)


def _clear_generated_artifacts(*, delete_files: bool) -> None:
    if delete_files:
        _cleanup_generation_output_dirs()
    st.session_state["bol_docx_result"] = None
    st.session_state["bol_pdf_result"] = None
    st.session_state["bol_pdf_source_signature"] = None
    st.session_state["bol_bundle_result"] = None
    st.session_state["bol_bundle_error"] = None
    st.session_state["bol_all_files_bundle_requested"] = False


def _clear_review_state(*, delete_files: bool = False) -> None:
    st.session_state["bol_record_comments"] = {}
    st.session_state["bol_record_selection"] = {}
    _clear_generated_artifacts(delete_files=delete_files)
    st.session_state["bol_generation_status"] = "Waiting for generation action."


def _clear_input_state(*, delete_files: bool = False) -> None:
    st.session_state["bol_uploaded_filename"] = None
    st.session_state["bol_uploaded_file_signature"] = None
    st.session_state["bol_selected_worksheet"] = None
    st.session_state["bol_parsed_worksheet"] = None
    st.session_state["bol_uploaded_doc_filename"] = None
    st.session_state["bol_doc_upload_parse_result"] = None
    st.session_state["bol_parse_requested"] = False
    st.session_state["bol_parse_error"] = None
    st.session_state["bol_parsed_rows"] = []
    st.session_state["bol_grouped_records"] = []
    st.session_state["bol_batch_comment"] = ""
    st.session_state["bol_batch_comment_textarea"] = ""
    _clear_review_state(delete_files=delete_files)


def _clear_worksheet_dependent_state() -> None:
    st.session_state["bol_parse_requested"] = False
    st.session_state["bol_parse_error"] = None
    st.session_state["bol_parsed_rows"] = []
    st.session_state["bol_grouped_records"] = []
    st.session_state["bol_parsed_worksheet"] = None
    _clear_review_state(delete_files=False)


def _uploaded_file_signature(uploaded_file: Any) -> tuple[Any, ...] | None:
    if uploaded_file is None:
        return None

    return (
        getattr(uploaded_file, "file_id", None),
        getattr(uploaded_file, "name", None),
        getattr(uploaded_file, "size", None),
    )


def _default_worksheet_selection(
    sheet_names: list[str],
    previous_selection: str | None,
) -> str | None:
    if not sheet_names:
        return None
    for candidate in (previous_selection, "Revised LS", "Load Sheet"):
        if candidate in sheet_names:
            return candidate
    return sheet_names[0]


def _summary_worksheet_label(input_source: str, mode: str) -> str:
    if input_source == "Doc upload":
        return "DOCX Shipment Request Form"
    if mode in ("Standard", "No Recourse"):
        return str(st.session_state.get("bol_parsed_worksheet") or "")
    return "MAIN LOAD SHEET"


def _refresh_bundles() -> StandardBundleResult | None:
    docx_result = st.session_state["bol_docx_result"]
    pdf_result = st.session_state["bol_pdf_result"]
    mode = st.session_state.get("bol_mode", "Standard")
    include_all_files_bundle = bool(st.session_state.get("bol_all_files_bundle_requested", False))

    if not isinstance(docx_result, StandardDocxGenerationResult):
        st.session_state["bol_bundle_result"] = None
        st.session_state["bol_bundle_error"] = None
        return None

    try:
        if mode == "Multistop":
            bundle_result = create_multistop_bundles(
                generated_docx_files=docx_result.generated_files,
                converted_pdf_files=(
                    pdf_result.converted_files
                    if isinstance(pdf_result, StandardPdfConversionResult)
                    else []
                ),
                bundle_name_prefix="multistop_bol",
                batch_name=st.session_state.get("bol_batch_name", ""),
                include_all_files_bundle=include_all_files_bundle,
            )
        else:
            bundle_result = create_standard_bundles(
                generated_docx_files=docx_result.generated_files,
                converted_pdf_files=(
                    pdf_result.converted_files
                    if isinstance(pdf_result, StandardPdfConversionResult)
                    else []
                ),
                bundle_name_prefix=resolve_output_filename_prefix_for_mode(mode),
                batch_name=st.session_state.get("bol_batch_name", ""),
                include_all_files_bundle=include_all_files_bundle,
            )
        st.session_state["bol_bundle_result"] = bundle_result
        if docx_result.generated_count > 0 and bundle_result.docx_bundle is None:
            st.session_state["bol_bundle_error"] = (
                "DOCX bundle creation failed: generated DOCX files were not found on disk."
            )
        elif (
            isinstance(pdf_result, StandardPdfConversionResult)
            and pdf_result.converted_count > 0
            and bundle_result.pdf_bundle is None
        ):
            st.session_state["bol_bundle_error"] = (
                "PDF bundle creation failed: converted PDF files were not found on disk."
            )
        elif (
            mode == "Multistop"
            and bundle_result.docx_bundle is not None
            and bundle_result.docx_bundle.missing_count > 0
        ):
            st.session_state["bol_bundle_error"] = (
                "Multistop DOCX bundle was created with "
                f"{bundle_result.docx_bundle.missing_count} missing generated source file(s)."
            )
        elif (
            mode == "Multistop"
            and bundle_result.pdf_bundle is not None
            and bundle_result.pdf_bundle.missing_count > 0
        ):
            st.session_state["bol_bundle_error"] = (
                "Multistop PDF bundle was created with "
                f"{bundle_result.pdf_bundle.missing_count} missing converted source file(s)."
            )
        elif bundle_result.combined_pdf_error:
            st.session_state["bol_bundle_error"] = bundle_result.combined_pdf_error
        else:
            st.session_state["bol_bundle_error"] = None
        _remember_generation_output_dir(bundle_result.output_dir)
        return bundle_result
    except Exception as exc:
        st.session_state["bol_bundle_result"] = None
        st.session_state["bol_bundle_error"] = f"Bundle creation failed: {exc}"
        return None


def _clear_generation_state(*, delete_files: bool = False) -> None:
    _clear_generated_artifacts(delete_files=delete_files)
    st.session_state["bol_generation_status"] = "Waiting for generation action."


def _clear_generation_state_references() -> None:
    _clear_generation_state(delete_files=False)


def _prepare_parse_state() -> None:
    st.session_state["bol_parse_requested"] = True
    st.session_state["bol_parse_error"] = None
    _clear_generation_state_references()


def _log_parse_timing(step_name: str, started_at: float) -> float:
    finished_at = perf_counter()
    print(f"BOL parse UI timing: {step_name}={(finished_at - started_at) * 1000:.1f}ms")
    return finished_at


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


def _artifact_exists(path: str | None) -> bool:
    if not path:
        return False
    file_path = Path(path)
    return file_path.exists() and file_path.is_file()


def _download_artifact_button(
    *,
    label: str,
    artifact: Any | None,
    fallback_file_name: str,
    mime: str,
    disabled: bool = False,
) -> bool:
    file_path = Path(artifact.file_path) if artifact is not None else None
    file_exists = file_path is not None and file_path.exists() and file_path.is_file()
    if disabled or not file_exists:
        st.download_button(
            label,
            data=b"",
            file_name=artifact.file_name if artifact is not None else fallback_file_name,
            mime=mime,
            disabled=True,
            use_container_width=True,
        )
        return False

    with file_path.open("rb") as download_file:
        st.download_button(
            label,
            data=download_file,
            file_name=artifact.file_name,
            mime=mime,
            use_container_width=True,
        )
    return True


def _docx_result_signature(docx_result: StandardDocxGenerationResult) -> tuple[tuple[Any, ...], ...]:
    signature_parts: list[tuple[Any, ...]] = []
    for generated_file in docx_result.generated_files:
        file_path = Path(generated_file.file_path)
        try:
            stat_result = file_path.stat()
            modified_ns: int | None = stat_result.st_mtime_ns
            file_size: int | None = stat_result.st_size
        except OSError:
            modified_ns = None
            file_size = None

        signature_parts.append(
            (
                generated_file.bol_number,
                generated_file.file_name,
                str(file_path.resolve()) if file_path.exists() else str(file_path),
                modified_ns,
                file_size,
                getattr(generated_file, "document_type", ""),
                getattr(generated_file, "load_number", ""),
                getattr(generated_file, "stop_number", None),
            )
        )
    return tuple(signature_parts)


def _pdf_result_matches_docx_result(docx_result: StandardDocxGenerationResult) -> bool:
    pdf_result = st.session_state.get("bol_pdf_result")
    if not isinstance(pdf_result, StandardPdfConversionResult):
        return False
    if not pdf_result.conversion_available or pdf_result.converted_count == 0:
        return False
    if pdf_result.failed_count > 0:
        return False
    if st.session_state.get("bol_pdf_source_signature") != _docx_result_signature(docx_result):
        return False
    return all(Path(pdf_file.file_path).exists() for pdf_file in pdf_result.converted_files)


def _generate_pdf_result(
    *,
    mode: str,
    docx_result: StandardDocxGenerationResult,
    grouped_records: list[Any],
    progress_callback: Any,
) -> StandardPdfConversionResult:
    return stamp_bol_pdf_set(
        grouped_records,
        selected_facility=st.session_state["bol_selected_facility"],
        generated_docx_files=docx_result.generated_files,
        mode=mode,
        bol_type=st.session_state.get("bol_type_selector", "PLT"),
        qty_type=st.session_state.get("bol_qty_type_selector", "PLT"),
        batch_comment=st.session_state.get("bol_batch_comment_textarea", ""),
        progress_callback=progress_callback,
    )


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
    if st.session_state["bol_mode"] == "Multistop":
        st.session_state["bol_input_source"] = "Excel upload"

    st.markdown("---")

    st.subheader("BOL Output Options")
    current_bol_type = st.session_state.get("bol_type_selector", "PLT")
    if current_bol_type not in ("PLT", "CASE"):
        current_bol_type = "PLT"
        st.session_state["bol_type_selector"] = current_bol_type
    st.selectbox(
        "TYPE",
        options=["PLT", "CASE"],
        index=["PLT", "CASE"].index(current_bol_type),
        key="bol_type_selector",
        on_change=_clear_generation_state,
    )
    current_qty_type = st.session_state.get("bol_qty_type_selector", "PLT")
    if current_qty_type not in ("PLT", "Case"):
        current_qty_type = "PLT"
        st.session_state["bol_qty_type_selector"] = current_qty_type
    st.selectbox(
        "Qty Type",
        options=["PLT", "Case"],
        index=["PLT", "Case"].index(current_qty_type),
        key="bol_qty_type_selector",
        on_change=_clear_generation_state,
    )
    st.text_input(
        "Batch name",
        key="bol_batch_name",
        placeholder="Optional name for download bundles.",
        on_change=_clear_generation_state,
    )

    st.markdown("---")

    st.subheader("Input")
    input_source = st.session_state.get("bol_input_source", "Excel upload")
    if st.session_state["bol_mode"] in ("Standard", "No Recourse"):
        input_source = st.selectbox(
            "Input source",
            options=["Excel upload", "Doc upload"],
            index=["Excel upload", "Doc upload"].index(input_source),
            key="bol_input_source",
            on_change=_clear_input_state,
        )
    else:
        input_source = "Excel upload"
        st.caption("Multistop BOLs use Excel upload.")

    uploaded_file = None
    uploaded_doc_file = None

    if input_source == "Doc upload":
        st.caption("Accepted file type: .docx")
        uploaded_doc_file = st.file_uploader(
            "Doc upload",
            type=["docx"],
            key="bol_doc_uploader",
        )

        previous_doc_filename = st.session_state["bol_uploaded_doc_filename"]
        if uploaded_doc_file is None:
            st.info("No DOCX file uploaded yet.")
            st.session_state["bol_uploaded_doc_filename"] = None
            _set_selected_facility(None)
            if previous_doc_filename is not None:
                _clear_input_state(delete_files=False)
        else:
            st.session_state["bol_uploaded_doc_filename"] = uploaded_doc_file.name
            st.success(f"Uploaded file: {uploaded_doc_file.name}")
            if previous_doc_filename != uploaded_doc_file.name:
                _clear_input_state(delete_files=False)
                st.session_state["bol_uploaded_doc_filename"] = uploaded_doc_file.name
    else:
        st.caption("Accepted file types: .xlsx, .xlsm, .xls")
        uploaded_file = st.file_uploader(
            "Upload Excel input",
            type=["xlsx", "xlsm", "xls"],
            key="bol_excel_uploader",
        )

        previous_filename = st.session_state["bol_uploaded_filename"]
        previous_file_signature = st.session_state["bol_uploaded_file_signature"]
        if uploaded_file is None:
            st.info("No Excel file uploaded yet.")
            st.session_state["bol_uploaded_filename"] = None
            st.session_state["bol_uploaded_file_signature"] = None
            st.session_state["bol_selected_worksheet"] = None
            st.session_state["bol_parsed_worksheet"] = None
            st.session_state.pop("bol_selected_worksheet_selectbox", None)
            st.session_state["bol_doc_upload_parse_result"] = None
            st.session_state["bol_parsed_rows"] = []
            st.session_state["bol_grouped_records"] = []
            st.session_state["bol_batch_comment"] = ""
            st.session_state["bol_batch_comment_textarea"] = ""
            _set_selected_facility(None)
            _clear_review_state(delete_files=False)
            st.session_state["bol_parse_error"] = None
            st.session_state["bol_parse_requested"] = False
        else:
            st.session_state["bol_uploaded_filename"] = uploaded_file.name
            current_file_signature = _uploaded_file_signature(uploaded_file)
            st.session_state["bol_uploaded_file_signature"] = current_file_signature
            st.success(f"Uploaded file: {uploaded_file.name}")
            if previous_file_signature != current_file_signature:
                st.session_state["bol_parsed_rows"] = []
                st.session_state["bol_grouped_records"] = []
                st.session_state["bol_doc_upload_parse_result"] = None
                st.session_state["bol_parsed_worksheet"] = None
                st.session_state["bol_batch_comment"] = ""
                st.session_state["bol_batch_comment_textarea"] = ""
                default_facility_label = BOL_FACILITY_OPTIONS[0] if BOL_FACILITY_OPTIONS else None
                _set_selected_facility(default_facility_label)
                _clear_review_state(delete_files=False)
                st.session_state["bol_parse_error"] = None
                st.session_state["bol_parse_requested"] = False
                st.session_state["bol_selected_worksheet"] = None
                st.session_state.pop("bol_selected_worksheet_selectbox", None)

            if st.session_state["bol_mode"] in ("Standard", "No Recourse"):
                try:
                    sheet_names = get_excel_sheet_names(uploaded_file)
                except ValueError as exc:
                    st.error(str(exc))
                    sheet_names = []

                if sheet_names:
                    previous_selection = st.session_state.get("bol_selected_worksheet")
                    default_selection = _default_worksheet_selection(
                        sheet_names,
                        previous_selection,
                    )
                    selected_worksheet = st.selectbox(
                        "Select worksheet",
                        options=sheet_names,
                        index=sheet_names.index(default_selection),
                        key="bol_selected_worksheet_selectbox",
                    )
                    if previous_selection != selected_worksheet:
                        st.session_state["bol_selected_worksheet"] = selected_worksheet
                        _clear_worksheet_dependent_state()
                    else:
                        st.session_state["bol_selected_worksheet"] = selected_worksheet

    st.markdown("---")

    st.subheader("Facility Selection")
    if input_source == "Doc upload":
        st.caption("DOCX uploads use the Origin fields from the Shipment Request Form.")

        doc_parse_result = st.session_state.get("bol_doc_upload_parse_result")
        if isinstance(doc_parse_result, BolDocUploadParseResult) and doc_parse_result.records:
            origin = doc_parse_result.records[0].ship_from
            st.caption(
                "Parsed origin: "
                f"{origin.company} | {origin.city_state_zip} | {origin.street}"
            )
        else:
            st.info("Upload and parse a DOCX file to use its Origin fields.")
    else:
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
    batch_comment_input = st.text_input(
        "Batch-level comment (optional)",
        key="bol_batch_comment_textarea",
        placeholder="Optional comment for the entire generated set.",
    )
    st.session_state["bol_batch_comment"] = batch_comment_input

    st.markdown("---")

    st.subheader("Parse")
    parse_disabled = (
        st.session_state["bol_uploaded_doc_filename"] is None
        if input_source == "Doc upload"
        else st.session_state["bol_uploaded_filename"] is None
    )
    parse_button_label = "Parse Doc" if input_source == "Doc upload" else "Parse Excel"
    if st.button(parse_button_label, disabled=parse_disabled):
        parse_started_at = perf_counter()
        _prepare_parse_state()
        last_parse_step_at = _log_parse_timing("state_clear", parse_started_at)

        selected_mode = st.session_state["bol_mode"]
        try:
            if input_source == "Doc upload":
                if selected_mode not in ("Standard", "No Recourse"):
                    raise ValueError("DOCX upload is only supported for Standard and No Recourse BOLs.")
                doc_parse_result = parse_bol_doc_upload(uploaded_doc_file)
                last_parse_step_at = _log_parse_timing("doc_parse", last_parse_step_at)
                grouped_records = doc_parse_result.records
                parsed_rows = []
                st.session_state["bol_record_comments"] = {}
                st.session_state["bol_record_selection"] = {}
                st.session_state["bol_doc_upload_parse_result"] = doc_parse_result
                if grouped_records:
                    origin = grouped_records[0].ship_from
                    st.session_state["bol_selected_facility_label"] = None
                    st.session_state["bol_selected_facility"] = {
                        "facility": "DOC_UPLOAD_ORIGIN",
                        "facility_name": origin.company,
                        "location": origin.city_state_zip,
                        "address": origin.street,
                    }
                    st.session_state["bol_batch_comment"] = grouped_records[0].comments
            elif selected_mode == "Multistop":
                parsed_rows = parse_multistop_bol_excel(uploaded_file)
                last_parse_step_at = _log_parse_timing("excel_parse", last_parse_step_at)
                grouped_records = map_multistop_rows_to_records(parsed_rows)
                last_parse_step_at = _log_parse_timing("mapping", last_parse_step_at)
            else:
                selected_worksheet = st.session_state.get("bol_selected_worksheet")
                parsed_rows = parse_standard_bol_excel(
                    uploaded_file,
                    worksheet_name=selected_worksheet,
                )
                last_parse_step_at = _log_parse_timing("excel_parse", last_parse_step_at)
                grouped_records = map_standard_rows_to_records(parsed_rows)
                last_parse_step_at = _log_parse_timing("mapping", last_parse_step_at)

            if not grouped_records:
                raise ValueError("No grouped BOL records were created from parsed rows.")
            st.session_state["bol_parsed_rows"] = parsed_rows
            st.session_state["bol_grouped_records"] = grouped_records
            st.session_state["bol_parsed_worksheet"] = (
                None
                if input_source == "Doc upload" or selected_mode == "Multistop"
                else st.session_state.get("bol_selected_worksheet")
            )
            _sync_review_state(st.session_state["bol_grouped_records"])
            last_parse_step_at = _log_parse_timing("review_sync", last_parse_step_at)
            _log_parse_timing("total", parse_started_at)
        except ValueError as exc:
            st.session_state["bol_parsed_rows"] = []
            st.session_state["bol_grouped_records"] = []
            st.session_state["bol_doc_upload_parse_result"] = None
            _clear_review_state(delete_files=False)
            st.session_state["bol_parse_error"] = str(exc)
        except Exception as exc:
            st.session_state["bol_parsed_rows"] = []
            st.session_state["bol_grouped_records"] = []
            st.session_state["bol_doc_upload_parse_result"] = None
            _clear_review_state(delete_files=False)
            st.session_state["bol_parse_error"] = f"Unexpected parse error: {exc}"

    if parse_disabled:
        if input_source == "Doc upload":
            st.info("Upload a DOCX file to enable parsing.")
        else:
            st.info("Upload an Excel file to enable parsing.")
    elif st.session_state["bol_parse_error"]:
        st.error(st.session_state["bol_parse_error"])
    elif st.session_state["bol_parse_requested"] and st.session_state["bol_grouped_records"]:
        parsed_rows: list[Any] = st.session_state["bol_parsed_rows"]
        grouped_records: list[BolStandardRecord] = st.session_state["bol_grouped_records"]
        if input_source == "Doc upload":
            unique_bol_count = len({record.bol_number for record in grouped_records if record.bol_number})
        else:
            unique_bol_count = len({row.bol_number for row in parsed_rows if row.bol_number})
        ready_count = sum(1 for record in grouped_records if record.is_ready)
        issue_count = len(grouped_records) - ready_count
        st.success("Parse complete.")
        st.write(
            {
                "source_file": (
                    st.session_state["bol_uploaded_doc_filename"]
                    if input_source == "Doc upload"
                    else st.session_state["bol_uploaded_filename"]
                ),
                "mode": st.session_state["bol_mode"],
                "input_source": input_source,
                "worksheet": _summary_worksheet_label(input_source, selected_mode),
                "rows_parsed": (
                    len(grouped_records) if input_source == "Doc upload" else len(parsed_rows)
                ),
                "unique_bol_numbers_found": unique_bol_count,
                "grouped_records": len(grouped_records),
                "ready_records": ready_count,
                "records_with_issues": issue_count,
                "selected_facility": st.session_state["bol_selected_facility"],
            }
        )
    else:
        st.info(f"Parse summary will appear here after clicking {parse_button_label}.")

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
    pdf_generation_mode_supported = st.session_state["bol_mode"] in (
        "Standard",
        "No Recourse",
        "Multistop",
    )
    generate_all_mode_supported = st.session_state["bol_mode"] in (
        "Standard",
        "No Recourse",
        "Multistop",
    )

    generate_docx_disabled = (not docx_generation_mode_supported) or not any(
        record.selected_for_generation and record.is_ready for record in grouped_records
    )

    if st.session_state["bol_mode"] == "Multistop":
        individual_template_options = ["Standard", "No Recourse"]
        current_individual_template = st.session_state.get(
            "bol_multistop_individual_template_mode",
            "Standard",
        )
        if current_individual_template not in individual_template_options:
            current_individual_template = "Standard"
            st.session_state["bol_multistop_individual_template_mode"] = current_individual_template

        st.selectbox(
            "Choose individual BOL template",
            options=individual_template_options,
            index=individual_template_options.index(current_individual_template),
            key="bol_multistop_individual_template_mode",
            on_change=_clear_generation_state,
        )

    if st.button("Generate DOCX Set", disabled=generate_docx_disabled, use_container_width=True):
        try:
            _clear_generated_artifacts(delete_files=True)
            mode = st.session_state["bol_mode"]
            if mode == "Multistop":
                individual_template_mode = st.session_state.get(
                    "bol_multistop_individual_template_mode",
                    "Standard",
                )
                result = generate_multistop_docx_set(
                    grouped_records,
                    selected_facility=st.session_state["bol_selected_facility"],
                    batch_comment=st.session_state.get("bol_batch_comment_textarea", ""),
                    bol_type=st.session_state.get("bol_type_selector", "PLT"),
                    template_path=MULTISTOP_TEMPLATE_PATH,
                    individual_stop_template_path=resolve_template_path_for_mode(
                        individual_template_mode
                    ),
                    file_name_prefix="multistop_bol",
                )
            else:
                _, template_path = _resolve_generation_context()
                result = generate_standard_docx_set(
                    grouped_records,
                    selected_facility=st.session_state["bol_selected_facility"],
                    batch_comment=st.session_state.get("bol_batch_comment_textarea", ""),
                    bol_type=st.session_state.get("bol_type_selector", "PLT"),
                    qty_type=st.session_state.get("bol_qty_type_selector", "PLT"),
                    template_path=template_path,
                    file_name_prefix=resolve_output_filename_prefix_for_mode(mode),
                )
            st.session_state["bol_docx_result"] = result
            _remember_generation_output_dir(result.output_dir)
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
            _clear_generated_artifacts(delete_files=True)
            st.session_state["bol_generation_status"] = (
                f"{mode} DOCX generation failed: selected template file was not found ({exc})."
            )
        except ValueError as exc:
            mode = st.session_state["bol_mode"]
            _clear_generated_artifacts(delete_files=True)
            st.session_state["bol_generation_status"] = f"{mode} DOCX generation failed: {exc}"
        except Exception as exc:
            mode = st.session_state["bol_mode"]
            _clear_generated_artifacts(delete_files=True)
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

            if _pdf_result_matches_docx_result(docx_result):
                pdf_result = st.session_state["bol_pdf_result"]
                _refresh_bundles()
                st.session_state["bol_generation_status"] = (
                    f"{mode} PDF conversion skipped. Existing PDFs already match "
                    "the current DOCX set."
                )
                st.toast("Existing PDFs already match the current DOCX set.")
            else:
                stale_pdf_result = st.session_state.get("bol_pdf_result")
                stale_bundle_result = st.session_state.get("bol_bundle_result")
                for stale_result in (stale_pdf_result, stale_bundle_result):
                    if isinstance(stale_result, (StandardPdfConversionResult, StandardBundleResult)):
                        path = Path(stale_result.output_dir)
                        if _is_safe_bol_temp_output_dir(path):
                            shutil.rmtree(path, ignore_errors=True)
                st.session_state["bol_pdf_result"] = None
                st.session_state["bol_pdf_source_signature"] = None
                st.session_state["bol_bundle_result"] = None
                st.session_state["bol_bundle_error"] = None
                st.session_state["bol_all_files_bundle_requested"] = False
                progress_bar = st.progress(0)
                progress_status = st.empty()

                def _update_pdf_progress(
                    current_index: int,
                    total_files: int,
                    generated_file: Any,
                ) -> None:
                    progress_status.write(
                        f"Generating PDF {current_index} of {total_files}: "
                        f"{generated_file.file_name}"
                    )
                    progress_bar.progress(current_index / total_files if total_files else 0)

                pdf_result = _generate_pdf_result(
                    mode=mode,
                    docx_result=docx_result,
                    grouped_records=grouped_records,
                    progress_callback=_update_pdf_progress,
                )
                st.session_state["bol_pdf_result"] = pdf_result
                st.session_state["bol_pdf_source_signature"] = _docx_result_signature(docx_result)
                _remember_generation_output_dir(pdf_result.output_dir)
                _refresh_bundles()
                progress_bar.empty()
                progress_status.empty()

                if not pdf_result.conversion_available:
                    st.session_state["bol_generation_status"] = (
                        f"{mode} PDF conversion unavailable. "
                        f"{pdf_result.unavailable_reason}"
                    )
                else:
                    st.session_state["bol_generation_status"] = (
                        f"{mode} PDF generation complete. Created {pdf_result.converted_count}, "
                        f"failed {pdf_result.failed_count}."
                    )
        except Exception as exc:
            mode = st.session_state["bol_mode"]
            st.session_state["bol_pdf_result"] = None
            st.session_state["bol_pdf_source_signature"] = None
            _refresh_bundles()
            st.session_state["bol_generation_status"] = f"{mode} PDF generation failed: {exc}"
    if pdf_generation_mode_supported and generate_pdf_disabled:
        st.caption("Generate DOCX Set first to enable PDF conversion.")

    generate_all_disabled = (not generate_all_mode_supported) or generate_docx_disabled
    if st.button("Generate All", disabled=generate_all_disabled, use_container_width=True):
        try:
            _clear_generated_artifacts(delete_files=True)
            mode = st.session_state["bol_mode"]
            if mode == "Multistop":
                individual_template_mode = st.session_state.get(
                    "bol_multistop_individual_template_mode",
                    "Standard",
                )
                docx_result_all = generate_multistop_docx_set(
                    grouped_records,
                    selected_facility=st.session_state["bol_selected_facility"],
                    batch_comment=st.session_state.get("bol_batch_comment_textarea", ""),
                    bol_type=st.session_state.get("bol_type_selector", "PLT"),
                    template_path=MULTISTOP_TEMPLATE_PATH,
                    individual_stop_template_path=resolve_template_path_for_mode(
                        individual_template_mode
                    ),
                    file_name_prefix="multistop_bol",
                )
            else:
                _, template_path = _resolve_generation_context()
                docx_result_all = generate_standard_docx_set(
                    grouped_records,
                    selected_facility=st.session_state["bol_selected_facility"],
                    batch_comment=st.session_state.get("bol_batch_comment_textarea", ""),
                    bol_type=st.session_state.get("bol_type_selector", "PLT"),
                    qty_type=st.session_state.get("bol_qty_type_selector", "PLT"),
                    template_path=template_path,
                    file_name_prefix=resolve_output_filename_prefix_for_mode(mode),
                )
            st.session_state["bol_docx_result"] = docx_result_all
            _remember_generation_output_dir(docx_result_all.output_dir)

            progress_bar = st.progress(0)
            progress_status = st.empty()

            def _update_generate_all_pdf_progress(
                current_index: int,
                total_files: int,
                generated_file: Any,
            ) -> None:
                progress_status.write(
                    f"Generating PDF {current_index} of {total_files}: "
                    f"{generated_file.file_name}"
                )
                progress_bar.progress(current_index / total_files if total_files else 0)

            pdf_result_all = _generate_pdf_result(
                mode=mode,
                docx_result=docx_result_all,
                grouped_records=grouped_records,
                progress_callback=_update_generate_all_pdf_progress,
            )
            st.session_state["bol_pdf_result"] = pdf_result_all
            st.session_state["bol_pdf_source_signature"] = _docx_result_signature(docx_result_all)
            _remember_generation_output_dir(pdf_result_all.output_dir)
            _refresh_bundles()
            progress_bar.empty()
            progress_status.empty()

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
            _clear_generated_artifacts(delete_files=True)
            st.session_state["bol_generation_status"] = f"{mode} Generate All failed: {exc}"

    if pdf_generation_mode_supported:
        st.caption("DOCX and PDF generation are enabled for Standard, No Recourse, and Multistop modes.")

    st.markdown("---")

    st.subheader("Download")
    bundle_result = st.session_state["bol_bundle_result"]
    docx_bundle_ready = False
    pdf_bundle_ready = False
    combined_pdf_ready = False
    all_bundle_ready = False

    bundle_read_errors: list[str] = []
    if isinstance(bundle_result, StandardBundleResult):
        docx_bundle_ready = _artifact_exists(
            bundle_result.docx_bundle.file_path if bundle_result.docx_bundle else None
        )
        pdf_bundle_ready = _artifact_exists(
            bundle_result.pdf_bundle.file_path if bundle_result.pdf_bundle else None
        )
        combined_pdf_ready = _artifact_exists(
            bundle_result.combined_pdf.file_path if bundle_result.combined_pdf else None
        )
        all_bundle_ready = _artifact_exists(
            bundle_result.all_files_bundle.file_path if bundle_result.all_files_bundle else None
        )

        if bundle_result.docx_bundle and not docx_bundle_ready:
            bundle_read_errors.append(
                f"DOCX bundle file is missing on disk: {bundle_result.docx_bundle.file_path}"
            )
        if bundle_result.pdf_bundle and not pdf_bundle_ready:
            bundle_read_errors.append(
                f"PDF bundle file is missing on disk: {bundle_result.pdf_bundle.file_path}"
            )
        if bundle_result.combined_pdf and not combined_pdf_ready:
            bundle_read_errors.append(
                f"Combined PDF file is missing on disk: {bundle_result.combined_pdf.file_path}"
            )
        if bundle_result.all_files_bundle and not all_bundle_ready:
            bundle_read_errors.append(
                f"Combined bundle file is missing on disk: {bundle_result.all_files_bundle.file_path}"
            )
        if bundle_result.combined_pdf_error:
            bundle_read_errors.append(bundle_result.combined_pdf_error)

    if bundle_read_errors:
        st.session_state["bol_bundle_error"] = " | ".join(bundle_read_errors)

    if st.session_state["bol_mode"] == "Multistop":
        st.caption(
            "Multistop downloads use grouped load/BOL folders for DOCX, PDF, and combined bundles."
        )

    all_bundle_can_be_prepared = isinstance(bundle_result, StandardBundleResult) and (
        bundle_result.docx_bundle is not None or bundle_result.pdf_bundle is not None
    )
    if not st.session_state.get("bol_all_files_bundle_requested", False):
        st.caption("Download All Files is prepared only on request to keep DOCX/PDF generation responsive.")
        if st.button(
            "Prepare All Files Bundle",
            disabled=not all_bundle_can_be_prepared,
            use_container_width=True,
        ):
            st.session_state["bol_all_files_bundle_requested"] = True
            bundle_result = _refresh_bundles()
            if isinstance(bundle_result, StandardBundleResult):
                docx_bundle_ready = _artifact_exists(
                    bundle_result.docx_bundle.file_path if bundle_result.docx_bundle else None
                )
                pdf_bundle_ready = _artifact_exists(
                    bundle_result.pdf_bundle.file_path if bundle_result.pdf_bundle else None
                )
                combined_pdf_ready = _artifact_exists(
                    bundle_result.combined_pdf.file_path if bundle_result.combined_pdf else None
                )
                all_bundle_ready = _artifact_exists(
                    bundle_result.all_files_bundle.file_path if bundle_result.all_files_bundle else None
                )

    mode_prefix = st.session_state["bol_mode"].lower().replace(" ", "_")
    docx_artifact = (
        bundle_result.docx_bundle
        if isinstance(bundle_result, StandardBundleResult)
        else None
    )
    pdf_artifact = (
        bundle_result.pdf_bundle
        if isinstance(bundle_result, StandardBundleResult)
        else None
    )
    combined_pdf_artifact = (
        bundle_result.combined_pdf
        if isinstance(bundle_result, StandardBundleResult)
        else None
    )
    all_files_artifact = (
        bundle_result.all_files_bundle
        if isinstance(bundle_result, StandardBundleResult)
        else None
    )

    docx_bundle_ready = _download_artifact_button(
        label="Download DOCX Bundle",
        artifact=docx_artifact,
        fallback_file_name=f"{mode_prefix}_bol_docx_bundle.zip",
        mime="application/zip",
    )
    pdf_bundle_ready = _download_artifact_button(
        label="Download PDF Bundle",
        artifact=pdf_artifact,
        fallback_file_name=f"{mode_prefix}_bol_pdf_bundle.zip",
        mime="application/zip",
    )
    combined_pdf_ready = False
    if combined_pdf_artifact is not None:
        combined_pdf_ready = _download_artifact_button(
            label="Download Combined PDF",
            artifact=combined_pdf_artifact,
            fallback_file_name=f"{mode_prefix}_bol_combined.pdf",
            mime="application/pdf",
        )
    all_bundle_ready = _download_artifact_button(
        label="Download All Files",
        artifact=all_files_artifact,
        fallback_file_name=f"{mode_prefix}_bol_all_files_bundle.zip",
        mime="application/zip",
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
        selected_individual_template: str | None = None
        if selected_mode == "Multistop":
            selected_template = str(MULTISTOP_TEMPLATE_PATH)
            selected_individual_template = str(
                resolve_template_path_for_mode(
                    st.session_state.get("bol_multistop_individual_template_mode", "Standard")
                )
            )
        else:
            try:
                selected_template = str(resolve_template_path_for_mode(selected_mode))
            except ValueError:
                selected_template = None

        st.write(
            {
                "mode": selected_mode,
                "template_path": selected_template,
                "multistop_individual_template_path": selected_individual_template,
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
            multistop_docx_files = docx_result.generated_files
            combined_docs_generated = sum(
                1
                for generated in multistop_docx_files
                if getattr(generated, "document_type", "") == "combined"
            )
            stop_docs_generated = sum(
                1
                for generated in multistop_docx_files
                if getattr(generated, "document_type", "") == "stop"
            )
            st.write(
                {
                    "grouped_records_generated": combined_docs_generated,
                    "combined_docs_generated": combined_docs_generated,
                    "stop_docs_generated": stop_docs_generated,
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
                    "pdf_converter_path": pdf_result.converter_path,
                    "pdf_conversion_available": pdf_result.conversion_available,
                    "pdf_unavailable_reason": pdf_result.unavailable_reason,
                }
            )

        if isinstance(bundle_result, StandardBundleResult):
            st.write(
                {
                    "docx_bundle_ready": bundle_result.docx_bundle is not None and docx_bundle_ready,
                    "pdf_bundle_ready": bundle_result.pdf_bundle is not None and pdf_bundle_ready,
                    "combined_pdf_ready": bundle_result.combined_pdf is not None and combined_pdf_ready,
                    "combined_bundle_ready": bundle_result.all_files_bundle is not None and all_bundle_ready,
                }
            )

        if isinstance(bundle_result, StandardBundleResult):
            bundle_counts = {
                "docx_bundle_file_count": (
                    bundle_result.docx_bundle.file_count if bundle_result.docx_bundle else 0
                ),
                "pdf_bundle_file_count": (
                    bundle_result.pdf_bundle.file_count if bundle_result.pdf_bundle else 0
                ),
                "combined_pdf_file_count": (
                    bundle_result.combined_pdf.file_count if bundle_result.combined_pdf else 0
                ),
                "combined_bundle_file_count": (
                    bundle_result.all_files_bundle.file_count if bundle_result.all_files_bundle else 0
                ),
            }
            if selected_mode == "Multistop" and bundle_result.docx_bundle:
                bundle_counts.update(
                    {
                        "docx_bundle_group_count": bundle_result.docx_bundle.group_count,
                        "docx_bundle_combined_count": bundle_result.docx_bundle.combined_count,
                        "docx_bundle_stop_count": bundle_result.docx_bundle.stop_count,
                        "docx_bundle_missing_source_count": bundle_result.docx_bundle.missing_count,
                    }
                )
            if selected_mode == "Multistop" and bundle_result.pdf_bundle:
                bundle_counts.update(
                    {
                        "pdf_bundle_group_count": bundle_result.pdf_bundle.group_count,
                        "pdf_bundle_combined_count": bundle_result.pdf_bundle.combined_count,
                        "pdf_bundle_stop_count": bundle_result.pdf_bundle.stop_count,
                        "pdf_bundle_missing_source_count": bundle_result.pdf_bundle.missing_count,
                    }
                )
            st.write(bundle_counts)

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
                st.write(f"- {failed_pdf.bol_number}: {failed_pdf.error} ({failed_pdf.source_docx})")
