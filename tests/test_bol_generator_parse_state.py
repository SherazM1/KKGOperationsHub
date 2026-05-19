from __future__ import annotations

import importlib
import sys
from types import SimpleNamespace

from app.models.bol_standard_record import BolAddressBlock
from app.ui import bol_generator


def test_prepare_parse_state_only_clears_artifact_references(monkeypatch) -> None:
    def fail_if_called(*args, **kwargs):
        raise AssertionError("Parse preparation must not build bundles, stamp PDFs, or delete folders.")

    monkeypatch.setattr(bol_generator, "_cleanup_generation_output_dirs", fail_if_called)
    monkeypatch.setattr(bol_generator, "_refresh_bundles", fail_if_called)
    monkeypatch.setattr(bol_generator, "stamp_bol_pdf_set", fail_if_called)

    bol_generator.st.session_state["bol_parse_requested"] = False
    bol_generator.st.session_state["bol_parse_error"] = "old error"
    bol_generator.st.session_state["bol_docx_result"] = object()
    bol_generator.st.session_state["bol_pdf_result"] = object()
    bol_generator.st.session_state["bol_pdf_source_signature"] = object()
    bol_generator.st.session_state["bol_bundle_result"] = object()
    bol_generator.st.session_state["bol_bundle_error"] = "old bundle error"
    bol_generator.st.session_state["bol_all_files_bundle_requested"] = True

    bol_generator._prepare_parse_state()

    assert bol_generator.st.session_state["bol_parse_requested"] is True
    assert bol_generator.st.session_state["bol_parse_error"] is None
    assert bol_generator.st.session_state["bol_docx_result"] is None
    assert bol_generator.st.session_state["bol_pdf_result"] is None
    assert bol_generator.st.session_state["bol_pdf_source_signature"] is None
    assert bol_generator.st.session_state["bol_bundle_result"] is None
    assert bol_generator.st.session_state["bol_bundle_error"] is None
    assert bol_generator.st.session_state["bol_all_files_bundle_requested"] is False


def test_initialize_bol_state_stores_selected_worksheet() -> None:
    bol_generator.st.session_state.clear()

    bol_generator._initialize_bol_state()

    assert bol_generator.st.session_state["bol_selected_worksheet"] is None
    assert bol_generator.st.session_state["bol_parsed_worksheet"] is None
    assert bol_generator.st.session_state["bol_render_pickup_number"] == "Yes"


def test_default_worksheet_selection_prefers_previous_then_named_defaults() -> None:
    sheet_names = ["Tracker Info", "Revised LS", "Load Sheet"]

    assert (
        bol_generator._default_worksheet_selection(sheet_names, "Load Sheet")
        == "Load Sheet"
    )
    assert bol_generator._default_worksheet_selection(sheet_names, "Missing") == "Revised LS"
    assert (
        bol_generator._default_worksheet_selection(["Tracker Info", "Load Sheet"], None)
        == "Load Sheet"
    )
    assert bol_generator._default_worksheet_selection(["Rates"], None) == "Rates"


def test_clear_worksheet_dependent_state_clears_parsed_and_generated_state_cheaply(
    monkeypatch,
) -> None:
    def fail_if_called(*args, **kwargs):
        raise AssertionError("Worksheet changes must not recursively delete generated files.")

    monkeypatch.setattr(bol_generator, "_cleanup_generation_output_dirs", fail_if_called)

    bol_generator.st.session_state["bol_parse_requested"] = True
    bol_generator.st.session_state["bol_parse_error"] = "old error"
    bol_generator.st.session_state["bol_parsed_rows"] = [object()]
    bol_generator.st.session_state["bol_grouped_records"] = [object()]
    bol_generator.st.session_state["bol_parsed_worksheet"] = "Load Sheet"
    bol_generator.st.session_state["bol_record_comments"] = {"BOL-1": "comment"}
    bol_generator.st.session_state["bol_record_selection"] = {"BOL-1": True}
    bol_generator.st.session_state["bol_docx_result"] = object()
    bol_generator.st.session_state["bol_pdf_result"] = object()
    bol_generator.st.session_state["bol_pdf_source_signature"] = object()
    bol_generator.st.session_state["bol_bundle_result"] = object()
    bol_generator.st.session_state["bol_bundle_error"] = "old bundle error"
    bol_generator.st.session_state["bol_all_files_bundle_requested"] = True

    bol_generator._clear_worksheet_dependent_state()

    assert bol_generator.st.session_state["bol_parse_requested"] is False
    assert bol_generator.st.session_state["bol_parse_error"] is None
    assert bol_generator.st.session_state["bol_parsed_rows"] == []
    assert bol_generator.st.session_state["bol_grouped_records"] == []
    assert bol_generator.st.session_state["bol_parsed_worksheet"] is None
    assert bol_generator.st.session_state["bol_record_comments"] == {}
    assert bol_generator.st.session_state["bol_record_selection"] == {}
    assert bol_generator.st.session_state["bol_docx_result"] is None
    assert bol_generator.st.session_state["bol_pdf_result"] is None
    assert bol_generator.st.session_state["bol_pdf_source_signature"] is None
    assert bol_generator.st.session_state["bol_bundle_result"] is None
    assert bol_generator.st.session_state["bol_bundle_error"] is None
    assert bol_generator.st.session_state["bol_all_files_bundle_requested"] is False


def test_parse_summary_reports_selected_standard_worksheet() -> None:
    bol_generator.st.session_state["bol_parsed_worksheet"] = "Revised LS"

    assert bol_generator._summary_worksheet_label("Excel upload", "Standard") == "Revised LS"
    assert (
        bol_generator._summary_worksheet_label("Excel upload", "No Recourse")
        == "Revised LS"
    )


def test_summary_worksheet_label_handles_missing_state_without_parse_local() -> None:
    bol_generator.st.session_state.pop("bol_parsed_worksheet", None)
    bol_generator.st.session_state["bol_selected_worksheet"] = "Load Sheet"

    assert bol_generator._summary_worksheet_label("Excel upload", "Standard") == "Load Sheet"
    assert bol_generator._summary_worksheet_label("Doc upload", "Standard") == "N/A"


def test_pdf_generation_uses_grouped_records_from_selected_parse(monkeypatch) -> None:
    selected_parse_records = [object()]
    captured: dict[str, object] = {}

    def fake_stamp_bol_pdf_set(records, **kwargs):
        captured["records"] = records
        captured["kwargs"] = kwargs
        return "pdf-result"

    monkeypatch.setattr(bol_generator, "stamp_bol_pdf_set", fake_stamp_bol_pdf_set)
    bol_generator.st.session_state["bol_selected_facility"] = {"facility": "TEST"}
    bol_generator.st.session_state["bol_type_selector"] = "PLT"
    bol_generator.st.session_state["bol_qty_type_selector"] = "PLT"
    bol_generator.st.session_state["bol_batch_comment_textarea"] = ""
    bol_generator.st.session_state["bol_render_pickup_number"] = "No"

    result = bol_generator._generate_pdf_result(
        mode="Standard",
        docx_result=SimpleNamespace(generated_files=[]),
        grouped_records=selected_parse_records,
        progress_callback=None,
    )

    assert result == "pdf-result"
    assert captured["records"] is selected_parse_records
    assert captured["kwargs"]["render_pickup_number"] is False


def test_changing_selected_facility_updates_grouped_records_and_clears_generated_state() -> None:
    record = SimpleNamespace(
        ship_from=BolAddressBlock(
            company="Old",
            street="Old",
            city_state_zip="Old",
        )
    )
    bol_generator.st.session_state["bol_selected_facility_label"] = "SHORR"
    bol_generator.st.session_state["bol_selected_facility"] = None
    bol_generator.st.session_state["bol_grouped_records"] = [record]
    bol_generator.st.session_state["bol_docx_result"] = object()
    bol_generator.st.session_state["bol_pdf_result"] = object()
    bol_generator.st.session_state["bol_pdf_source_signature"] = object()
    bol_generator.st.session_state["bol_bundle_result"] = object()
    bol_generator.st.session_state["bol_bundle_error"] = "old"
    bol_generator.st.session_state["bol_all_files_bundle_requested"] = True

    bol_generator._set_selected_facility("PRODUCTIV-ESTERS")

    assert record.ship_from.company == "Kendal King C/O Productiv"
    assert record.ship_from.street == "2450 Esters BLVD Suite 100"
    assert record.ship_from.city_state_zip == "Grapevine, TX 76051"
    assert bol_generator.st.session_state["bol_docx_result"] is None
    assert bol_generator.st.session_state["bol_pdf_result"] is None
    assert bol_generator.st.session_state["bol_pdf_source_signature"] is None
    assert bol_generator.st.session_state["bol_bundle_result"] is None
    assert bol_generator.st.session_state["bol_bundle_error"] is None
    assert bol_generator.st.session_state["bol_all_files_bundle_requested"] is False


def test_importing_bol_generator_does_not_create_pdf_readers(monkeypatch) -> None:
    import pypdf

    original_stamper = sys.modules.pop("app.services.bol_pdf_template_stamper", None)
    original_ui = sys.modules.pop("app.ui.bol_generator", None)

    def fail_if_called(*args, **kwargs):
        raise AssertionError("Import must not open PDF templates or create PDF readers.")

    monkeypatch.setattr(pypdf, "PdfReader", fail_if_called)
    try:
        imported = importlib.import_module("app.ui.bol_generator")
        assert imported is not None
    finally:
        sys.modules.pop("app.ui.bol_generator", None)
        sys.modules.pop("app.services.bol_pdf_template_stamper", None)
        if original_stamper is not None:
            sys.modules["app.services.bol_pdf_template_stamper"] = original_stamper
        if original_ui is not None:
            sys.modules["app.ui.bol_generator"] = original_ui


def test_importing_pdf_stamper_does_not_read_template_files(monkeypatch) -> None:
    import pypdf

    original_stamper = sys.modules.pop("app.services.bol_pdf_template_stamper", None)

    def fail_if_called(*args, **kwargs):
        raise AssertionError("Import must not open PDF templates or create PDF readers.")

    monkeypatch.setattr(pypdf, "PdfReader", fail_if_called)
    try:
        imported = importlib.import_module("app.services.bol_pdf_template_stamper")
        assert imported is not None
    finally:
        sys.modules.pop("app.services.bol_pdf_template_stamper", None)
        if original_stamper is not None:
            sys.modules["app.services.bol_pdf_template_stamper"] = original_stamper
