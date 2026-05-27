"""Tests for Andersons Label Maker session behavior."""

from __future__ import annotations

from types import SimpleNamespace

from app import main


def test_default_andersons_worksheet_selection_prefers_previous_then_named_defaults() -> None:
    sheet_names = ["Tracker Info", "Revised LS", "Load Sheet"]

    assert (
        main._default_andersons_worksheet_selection(sheet_names, "Load Sheet")
        == "Load Sheet"
    )
    assert (
        main._default_andersons_worksheet_selection(sheet_names, "Missing")
        == "Revised LS"
    )
    assert (
        main._default_andersons_worksheet_selection(["Tracker Info", "Load Sheet"], None)
        == "Load Sheet"
    )
    assert main._default_andersons_worksheet_selection(["Rates"], None) == "Rates"


def test_andersons_session_stores_selected_worksheet() -> None:
    main.st.session_state.clear()
    main._initialize_andersons_state()

    main._handle_andersons_worksheet_change("Revised LS")

    assert main.st.session_state["andersons_selected_worksheet"] == "Revised LS"
    assert main.st.session_state["andersons_labels"] == []
    assert main.st.session_state["andersons_parsed_worksheet"] is None
    assert main.st.session_state["andersons_pdf_bytes"] is None


def test_changing_andersons_worksheet_clears_stale_andersons_generated_state_only() -> None:
    main.st.session_state.clear()
    main._initialize_andersons_state()
    main.st.session_state["andersons_selected_worksheet"] = "Load Sheet"
    main.st.session_state["andersons_labels"] = [object()]
    main.st.session_state["andersons_parsed_worksheet"] = "Load Sheet"
    main.st.session_state["andersons_parse_error"] = "old error"
    main.st.session_state["andersons_pdf_bytes"] = b"old pdf"
    main.st.session_state["bol_grouped_records"] = [object()]
    main.st.session_state["skid_tags_rows"] = [object()]

    main._handle_andersons_worksheet_change("Revised LS")

    assert main.st.session_state["andersons_selected_worksheet"] == "Revised LS"
    assert main.st.session_state["andersons_labels"] == []
    assert main.st.session_state["andersons_parsed_worksheet"] is None
    assert main.st.session_state["andersons_parse_error"] is None
    assert main.st.session_state["andersons_pdf_bytes"] is None
    assert len(main.st.session_state["bol_grouped_records"]) == 1
    assert len(main.st.session_state["skid_tags_rows"]) == 1


def test_changing_andersons_upload_clears_parsed_generated_and_selected_state() -> None:
    main.st.session_state.clear()
    main._initialize_andersons_state()
    main.st.session_state["andersons_uploaded_file_signature"] = ("old-id", "old.xlsx", 100)
    main.st.session_state["andersons_selected_worksheet"] = "Load Sheet"
    main.st.session_state["andersons_selected_worksheet_selectbox"] = "Load Sheet"
    main.st.session_state["andersons_labels"] = [object()]
    main.st.session_state["andersons_parsed_worksheet"] = "Load Sheet"
    main.st.session_state["andersons_pdf_bytes"] = b"old pdf"

    new_upload = SimpleNamespace(file_id="new-id", name="new.xlsx", size=200)
    main._handle_andersons_file_change(new_upload)

    assert main.st.session_state["andersons_uploaded_file_signature"] == (
        "new-id",
        "new.xlsx",
        200,
    )
    assert main.st.session_state["andersons_selected_worksheet"] is None
    assert "andersons_selected_worksheet_selectbox" not in main.st.session_state
    assert main.st.session_state["andersons_labels"] == []
    assert main.st.session_state["andersons_parsed_worksheet"] is None
    assert main.st.session_state["andersons_pdf_bytes"] is None
