from __future__ import annotations

import importlib
import sys

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
