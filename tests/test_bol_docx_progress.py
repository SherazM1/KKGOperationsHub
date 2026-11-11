"""Word progress counts actual documents and handles skipped or failed work."""

from copy import deepcopy
from types import SimpleNamespace

import pytest

from app.services import bol_multistop_docx_generator as multistop
from app.services.bol_multistop_mapper import map_multistop_rows_to_records
from app.services.bol_standard_docx_generator import (
    generate_standard_docx_set,
    resolve_template_path_for_mode,
)
from app.ui import bol_generator
from app.utils.bol_facilities import BOL_FACILITY_LOOKUP
from tests.test_bol_multistop_grouping import _fake_multistop_saves, _row
from tests.test_bol_standard_docx_generator import _ready_record


FACILITY = BOL_FACILITY_LOOKUP["Green Bay Packaging"]


@pytest.mark.parametrize("mode", ["Standard", "No Recourse"])
@pytest.mark.parametrize("allow_incomplete", [False, True])
def test_standard_progress_counts_only_eligible_documents_and_updates_after_save(tmp_path, mode, allow_incomplete):
    ready = _ready_record()
    excluded, blocked, incomplete = (deepcopy(ready) for _ in range(3))
    excluded.selected_for_generation = False
    blocked.import_issues = ["Invalid quantity"]
    incomplete.is_ready = False
    events = []

    def progress(completed, total, detail):
        events.append((completed, total, detail))
        if completed:
            assert len(list(tmp_path.glob("*.docx"))) == completed

    result = generate_standard_docx_set(
        [ready, excluded, blocked, incomplete], FACILITY, output_dir=tmp_path,
        template_path=resolve_template_path_for_mode(mode), progress_callback=progress,
        allow_incomplete_records=allow_incomplete,
    )
    expected = 2 if allow_incomplete else 1
    assert result.generated_count == expected
    assert result.skipped_count == 4 - expected
    assert all(total == expected for _, total, _ in events)
    assert events[0][:2] == events[1][:2] == (0, expected)
    assert events[-1][:2] == (expected, expected)
    assert "Generating BOL" in events[1][2]
    assert "Created BOL" in events[-1][2]


def test_standard_progress_reports_failure_without_stopping_the_batch(tmp_path, monkeypatch):
    first, second = _ready_record(), _ready_record()
    second.bol_number = "SECOND"
    from app.services import bol_standard_docx_generator as standard
    original = standard.Document
    attempts = 0

    def fail_first(*args, **kwargs):
        nonlocal attempts
        attempts += 1
        if attempts == 1:
            raise OSError("Test save failure")
        return original(*args, **kwargs)

    monkeypatch.setattr(standard, "Document", fail_first)
    events = []
    result = generate_standard_docx_set(
        [first, second], FACILITY, output_dir=tmp_path,
        progress_callback=lambda *event: events.append(event),
    )
    assert result.generated_count == result.failed_count == 1
    assert any(done == 1 and "Failed BOL" in detail for done, _, detail in events)
    assert events[-1][:2] == (2, 2)


def test_multistop_progress_counts_master_and_each_child(tmp_path, monkeypatch):
    _fake_multistop_saves(monkeypatch)
    records = map_multistop_rows_to_records([
        _row(kk_load="1", stop=number, bol_number=f"BOL-{number}") for number in range(1, 6)
    ])
    excluded = deepcopy(records[0])
    excluded.selected_for_generation = False
    blocked = deepcopy(records[0])
    blocked.import_issues = ["Invalid quantity"]
    events = []

    def progress(completed, total, detail):
        events.append((completed, total, detail))
        assert len(list(tmp_path.glob("*.docx"))) == completed

    result = multistop.generate_multistop_docx_set(
        [*records, excluded, blocked], FACILITY, output_dir=tmp_path,
        progress_callback=progress,
    )
    assert result.generated_count == 6
    assert result.skipped_count == 2
    assert all(total == 6 for _, total, _ in events)
    assert events[-1][:2] == (6, 6)
    assert any("Generating master BOL" in detail for _, _, detail in events)
    assert any("Generating stop 5, BOL BOL-5" in detail for _, _, detail in events)


def test_multistop_progress_accounts_for_abandoned_children_after_failure(tmp_path, monkeypatch):
    _fake_multistop_saves(monkeypatch)
    save_stop = multistop._save_individual_stop_docx

    def fail_second_stop(**kwargs):
        if kwargs["record"].kk_load_number == "1" and kwargs["stop_index"] == 1:
            raise OSError("Test child failure")
        return save_stop(**kwargs)

    monkeypatch.setattr(multistop, "_save_individual_stop_docx", fail_second_stop)
    records = map_multistop_rows_to_records([
        _row(kk_load=str(load), stop=number, bol_number=f"BOL-{load}-{number}")
        for load in (1, 2) for number in (1, 2, 3)
    ])
    events = []
    result = multistop.generate_multistop_docx_set(
        records, FACILITY, output_dir=tmp_path,
        progress_callback=lambda *event: events.append(event),
    )
    assert result.generated_count == 6
    assert result.failed_count == 1
    assert any(done == 4 and "remaining documents skipped" in detail
               for done, _, detail in events)
    assert events[-1][:2] == (8, 8)
    assert [done for done, _, _ in events] == sorted(done for done, _, _ in events)


@pytest.mark.parametrize("fail", [False, True])
def test_word_progress_ui_updates_and_cleans_up_on_success_or_error(monkeypatch, fail):
    events = []
    bar = SimpleNamespace(
        progress=lambda value: events.append(("progress", value)),
        empty=lambda: events.append(("clear_bar",)),
    )
    status = SimpleNamespace(
        write=lambda text: events.append(("status", text)),
        empty=lambda: events.append(("clear_status",)),
    )
    monkeypatch.setattr(bol_generator.st, "progress", lambda initial: bar)
    monkeypatch.setattr(bol_generator.st, "empty", lambda: status)
    expected = object()

    def generate(records, *, progress_callback):
        assert records == ["record"]
        progress_callback(0, 2, "Generating master BOL")
        progress_callback(1, 2, "Generating stop 1")
        if fail:
            raise OSError("Generation failed")
        progress_callback(2, 2, "Created stop 1")
        return expected

    if fail:
        with pytest.raises(OSError, match="Generation failed"):
            bol_generator._generate_docx_with_progress(generate, ["record"])
    else:
        assert bol_generator._generate_docx_with_progress(generate, ["record"]) is expected
        assert ("progress", 1) in events
    assert ("progress", 0.5) in events
    assert any("1 of 2" in event[1] for event in events if event[0] == "status")
    assert events[-2:] == [("clear_bar",), ("clear_status",)]
