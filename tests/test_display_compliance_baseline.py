"""Tests for the Display Compliance baseline scaffold."""

from __future__ import annotations

from pathlib import Path
import struct
import zlib

import pytest

from app.display_compliance.baseline import create_baseline
from app.display_compliance.models import DisplayBaseline, ProductRegion
from app.display_compliance import page as display_page
from app.display_compliance.storage import LocalBaselineStorage, baseline_from_dict


def _png_bytes(width: int = 3, height: int = 2) -> bytes:
    def chunk(chunk_type: bytes, data: bytes) -> bytes:
        checksum = zlib.crc32(chunk_type + data) & 0xFFFFFFFF
        return struct.pack(">I", len(data)) + chunk_type + data + struct.pack(">I", checksum)

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    raw_rows = b"".join(b"\x00" + (b"\xff\x00\x00" * width) for _ in range(height))
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(
        b"IDAT", zlib.compress(raw_rows)
    ) + chunk(b"IEND", b"")


def test_create_baseline_from_valid_image_records_domain_metadata() -> None:
    baseline = create_baseline(
        name="Endcap Reference",
        filename="reference.png",
        image_bytes=_png_bytes(7, 5),
    )

    assert isinstance(baseline, DisplayBaseline)
    assert baseline.name == "Endcap Reference"
    assert baseline.reference_filename == "reference.png"
    assert baseline.reference_width == 7
    assert baseline.reference_height == 5
    assert baseline.regions == []
    assert baseline.baseline_id


def test_blank_baseline_name_is_rejected() -> None:
    with pytest.raises(ValueError, match="name is required"):
        create_baseline(name="  ", filename="reference.png", image_bytes=_png_bytes())


@pytest.mark.parametrize("image_bytes", [b"", b"not an image"])
def test_empty_or_invalid_image_bytes_are_rejected(image_bytes: bytes) -> None:
    with pytest.raises(ValueError):
        create_baseline(name="Reference", filename="reference.png", image_bytes=image_bytes)


def test_baseline_metadata_round_trips_through_local_storage(tmp_path: Path) -> None:
    image_bytes = _png_bytes(4, 6)
    baseline = create_baseline(
        name="Baseline A",
        filename="baseline a.png",
        image_bytes=image_bytes,
    )
    storage = LocalBaselineStorage(tmp_path)

    storage.save_baseline(baseline, image_bytes)
    loaded = storage.load_baseline(baseline.baseline_id)

    assert loaded == baseline
    assert storage.list_baselines() == [baseline]
    assert (tmp_path / baseline.baseline_id / "baseline.json").exists()
    assert (tmp_path / baseline.baseline_id / "baseline_a.png").read_bytes() == image_bytes


def test_baseline_region_metadata_round_trips_through_local_storage(tmp_path: Path) -> None:
    image_bytes = _png_bytes(10, 8)
    baseline = create_baseline(
        name="Baseline With Regions",
        filename="baseline.png",
        image_bytes=image_bytes,
    )
    baseline.regions.append(
        ProductRegion(
            region_id="region_001",
            bbox=(1, 2, 3, 4),
            polygon=((1, 2), (4, 2), (4, 6), (1, 6)),
        )
    )
    storage = LocalBaselineStorage(tmp_path)

    storage.save_baseline(baseline, image_bytes)
    loaded = storage.load_baseline(baseline.baseline_id)

    assert loaded == baseline


def test_existing_zero_region_baseline_metadata_remains_loadable() -> None:
    baseline = baseline_from_dict(
        {
            "baseline_id": "baseline-1",
            "name": "Pass 1 Baseline",
            "reference_filename": "reference.png",
            "reference_width": 100,
            "reference_height": 80,
            "regions": [],
        }
    )

    assert baseline.regions == []


class _StopRender(Exception):
    pass


class _FakeDisplayComplianceStorage:
    def __init__(self, baselines: list[DisplayBaseline]) -> None:
        self._baselines = baselines
        self.saved_metadata: list[DisplayBaseline] = []
        self.loaded_images: list[str] = []

    def list_baselines(self) -> list[DisplayBaseline]:
        return self._baselines

    def load_baseline(self, baseline_id: str) -> DisplayBaseline:
        for baseline in self._baselines:
            if baseline.baseline_id == baseline_id:
                return baseline
        raise FileNotFoundError(baseline_id)

    def load_reference_image_bytes(self, baseline: DisplayBaseline) -> bytes:
        self.loaded_images.append(baseline.baseline_id)
        return _png_bytes(10, 8)

    def save_baseline_metadata(self, baseline: DisplayBaseline) -> None:
        self.saved_metadata.append(baseline)


class _FakeStreamlit:
    def __init__(
        self,
        *,
        session_state: dict[str, object] | None = None,
        clicked_buttons: set[str] | None = None,
    ) -> None:
        self.session_state = session_state or {}
        self.clicked_buttons = clicked_buttons or set()
        self.selectbox_calls: list[dict[str, object]] = []
        self.button_calls: list[dict[str, object]] = []
        self.info_messages: list[str] = []
        self.errors: list[str] = []

    def button(self, label: str, **kwargs: object) -> bool:
        key = str(kwargs.get("key", label))
        self.button_calls.append({"label": label, **kwargs})
        return key in self.clicked_buttons and not bool(kwargs.get("disabled", False))

    def selectbox(self, label: str, *, options: list[object], key: str, **kwargs: object) -> object:
        self.selectbox_calls.append({"label": label, "options": options, "key": key, **kwargs})
        value = self.session_state.get(key)
        if value not in options:
            value = options[0]
            self.session_state[key] = value
        return value

    def text_input(self, *args: object, **kwargs: object) -> str:
        return ""

    def file_uploader(self, *args: object, **kwargs: object) -> None:
        return None

    def dataframe(self, *args: object, **kwargs: object) -> None:
        return None

    def image(self, *args: object, **kwargs: object) -> None:
        return None

    def metric(self, *args: object, **kwargs: object) -> None:
        return None

    def title(self, *args: object, **kwargs: object) -> None:
        return None

    def subheader(self, *args: object, **kwargs: object) -> None:
        return None

    def caption(self, *args: object, **kwargs: object) -> None:
        return None

    def success(self, *args: object, **kwargs: object) -> None:
        return None

    def warning(self, *args: object, **kwargs: object) -> None:
        return None

    def error(self, message: str) -> None:
        self.errors.append(message)

    def info(self, message: str) -> None:
        self.info_messages.append(message)

    def stop(self) -> None:
        raise _StopRender


def _baseline_for_page(baseline_id: str = "baseline-1") -> DisplayBaseline:
    return DisplayBaseline(
        baseline_id=baseline_id,
        name="Endcap Reference",
        reference_filename="reference.png",
        reference_width=10,
        reference_height=8,
        regions=[],
    )


def _render_page_with_fakes(
    monkeypatch: pytest.MonkeyPatch,
    *,
    baselines: list[DisplayBaseline],
    session_state: dict[str, object] | None = None,
    clicked_buttons: set[str] | None = None,
) -> tuple[_FakeStreamlit, _FakeDisplayComplianceStorage]:
    fake_st = _FakeStreamlit(session_state=session_state, clicked_buttons=clicked_buttons)
    fake_storage = _FakeDisplayComplianceStorage(baselines)
    monkeypatch.setattr(display_page, "st", fake_st)
    monkeypatch.setattr(display_page, "LocalBaselineStorage", lambda: fake_storage)
    monkeypatch.setattr(display_page, "detect_baseline_regions", lambda **kwargs: [])
    monkeypatch.setattr(display_page, "render_annotated_preview", lambda **kwargs: b"preview")

    display_page.render_display_compliance_view()

    return fake_st, fake_storage


def test_display_compliance_renders_when_no_saved_baselines_exist(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    fake_st, fake_storage = _render_page_with_fakes(monkeypatch, baselines=[])

    assert fake_st.selectbox_calls == []
    assert "No saved baselines yet. Create a baseline above to begin." in fake_st.info_messages
    assert fake_storage.loaded_images == []


def test_display_compliance_placeholder_selection_does_not_lookup_or_detect(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    fake_st, fake_storage = _render_page_with_fakes(
        monkeypatch,
        baselines=[_baseline_for_page()],
        session_state={"display_compliance_selected_baseline_id": None},
    )

    assert fake_st.selectbox_calls[0]["options"][0] is None
    assert "Select a saved baseline to detect candidate product regions." in fake_st.info_messages
    assert fake_storage.loaded_images == []


def test_display_compliance_uses_valid_baseline_id_selection(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    baseline = _baseline_for_page()
    fake_st, fake_storage = _render_page_with_fakes(
        monkeypatch,
        baselines=[baseline],
        session_state={"display_compliance_selected_baseline_id": baseline.baseline_id},
    )

    assert fake_st.selectbox_calls[0]["options"] == [None, baseline.baseline_id]
    assert any(call["key"] == "display_compliance_detect_product_regions" for call in fake_st.button_calls)
    assert fake_storage.loaded_images == []


def test_display_compliance_sanitizes_stale_selected_baseline_before_selectbox(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    baseline = _baseline_for_page()
    fake_st, fake_storage = _render_page_with_fakes(
        monkeypatch,
        baselines=[baseline],
        session_state={
            "display_compliance_selected_baseline_id": "missing-baseline",
            "display_compliance_annotated_preview_baseline_id": "missing-baseline",
            "display_compliance_annotated_preview_bytes": b"old-preview",
        },
    )

    assert fake_st.session_state.get("display_compliance_selected_baseline_id") is None
    assert "display_compliance_annotated_preview_baseline_id" not in fake_st.session_state
    assert "display_compliance_annotated_preview_bytes" not in fake_st.session_state
    assert fake_storage.loaded_images == []


def test_display_compliance_valid_selection_permits_detection_and_rerun(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    baseline = _baseline_for_page()
    fake_st, fake_storage = _render_page_with_fakes(
        monkeypatch,
        baselines=[baseline],
        session_state={"display_compliance_selected_baseline_id": baseline.baseline_id},
        clicked_buttons={"display_compliance_detect_product_regions"},
    )

    assert fake_storage.loaded_images == [baseline.baseline_id]
    assert fake_storage.saved_metadata
    assert (
        fake_st.session_state["display_compliance_annotated_preview_baseline_id"]
        == baseline.baseline_id
    )


def test_display_compliance_back_to_home_still_stops_render(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    fake_st = _FakeStreamlit(clicked_buttons={"display_compliance_back_button"})
    monkeypatch.setattr(display_page, "st", fake_st)
    monkeypatch.setattr(display_page, "LocalBaselineStorage", lambda: _FakeDisplayComplianceStorage([]))

    with pytest.raises(_StopRender):
        display_page.render_display_compliance_view()

    assert fake_st.session_state["page"] == "home"
