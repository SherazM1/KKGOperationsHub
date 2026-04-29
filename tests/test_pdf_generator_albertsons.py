"""Tests for Albertsons PDF generation."""

from __future__ import annotations

import pytest

from app.models.albertsons_label import AlbertsonsLabel
from app.services import pdf_generator_albertsons
from app.services.pdf_generator_albertsons import _draw_label_page, generate_albertsons_pdf


class RecordingCanvas:
    def __init__(self) -> None:
        self.strings: list[str] = []
        self.page_count = 0

    def setFillColorRGB(self, *_args: object) -> None:
        pass

    def setFont(self, *_args: object) -> None:
        pass

    def drawCentredString(self, _x: float, _y: float, text: str) -> None:
        self.strings.append(text)

    def drawRightString(self, _x: float, _y: float, text: str) -> None:
        self.strings.append(text)

    def drawString(self, _x: float, _y: float, text: str) -> None:
        self.strings.append(text)

    def setLineWidth(self, *_args: object) -> None:
        pass

    def line(self, *_args: object) -> None:
        pass

    def showPage(self) -> None:
        self.page_count += 1

    def save(self) -> None:
        pass


def _label() -> AlbertsonsLabel:
    return AlbertsonsLabel(
        ship_to_name="Store",
        ship_to_address="123 Main St",
        ship_to_city="Dallas",
        ship_to_state="TX",
        ship_to_zip="75001",
        po_number="PO-1",
        item_number="EXCEL-ITEM",
        description="Display",
        quantity="12",
        dc_label="DC#",
        dc_value="WNCA",
        carton_number="1",
    )


def test_albertsons_manual_values_override_label_values() -> None:
    canvas = RecordingCanvas()

    _draw_label_page(
        canvas,
        _label(),
        manual_item_number="MANUAL-ITEM",
        manual_qty="24",
        manual_po_type="PO-TYPE",
    )

    assert "MANUAL-ITEM" in canvas.strings
    assert "Qty 24" in canvas.strings
    assert "PO-TYPE" in canvas.strings
    assert "EXCEL-ITEM" not in canvas.strings
    assert "Qty 12" not in canvas.strings


def test_albertsons_auto_qty_uses_label_quantity() -> None:
    canvas = RecordingCanvas()

    _draw_label_page(canvas, _label(), qty_mode="auto")

    assert "Qty 12" in canvas.strings


def test_albertsons_blank_manual_values_preserve_non_qty_defaults() -> None:
    canvas = RecordingCanvas()

    _draw_label_page(canvas, _label())

    assert "EXCEL-ITEM" in canvas.strings
    assert "Qty " in canvas.strings
    assert "1" in canvas.strings


def test_albertsons_pdf_generates_two_pages_per_label(monkeypatch: pytest.MonkeyPatch) -> None:
    canvas = RecordingCanvas()
    monkeypatch.setattr(pdf_generator_albertsons.canvas, "Canvas", lambda *_args, **_kwargs: canvas)

    generate_albertsons_pdf([_label()], qty_mode="auto")

    assert canvas.page_count == 2
    assert canvas.strings.count("Qty 12") == 2
