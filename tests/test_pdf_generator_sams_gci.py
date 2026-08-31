"""Tests for Sam's GCI PDF layout decisions."""

from __future__ import annotations

from reportlab.pdfbase import pdfmetrics

from app.services.pdf_generator_sams_gci import _draw_bottom_row_box


class RecordingCanvas:
    def __init__(self) -> None:
        self.drawn_strings: list[tuple[str, float, float, str]] = []
        self.right_strings: list[tuple[str, float, float, str]] = []
        self.current_font = "Helvetica"
        self.current_font_size = 8.0

    def setFont(self, font_name: str, font_size: float) -> None:
        self.current_font = font_name
        self.current_font_size = font_size

    def stringWidth(self, text: str, font_name: str, font_size: float) -> float:
        return pdfmetrics.stringWidth(text, font_name, font_size)

    def drawString(self, x: float, y: float, text: str) -> None:
        self.drawn_strings.append((text, x, y, self.current_font))

    def drawRightString(self, x: float, y: float, text: str) -> None:
        self.right_strings.append((text, x, y, self.current_font))


def test_bottom_row_moves_qty_below_long_item_number_to_prevent_overlap() -> None:
    canvas = RecordingCanvas()

    _draw_bottom_row_box(
        canvas,  # type: ignore[arg-type]
        {
            "item_number": "990553354",
            "quantity": "12",
            "description": "2Pk Boxes Trees",
            "barcode_value": "",
        },
        row_top=200,
        row_bottom=160,
        text_left_x=12,
        text_right_x=82,
        barcode_left_x=90,
        barcode_right_x=250,
        barcode_cache={},
        wrap_cache={},
    )

    item_draw = next(call for call in canvas.drawn_strings if call[0] == "ITEM#: 990553354")
    qty_draw = next(call for call in canvas.drawn_strings if call[0] == "QTY: 12")

    assert not any(call[0] == "QTY: 12" for call in canvas.right_strings)
    assert qty_draw[2] < item_draw[2]
