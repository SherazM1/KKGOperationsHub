"""PDF template stamping service for BOL outputs."""

from __future__ import annotations

from dataclasses import dataclass, replace
from io import BytesIO
from pathlib import Path
import re
from tempfile import mkdtemp
from typing import Any, Callable

from pypdf import PdfReader, PdfWriter
from pypdf.generic import ArrayObject, ContentStream, NameObject, TextStringObject
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

from app.models.bol_multistop_record import BolMultistopRecord
from app.models.bol_standard_record import BolStandardItemLine, BolStandardRecord
from app.services.bol_multistop_docx_generator import _format_multistop_item_description
from app.services.bol_standard_docx_generator import (
    GeneratedDocxFile,
    _format_number,
    _format_ship_date_for_template,
    _normalize_bol_type,
    _parse_numeric,
    _qty_type_header,
)
from app.services.bol_standard_pdf_converter import (
    ConvertedPdfFile,
    FailedPdfConversion,
    StandardPdfConversionResult,
)
from app.utils.bol_facilities import BolFacilityRecord, facility_to_ship_from


STANDARD_PDF_TEMPLATE_PATH = Path("app/templates/standard_pdf_template.pdf")
NO_RECOURSE_PDF_TEMPLATE_PATH = Path("app/templates/no_recourse_pdf_template.pdf")
MULTISTOP_PDF_TEMPLATE_PATH = Path("app/templates/multistop_pdf_template.pdf")

PLACEHOLDER_TEXT_PATTERN = re.compile(r"\u00ab[^\u00bb]+\u00bb")
KNOWN_PLACEHOLDER_TOKENS = frozenset(
    {
        "QTY_2",
        "QTY_3",
        "QTY_4",
        "TYPE_2",
        "TYPE_3",
        "TYPE_4",
        "PO_2",
        "PO_3",
        "PO_4",
        "WEIGHT_2",
        "WEIGHT_3",
        "WEIGHT_4",
        "SHIP_TO_CITY_STATE_ZIP",
        "BILL_TO",
        "BILL_TO_ADDRESS",
        "BILL_TO_CITY_STATE_ZIP",
        "TOTAL_QTY",
        "ITEM_2",
        "ITEM_3",
        "ITEM_4",
        "UPC_2",
        "UPC_3",
        "UPC_4",
        "Item #:",
        "UPC #:",
        "Item",
        "UPC",
        "#:",
        "TOTALS",
        "Pallet Qty",
        "Case Qty",
        "QTY",
        "Pallet",
        "Case",
        "Qty",
    }
)

PAGE_WIDTH, PAGE_HEIGHT = letter
FONT_NAME = "Helvetica"
FONT_BOLD = "Helvetica-Bold"

GRID_WIDTHS = [
    1710,
    743,
    67,
    169,
    1531,
    1617,
    454,
    250,
    1806,
    412,
    421,
    677,
    43,
    665,
    74,
    434,
    74,
    473,
    74,
]
COL_WIDTHS = [value / 20 for value in GRID_WIDTHS]
TABLE_WIDTH = sum(COL_WIDTHS)
TABLE_X = (PAGE_WIDTH - TABLE_WIDTH) / 2
TABLE_TOP = PAGE_HEIGHT - 5
ROW_HEIGHTS = [
    446,
    287,
    270,
    270,
    261,
    288,
    269,
    279,
    237,
    279,
    279,
    173,
    381,
    272,
    372,
    338,
    325,
    310,
    337,
    365,
    356,
    551,
    2009,
    263,
    519,
    310,
    488,
    310,
    316,
    316,
    310,
    310,
    327,
    704,
    353,
]
ROW_HEIGHTS_PT = [value / 20 for value in ROW_HEIGHTS]


@dataclass(frozen=True, slots=True)
class TextBox:
    x: float
    y: float
    width: float
    height: float
    font_size: float = 8.0
    min_font_size: float = 5.0
    bold: bool = False
    align: str = "left"
    multiline: bool = False
    leading: float | None = None
    whiteout: bool = True
    vertical_align: str = "top"


@dataclass(frozen=True, slots=True)
class PdfTemplateConfig:
    mode: str
    template_path: Path
    fields: dict[str, TextBox]
    item_columns: dict[str, TextBox]
    item_start_y: float
    item_row_height: float
    max_item_rows: int
    totals: dict[str, TextBox]
    item_row_baselines: tuple[float, ...] = ()


def _col_x(col_index: int) -> float:
    return TABLE_X + sum(COL_WIDTHS[:col_index])


def _col_width(start_col: int, end_col: int) -> float:
    return sum(COL_WIDTHS[start_col:end_col])


def _row_top(row_index: int) -> float:
    return TABLE_TOP - sum(ROW_HEIGHTS_PT[:row_index])


def _row_bottom(row_index: int) -> float:
    return _row_top(row_index) - ROW_HEIGHTS_PT[row_index]


def _row_span_bottom(row_start: int, row_end: int) -> float:
    return TABLE_TOP - sum(ROW_HEIGHTS_PT[:row_end])


def clean_value(value: Any) -> str:
    text = str(value or "").strip()
    if text.startswith("\u00ab") and text.endswith("\u00bb"):
        return ""
    if "\u00ab" in text or "\u00bb" in text:
        return ""
    return text


def _safe_text(value: Any) -> str:
    return clean_value(value)


def _without_whiteout(box: TextBox) -> TextBox:
    return replace(box, whiteout=False)


def _without_whiteout_map(boxes: dict[str, TextBox]) -> dict[str, TextBox]:
    return {name: _without_whiteout(box) for name, box in boxes.items()}


def _box_for_baseline(
    *,
    x: float,
    baseline: float,
    width: float,
    height: float = 10.5,
    font_size: float = 8.4,
    min_font_size: float = 5.6,
    align: str = "left",
    bold: bool = False,
    multiline: bool = False,
    leading: float | None = None,
) -> TextBox:
    if multiline:
        y = baseline - height + font_size
    else:
        y = baseline - max((height - font_size) / 2, 0) - 1.2
    return TextBox(
        x=x,
        y=y,
        width=width,
        height=height,
        font_size=font_size,
        min_font_size=min_font_size,
        bold=bold,
        align=align,
        multiline=multiline,
        leading=leading,
    )


def _top_value_box(baseline: float, *, x: float = 444.4, width: float = 128.0) -> TextBox:
    return _box_for_baseline(
        x=x,
        baseline=baseline,
        width=width,
        height=10.4,
        font_size=8.4,
        min_font_size=5.4,
    )


def _no_recourse_top_value_box(baseline: float, *, x: float = 444.4, width: float = 128.0) -> TextBox:
    return _box_for_baseline(
        x=x,
        baseline=baseline,
        width=width,
        height=10.4,
        font_size=8.8,
        min_font_size=6.5,
    )


def _right_value_box(row: int, value_cols: tuple[int, int] = (9, 14)) -> TextBox:
    return TextBox(
        x=_col_x(value_cols[0]) + 2,
        y=_row_bottom(row) - 3.8,
        width=_col_width(*value_cols) - 4,
        height=ROW_HEIGHTS_PT[row] - 2,
        font_size=8.4,
        min_font_size=5.4,
    )


def _left_value_box(row: int, value_row: int | None = None) -> TextBox:
    resolved_row = value_row if value_row is not None else row
    return TextBox(
        x=_col_x(1) + 2,
        y=_row_bottom(resolved_row) + 1.2,
        width=_col_width(1, 7) - 4,
        height=ROW_HEIGHTS_PT[resolved_row] - 2,
        font_size=7.0,
        min_font_size=5.0,
    )


def _wide_value_box(col_start: int, col_end: int, row_start: int, row_end: int) -> TextBox:
    return TextBox(
        x=_col_x(col_start) + 3,
        y=_row_span_bottom(row_start, row_end) + 2,
        width=_col_width(col_start, col_end) - 6,
        height=_row_top(row_start) - _row_span_bottom(row_start, row_end) - 4,
        font_size=7.2,
        min_font_size=5.0,
        multiline=True,
    )


def _item_box(col_start: int, col_end: int, *, font_size: float = 6.5, align: str = "left") -> TextBox:
    return TextBox(
        x=_col_x(col_start) + 2,
        y=0,
        width=_col_width(col_start, col_end) - 4,
        height=0,
        font_size=font_size,
        min_font_size=5.0,
        align=align,
    )


def _standard_fields() -> dict[str, TextBox]:
    return {
        "bol_number": _top_value_box(749.4),
        "ship_date": _top_value_box(735.9),
        "carrier": _top_value_box(722.4, width=150.0),
        "carrier_pro_number": _top_value_box(709.3),
        "po_number": _top_value_box(694.9),
        "ship_from_company": _box_for_baseline(x=112.5, baseline=666.1, width=220.0, font_size=8.8, min_font_size=6.5),
        "ship_from_street": _box_for_baseline(x=112.5, baseline=629.7, width=220.0, font_size=8.8, min_font_size=6.5),
        "ship_from_city_state_zip": _box_for_baseline(x=112.5, baseline=615.2, width=220.0, font_size=8.8, min_font_size=6.5),
        "consignee_company": _box_for_baseline(x=112.5, baseline=539.4, width=210.0, font_size=8.8, min_font_size=6.5),
        "consignee_street": _box_for_baseline(x=112.5, baseline=504.3, width=210.0, font_size=8.8, min_font_size=6.5),
        "consignee_city_state_zip": _box_for_baseline(x=112.5, baseline=487.8, width=220.0, font_size=8.8, min_font_size=6.5),
        "bill_to": TextBox(
            x=365.6,
            y=425.0,
            width=200.0,
            height=49.0,
            font_size=9.0,
            min_font_size=6.5,
            multiline=True,
            leading=11.3,
            vertical_align="middle",
        ),
        "tracker_number": _top_value_box(681.0),
        "kk_po_number": _top_value_box(666.1),
        "kk_load_number": _top_value_box(653.2, width=150.0),
        "delivery_appt": _top_value_box(629.7, width=150.0),
        "appt_number": _top_value_box(606.2, width=150.0),
        "comments": _box_for_baseline(
            x=444.4,
            baseline=589.6,
            width=128.0,
            height=18.0,
            font_size=7.8,
            min_font_size=5.2,
            multiline=True,
        ),
        "seal_number": _top_value_box(573.6),
        "appointment_number": _box_for_baseline(x=112.5, baseline=454.1, width=210.0, font_size=8.5, min_font_size=6.5),
        "dc_number": _box_for_baseline(x=112.5, baseline=445.0, width=210.0, font_size=8.5, min_font_size=6.5),
    }


def _no_recourse_fields() -> dict[str, TextBox]:
    return {
        "bol_number": _no_recourse_top_value_box(743.5),
        "ship_date": _no_recourse_top_value_box(732.0),
        "carrier": _no_recourse_top_value_box(720.5, width=150.0),
        "carrier_pro_number": _no_recourse_top_value_box(709.0),
        "po_number": _no_recourse_top_value_box(697.5),
        "kk_po_number": _no_recourse_top_value_box(686.6),
        "kk_load_number": _no_recourse_top_value_box(674.7),
        "seal_number": _no_recourse_top_value_box(662.7),
        "pickup_number": _no_recourse_top_value_box(650.7, width=150.0),
        "comments": _box_for_baseline(
            x=444.4,
            baseline=626.4,
            width=128.0,
            height=18.0,
            font_size=8.2,
            min_font_size=6.5,
            multiline=True,
        ),
        "ship_from_company": _box_for_baseline(x=112.5, baseline=674.7, width=220.0, font_size=8.7, min_font_size=6.5),
        "ship_from_street": _box_for_baseline(x=112.5, baseline=650.7, width=220.0, font_size=8.7, min_font_size=6.5),
        "ship_from_city_state_zip": _box_for_baseline(x=115.0, baseline=638.7, width=217.0, font_size=8.7, min_font_size=6.5),
        "consignee_company": _box_for_baseline(x=112.5, baseline=590.4, width=210.0, font_size=8.7, min_font_size=6.5),
        "consignee_street": _box_for_baseline(x=112.5, baseline=566.4, width=210.0, font_size=8.7, min_font_size=6.5),
        "consignee_city_state_zip": _box_for_baseline(x=112.5, baseline=554.4, width=220.0, font_size=8.7, min_font_size=6.5),
        "bill_to": TextBox(
            x=398.0,
            y=499.0,
            width=176.0,
            height=45.0,
            font_size=9.3,
            min_font_size=7.4,
            multiline=True,
            leading=11.8,
            vertical_align="middle",
        ),
        "dc_number": _box_for_baseline(x=112.5, baseline=510.3, width=69.0, font_size=8.5, min_font_size=6.5),
    }


STANDARD_CONFIG = PdfTemplateConfig(
    mode="Standard",
    template_path=STANDARD_PDF_TEMPLATE_PATH,
    fields=_without_whiteout_map(_standard_fields()),
    item_columns=_without_whiteout_map(
        {
            "qty_header": _box_for_baseline(x=31.0, baseline=278.2, width=56.0, height=10.0, font_size=7.4, min_font_size=6.0, align="center"),
            "qty": TextBox(37.0, 0, 50.0, 0, 9.2, min_font_size=6.8, align="center"),
            "type": _item_box(1, 3, font_size=7.0, align="center"),
            "po": TextBox(144.0, 0, 68.0, 0, 8.4, min_font_size=6.8, align="center"),
            "description": TextBox(238.0, 0, 236.0, 0, 8.6, min_font_size=7.0, multiline=True, leading=10.6, vertical_align="middle"),
            "skids": TextBox(490.0, 0, 44.0, 0, 9.2, min_font_size=6.8, align="center"),
            "weight": TextBox(542.0, 0, 48.0, 0, 9.2, min_font_size=6.8, align="center"),
        }
    ),
    item_start_y=0,
    item_row_height=18.0,
    max_item_rows=8,
    totals=_without_whiteout_map(
        {
            "qty": _box_for_baseline(x=37.0, baseline=69.9, width=50.0, height=12.0, font_size=8.9, min_font_size=6.5, bold=True, align="center"),
            "label": _box_for_baseline(x=_col_x(5) + 2, baseline=69.9, width=_col_width(5, 11) - 4, height=12.0, font_size=8.9, min_font_size=6.5, bold=True, align="center"),
            "skids": _box_for_baseline(x=490.0, baseline=69.9, width=44.0, height=12.0, font_size=8.9, min_font_size=6.5, bold=True, align="center"),
            "weight": _box_for_baseline(x=542.0, baseline=69.9, width=48.0, height=12.0, font_size=8.9, min_font_size=6.5, bold=True, align="center"),
        }
    ),
    item_row_baselines=(264.1, 238.2, 212.3, 186.4, 160.5, 134.6, 108.7, 82.8),
)

NO_RECOURSE_CONFIG = replace(
    STANDARD_CONFIG,
    mode="No Recourse",
    template_path=NO_RECOURSE_PDF_TEMPLATE_PATH,
    fields=_without_whiteout_map(_no_recourse_fields()),
    item_columns=_without_whiteout_map(
        {
            # No Recourse removes template placeholder text objects first, then
            # draws values only. These boxes must not paint over form borders.
            "qty_header": _box_for_baseline(x=35.0, baseline=398.3, width=60.0, height=10.0, font_size=7.8, min_font_size=6.2, align="center"),
            "qty": TextBox(37.0, 0, 50.0, 0, 8.8, min_font_size=6.5, align="center"),
            "type": TextBox(96.0, 0, 54.0, 0, 8.8, min_font_size=6.5, align="center"),
            "po": TextBox(159.0, 0, 72.0, 0, 8.0, min_font_size=6.5, align="center"),
            "description": TextBox(
                242.0,
                0,
                234.0,
                0,
                8.0,
                min_font_size=6.6,
                multiline=True,
                leading=10.4,
                vertical_align="middle",
            ),
            "skids": TextBox(490.0, 0, 44.0, 0, 8.8, min_font_size=6.5, align="center"),
            "weight": TextBox(542.0, 0, 48.0, 0, 8.8, min_font_size=6.5, align="center"),
        }
    ),
    item_row_height=22.0,
    max_item_rows=4,
    totals=_without_whiteout_map(
        {
            "qty": _box_for_baseline(x=37.0, baseline=207.2, width=50.0, height=12.0, font_size=8.6, min_font_size=6.5, bold=True, align="center"),
            "label": _box_for_baseline(x=_col_x(5) + 2, baseline=207.2, width=_col_width(5, 11) - 4, height=12.0, font_size=8.6, min_font_size=6.5, bold=True, align="center"),
            "skids": _box_for_baseline(x=490.0, baseline=207.2, width=44.0, height=12.0, font_size=8.6, min_font_size=6.5, bold=True, align="center"),
            "weight": _box_for_baseline(x=542.0, baseline=207.2, width=48.0, height=12.0, font_size=8.6, min_font_size=6.5, bold=True, align="center"),
        }
    ),
    item_row_baselines=(386.3, 364.6, 332.5, 300.5),
)

MULTISTOP_CONFIG = PdfTemplateConfig(
    mode="Multistop",
    template_path=MULTISTOP_PDF_TEMPLATE_PATH,
    fields=_without_whiteout_map({
        "bol_number": _top_value_box(709.3),
        "ship_date": _top_value_box(694.9),
        "carrier": _top_value_box(681.0, width=150.0),
        "load_number": _top_value_box(666.1),
        "kk_po_number": _top_value_box(653.2),
        "kk_load_number": _top_value_box(638.2),
        "comments": _box_for_baseline(x=444.4, baseline=606.7, width=128.0, height=18.0, font_size=7.8, min_font_size=5.2, multiline=True),
        "ship_from_company": TextBox(92, 604, 242, 13, 7.5),
        "ship_from_street": TextBox(92, 584, 242, 13, 7.5),
        "ship_from_city_state_zip": TextBox(92, 565, 242, 13, 7.5),
        "bill_to": _box_for_baseline(x=360, baseline=510.0, width=205, height=58.0, font_size=7.0, min_font_size=5.0, multiline=True),
        "delivery_1_dc": _box_for_baseline(x=103.5, baseline=556.5, width=145.0, font_size=7.4),
        "delivery_1_address": _box_for_baseline(x=103.5, baseline=538.7, width=220.0, height=22.0, font_size=7.0, min_font_size=5.0, multiline=True),
        "delivery_2_dc": _box_for_baseline(x=103.5, baseline=521.4, width=145.0, font_size=7.4),
        "delivery_2_address": _box_for_baseline(x=103.5, baseline=504.9, width=220.0, height=22.0, font_size=7.0, min_font_size=5.0, multiline=True),
        "delivery_3_dc": _box_for_baseline(x=103.5, baseline=487.0, width=145.0, font_size=7.4),
        "delivery_3_address": _box_for_baseline(x=103.5, baseline=467.8, width=220.0, height=22.0, font_size=7.0, min_font_size=5.0, multiline=True),
    }),
    item_columns=_without_whiteout_map({
        "dc": TextBox(34, 0, 42, 0, 7.0, align="center"),
        "case": TextBox(78, 0, 46, 0, 7.0, align="center"),
        "po": TextBox(126, 0, 83, 0, 6.6, align="center"),
        "description": TextBox(212, 0, 194, 0, 6.3, multiline=True),
        "pallet": TextBox(408, 0, 66, 0, 7.0, align="center"),
        "weight": TextBox(476, 0, 85, 0, 7.0, align="center"),
    }),
    item_start_y=0,
    item_row_height=17.0,
    max_item_rows=3,
    totals=_without_whiteout_map({
        "case": _box_for_baseline(x=78, baseline=99.9, width=46, height=12.0, font_size=7.4, bold=True, align="center"),
        "label": _box_for_baseline(x=212, baseline=99.9, width=194, height=12.0, font_size=7.4, bold=True, align="center"),
        "pallet": _box_for_baseline(x=408, baseline=99.9, width=66, height=12.0, font_size=7.4, bold=True, align="center"),
        "weight": _box_for_baseline(x=476, baseline=99.9, width=85, height=12.0, font_size=7.4, bold=True, align="center"),
    }),
    item_row_baselines=(277.7, 253.3, 229.1),
)


def _font_name(bold: bool) -> str:
    return FONT_BOLD if bold else FONT_NAME


def whiteout_box(canv: canvas.Canvas, box: TextBox) -> None:
    if not box.whiteout:
        return
    canv.saveState()
    canv.setFillColor(colors.white)
    canv.setStrokeColor(colors.white)
    canv.rect(box.x, box.y, box.width, box.height, stroke=0, fill=1)
    canv.restoreState()


def _fit_font_size(
    canv: canvas.Canvas,
    text: str,
    font_name: str,
    font_size: float,
    max_width: float,
    min_font_size: float,
) -> float:
    size = font_size
    while size > min_font_size and canv.stringWidth(text, font_name, size) > max_width:
        size -= 0.25
    return max(size, min_font_size)


def draw_fitted_text(canv: canvas.Canvas, box: TextBox, value: Any) -> None:
    text = _safe_text(value)
    whiteout_box(canv, box)
    if not text:
        return
    font_name = _font_name(box.bold)
    font_size = _fit_font_size(canv, text, font_name, box.font_size, box.width, box.min_font_size)
    canv.setFont(font_name, font_size)
    baseline = box.y + max((box.height - font_size) / 2, 0) + 1.2
    if box.align == "right":
        canv.drawRightString(box.x + box.width, baseline, text)
    elif box.align == "center":
        canv.drawCentredString(box.x + box.width / 2, baseline, text)
    else:
        canv.drawString(box.x, baseline, text)


def _split_line_to_width(
    canv: canvas.Canvas,
    text: str,
    font_name: str,
    font_size: float,
    max_width: float,
) -> list[str]:
    words = text.split()
    if not words:
        return [""]
    lines: list[str] = []
    current = words[0]
    for word in words[1:]:
        candidate = f"{current} {word}"
        if canv.stringWidth(candidate, font_name, font_size) <= max_width:
            current = candidate
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


def draw_multiline_text(canv: canvas.Canvas, box: TextBox, value: Any) -> None:
    text = _safe_text(value)
    whiteout_box(canv, box)
    if not text:
        return

    font_name = _font_name(box.bold)
    font_size = box.font_size
    while font_size > box.min_font_size:
        lines: list[str] = []
        for raw_line in text.splitlines() or [text]:
            lines.extend(_split_line_to_width(canv, raw_line, font_name, font_size, box.width))
        leading = box.leading or font_size + 1.2
        if len(lines) * leading <= box.height:
            break
        font_size -= 0.25

    final_font_size = max(font_size, box.min_font_size)
    canv.setFont(font_name, final_font_size)
    leading = box.leading or final_font_size + 1.2
    text_height = len(lines) * leading
    if box.vertical_align == "middle":
        y = box.y + box.height - max((box.height - text_height) / 2, 0) - final_font_size
    else:
        y = box.y + box.height - final_font_size
    for line in lines:
        if y < box.y:
            break
        if box.align == "center":
            canv.drawCentredString(box.x + box.width / 2, y, line)
        elif box.align == "right":
            canv.drawRightString(box.x + box.width, y, line)
        else:
            canv.drawString(box.x, y, line)
        y -= leading


def _draw_box_value(canv: canvas.Canvas, box: TextBox, value: Any) -> None:
    if box.multiline:
        draw_multiline_text(canv, box, value)
    else:
        draw_fitted_text(canv, box, value)


def _box_at_row_baseline(base_box: TextBox, baseline: float, height: float) -> TextBox:
    if base_box.multiline:
        y = baseline - height + base_box.font_size - 1.0
    else:
        y = baseline - max((height - base_box.font_size) / 2, 0) - 1.2
    return replace(base_box, y=y, height=height)


def _standard_record_values(
    record: BolStandardRecord,
    selected_facility: BolFacilityRecord,
    batch_comment: str | None,
    *,
    render_pickup_number: bool = True,
) -> dict[str, str]:
    comment = _safe_text(record.comments) or _safe_text(batch_comment)
    bill_to_lines = "\n".join(
        part
        for part in (
            record.bill_to.company,
            record.bill_to.street,
            record.bill_to.city_state_zip,
            "Attn:",
        )
        if _safe_text(part)
    )
    pickup_number = (
        _safe_text(getattr(record, "pickup_number", ""))
        if render_pickup_number
        else ""
    )
    return {
        "bol_number": record.bol_number,
        "ship_date": _format_ship_date_for_template(record.ship_date),
        "carrier": record.carrier,
        "carrier_pro_number": record.carrier_pro_number,
        "po_number": record.po_number,
        "kk_po_number": record.kk_po_number,
        "kk_load_number": record.kk_load_number,
        "seal_number": record.seal_number_blank,
        "pickup_number": pickup_number,
        "delivery_appt": pickup_number,
        "appt_number": pickup_number,
        "appointment_number": pickup_number,
        "tracker_number": "",
        "comments": comment,
        "ship_from_company": record.ship_from.company,
        "ship_from_street": record.ship_from.street,
        "ship_from_city_state_zip": record.ship_from.city_state_zip,
        "consignee_company": record.consignee_company,
        "consignee_street": record.consignee_street,
        "consignee_city_state_zip": record.consignee_city_state_zip,
        "dc_number": record.dc_number,
        "bill_to": bill_to_lines,
    }


def _no_recourse_record_values(
    record: BolStandardRecord,
    selected_facility: BolFacilityRecord,
    batch_comment: str | None,
    *,
    render_pickup_number: bool = True,
) -> dict[str, str]:
    values = _standard_record_values(
        record,
        selected_facility,
        batch_comment,
        render_pickup_number=render_pickup_number,
    )
    no_recourse_comment = _safe_text(record.comments)
    no_recourse_bill_to = "\n".join(
        part
        for part in (
            record.bill_to.company,
            record.bill_to.street,
            record.bill_to.city_state_zip,
        )
        if _safe_text(part)
    )
    values.update(
        {
            "comments": no_recourse_comment,
            "delivery_appt": "",
            "appt_number": "",
            "appointment_number": "",
            "tracker_number": "",
            "bill_to": no_recourse_bill_to,
        }
    )
    return values


def _standard_pdf_record_values(
    record: BolStandardRecord,
    selected_facility: BolFacilityRecord,
    batch_comment: str | None,
    *,
    render_pickup_number: bool = True,
) -> dict[str, str]:
    values = _standard_record_values(
        record,
        selected_facility,
        batch_comment,
        render_pickup_number=render_pickup_number,
    )
    ship_from_street = _safe_text(record.ship_from.street)
    ship_from_location = _safe_text(record.ship_from.city_state_zip)
    location_suffixes = {
        ship_from_location,
        ship_from_location.replace(",", ", "),
    }
    for suffix in sorted(location_suffixes, key=len, reverse=True):
        if suffix and ship_from_street.lower().endswith(suffix.lower()):
            ship_from_street = ship_from_street[: -len(suffix)].rstrip(" ,")
            break
    bill_to_lines = "\n".join(
        part
        for part in (
            record.bill_to.company,
            record.bill_to.street,
            record.bill_to.city_state_zip,
        )
        if _safe_text(part)
    )
    values.update(
        {
            "delivery_appt": (
                _safe_text(getattr(record, "pickup_number", ""))
                if render_pickup_number
                else ""
            ),
            "appt_number": "",
            "appointment_number": "",
            "ship_from_street": ship_from_street,
            "bill_to": bill_to_lines,
        }
    )
    return values


def _line_has_data(line: BolStandardItemLine) -> bool:
    return any(
        _safe_text(value)
        for value in (
            line.pallet_qty,
            line.po_number,
            line.item_description,
            line.item_number,
            line.upc,
            line.skids,
            line.weight_each,
            getattr(line, "total_weight", ""),
        )
    )


def _standard_item_lines(record: BolStandardRecord, *, mode: str) -> list[BolStandardItemLine]:
    if mode == "No Recourse":
        return [line for line in record.item_lines if _line_has_data(line)]
    return record.item_lines


def _display_weight(line: BolStandardItemLine) -> str:
    return _safe_text(getattr(line, "weight_each", ""))


def _standard_totals(record: BolStandardRecord, item_lines: list[BolStandardItemLine], *, mode: str) -> tuple[str, str, str]:
    total_pallet_qty_value = 0.0
    total_skids_value = 0.0
    total_weight_value = 0.0
    has_pallet_qty_value = False
    use_line_total_weight = any(_safe_text(getattr(line, "total_weight", "")) for line in item_lines)
    has_total_weight_value = False

    for line in item_lines:
        numeric_pallet_qty = _parse_numeric(line.pallet_qty)
        if numeric_pallet_qty is not None:
            total_pallet_qty_value += numeric_pallet_qty
            has_pallet_qty_value = True

        numeric_skids = _parse_numeric(line.skids)
        if numeric_skids is not None:
            total_skids_value += numeric_skids

        weight_source = getattr(line, "total_weight", "") if use_line_total_weight else line.weight_each
        numeric_weight = _parse_numeric(weight_source)
        if numeric_weight is not None:
            total_weight_value += numeric_weight
            if use_line_total_weight:
                has_total_weight_value = True

    if use_line_total_weight and not has_total_weight_value:
        for line in item_lines:
            numeric_weight = _parse_numeric(line.weight_each)
            if numeric_weight is not None:
                total_weight_value += numeric_weight

    total_qty = (
        _format_number(total_pallet_qty_value)
        if mode == "No Recourse" and has_pallet_qty_value
        else _format_number(record.total_skids)
    )
    return total_qty, _format_number(total_skids_value), _format_number(total_weight_value)


def _description_value(line: BolStandardItemLine) -> str:
    detail_parts: list[str] = []
    if _safe_text(line.item_number):
        detail_parts.append(f"Item #: {_safe_text(line.item_number)}")
    if _safe_text(line.upc):
        detail_parts.append(f"UPC #: {_safe_text(line.upc)}")
    detail_line = "        ".join(detail_parts)
    if detail_line and _safe_text(line.item_description):
        return f"{_safe_text(line.item_description)}\n{detail_line}"
    return _safe_text(line.item_description) or detail_line


def _no_recourse_first_row_box(column_name: str, box: TextBox) -> TextBox:
    if column_name in {"qty", "skids", "weight"}:
        return replace(box, font_size=9.2, min_font_size=6.8)
    if column_name == "po":
        return replace(box, font_size=8.5, min_font_size=6.8)
    if column_name == "description":
        return replace(box, font_size=8.6, min_font_size=7.2, leading=10.8)
    return box


def _draw_two_line_item_description(
    canv: canvas.Canvas,
    box: TextBox,
    line: BolStandardItemLine,
    *,
    description_size: float,
    detail_size: float,
    min_description_size: float,
    min_detail_size: float,
    leading: float,
    second_line_y_offset: float = 0.0,
) -> None:
    description = _safe_text(line.item_description)
    detail_parts: list[str] = []
    if _safe_text(line.item_number):
        detail_parts.append(f"Item #: {_safe_text(line.item_number)}")
    if _safe_text(line.upc):
        detail_parts.append(f"UPC #: {_safe_text(line.upc)}")
    detail_line = "        ".join(detail_parts)

    lines = [value for value in (description, detail_line) if value]
    if not lines:
        return

    line_specs = (
        (lines[0], description_size, min_description_size),
        (lines[1], detail_size, min_detail_size),
    ) if len(lines) == 2 else ((lines[0], description_size, min_description_size),)
    total_height = len(line_specs) * leading
    baseline = box.y + box.height - max((box.height - total_height) / 2, 0) - line_specs[0][1]
    for line_index, (text, font_size, min_font_size) in enumerate(line_specs):
        fitted_size = _fit_font_size(canv, text, FONT_NAME, font_size, box.width, min_font_size)
        canv.setFont(FONT_NAME, fitted_size)
        draw_y = baseline + (second_line_y_offset if line_index == 1 else 0.0)
        canv.drawString(box.x, draw_y, text)
        baseline -= leading


def _draw_no_recourse_first_row_description(
    canv: canvas.Canvas,
    box: TextBox,
    line: BolStandardItemLine,
) -> None:
    _draw_two_line_item_description(
        canv,
        box,
        line,
        description_size=9.0,
        detail_size=8.0,
        min_description_size=7.2,
        min_detail_size=6.8,
        leading=10.8,
        second_line_y_offset=3.5,
    )


def _draw_standard_overlay(
    canv: canvas.Canvas,
    config: PdfTemplateConfig,
    record: BolStandardRecord,
    selected_facility: BolFacilityRecord,
    *,
    bol_type: str | None,
    qty_type: str,
    batch_comment: str | None,
    render_pickup_number: bool = True,
) -> None:
    record_values = (
        _no_recourse_record_values(
            record,
            selected_facility,
            batch_comment,
            render_pickup_number=render_pickup_number,
        )
        if config.mode == "No Recourse"
        else _standard_pdf_record_values(
            record,
            selected_facility,
            batch_comment,
            render_pickup_number=render_pickup_number,
        )
    )
    for field_name, box in config.fields.items():
        _draw_box_value(
            canv,
            box,
            record_values.get(field_name, ""),
        )

    qty_header_box = config.item_columns.get("qty_header")
    if qty_header_box is not None:
        _draw_box_value(canv, qty_header_box, _qty_type_header(qty_type))

    rendered_type = _normalize_bol_type(bol_type)
    item_lines = _standard_item_lines(record, mode=config.mode)
    row_count = config.max_item_rows if config.mode == "No Recourse" else len(item_lines[: config.max_item_rows])
    for row_offset in range(row_count):
        row_baseline = (
            config.item_row_baselines[row_offset]
            if row_offset < len(config.item_row_baselines)
            else config.item_start_y - row_offset * config.item_row_height
        )
        row_height = config.item_row_height - 2
        line = item_lines[row_offset] if row_offset < len(item_lines) else None
        values = (
            {
                "qty": line.pallet_qty,
                "type": rendered_type,
                "po": line.po_number,
                "description": _description_value(line),
                "skids": line.skids,
                "weight": _display_weight(line),
            }
            if line is not None
            else {
                "qty": "",
                "type": "",
                "po": "",
                "description": "",
                "skids": "",
                "weight": "",
            }
        )
        for column_name, value in values.items():
            base_box = config.item_columns[column_name]
            if config.mode == "No Recourse" and row_offset == 0:
                base_box = _no_recourse_first_row_box(column_name, base_box)
            box = _box_at_row_baseline(base_box, row_baseline, row_height)
            if config.mode == "No Recourse" and row_offset == 0 and column_name == "description" and line is not None:
                _draw_no_recourse_first_row_description(canv, box, line)
                continue
            if config.mode == "Standard" and column_name == "description" and line is not None:
                _draw_two_line_item_description(
                    canv,
                    box,
                    line,
                    description_size=9.0,
                    detail_size=8.0,
                    min_description_size=7.0,
                    min_detail_size=6.6,
                    leading=10.6,
                    second_line_y_offset=10.0 if row_offset == 0 else 0.0,
                )
                continue
            _draw_box_value(canv, box, value)

    total_qty, total_skids, total_weight = _standard_totals(record, item_lines, mode=config.mode)
    for total_name, value in {
        "qty": total_qty,
        "label": "TOTALS",
        "skids": total_skids,
        "weight": total_weight,
    }.items():
        _draw_box_value(canv, config.totals[total_name], value)


def _multistop_record_values(
    record: BolMultistopRecord,
    selected_facility: BolFacilityRecord,
    batch_comment: str | None,
) -> dict[str, str]:
    comment = _safe_text(record.comments) or _safe_text(batch_comment)
    bill_to_lines = "\n".join(
        part
        for part in (
            record.bill_to.company,
            record.bill_to.street,
            record.bill_to.city_state_zip,
            "Attn:",
        )
        if _safe_text(part)
    )
    return {
        "bol_number": record.bol_number,
        "ship_date": _format_ship_date_for_template(record.ship_date),
        "carrier": record.carrier,
        "load_number": record.load_number,
        "kk_po_number": record.kk_po_number,
        "kk_load_number": record.kk_load_number,
        "comments": comment,
        "ship_from_company": selected_facility["facility_name"],
        "ship_from_street": selected_facility["address"],
        "ship_from_city_state_zip": selected_facility["location"],
        "bill_to": bill_to_lines,
        "delivery_1_dc": record.delivery_1_dc,
        "delivery_1_address": record.delivery_1_address,
        "delivery_2_dc": record.delivery_2_dc,
        "delivery_2_address": record.delivery_2_address,
        "delivery_3_dc": record.delivery_3_dc,
        "delivery_3_address": record.delivery_3_address,
    }


def _draw_multistop_overlay(
    canv: canvas.Canvas,
    config: PdfTemplateConfig,
    record: BolMultistopRecord,
    selected_facility: BolFacilityRecord,
    *,
    batch_comment: str | None,
) -> None:
    values = _multistop_record_values(record, selected_facility, batch_comment)
    for field_name, box in config.fields.items():
        _draw_box_value(canv, box, values.get(field_name, ""))

    for row_offset, stop in enumerate(record.stops[: config.max_item_rows]):
        row_baseline = (
            config.item_row_baselines[row_offset]
            if row_offset < len(config.item_row_baselines)
            else config.item_start_y - row_offset * config.item_row_height
        )
        row_values = {
            "dc": stop.dc_number,
            "case": stop.cases,
            "po": stop.target_po_number,
            "description": _format_multistop_item_description(
                stop.pallet_description,
                stop.item_number,
                stop.upc,
            ),
            "pallet": stop.total_pallets,
            "weight": stop.weight,
        }
        for column_name, value in row_values.items():
            base_box = config.item_columns[column_name]
            box = _box_at_row_baseline(base_box, row_baseline, config.item_row_height - 3)
            _draw_box_value(canv, box, value)

    for total_name, value in {
        "case": _format_number(record.total_case),
        "label": "TOTALS",
        "pallet": _format_number(record.total_pallet),
        "weight": _format_number(record.total_ship_weight),
    }.items():
        _draw_box_value(canv, config.totals[total_name], value)


def _create_overlay_pdf(
    template_page,
    draw_callback: Callable[[canvas.Canvas], None],
) -> BytesIO:
    width = float(template_page.mediabox.width)
    height = float(template_page.mediabox.height)
    overlay_buffer = BytesIO()
    canv = canvas.Canvas(overlay_buffer, pagesize=(width, height))
    draw_callback(canv)
    canv.save()
    overlay_buffer.seek(0)
    return overlay_buffer


def _stamp_template_pdf(
    *,
    template_path: Path,
    destination_pdf: Path,
    draw_callback: Callable[[canvas.Canvas], None],
    strip_known_tokens: bool = False,
) -> None:
    destination_pdf.parent.mkdir(parents=True, exist_ok=True)
    template_reader = PdfReader(str(template_path))
    if not template_reader.pages:
        raise RuntimeError(f"PDF template has no pages: {template_path}")

    writer = PdfWriter()
    writer.add_page(template_reader.pages[0])

    template_page = writer.pages[0]
    _strip_template_placeholder_text(template_page, strip_known_tokens=strip_known_tokens)
    overlay_buffer = _create_overlay_pdf(template_page, draw_callback)
    overlay_reader = PdfReader(overlay_buffer)
    template_page.merge_page(overlay_reader.pages[0])

    with destination_pdf.open("wb") as output_file:
        writer.write(output_file)


def _strip_template_placeholder_text(page: Any, *, strip_known_tokens: bool = False) -> None:
    content = page.get_contents()
    if content is None:
        return

    content_stream = ContentStream(content, page.pdf)
    inside_placeholder = False
    for operands, operator in content_stream.operations:
        if operator in (b"Tj", b"'", b'"'):
            if operands:
                operands[0], inside_placeholder = _strip_placeholder_fragment(
                    operands[0],
                    inside_placeholder,
                    strip_known_tokens=strip_known_tokens,
                )
        elif operator == b"TJ" and operands:
            text_array = operands[0]
            if isinstance(text_array, ArrayObject):
                joined_text = "".join(value for value in text_array if isinstance(value, str))
                if _should_strip_text_fragment(joined_text, inside_placeholder, strip_known_tokens=strip_known_tokens):
                    for index, value in enumerate(text_array):
                        if isinstance(value, str):
                            text_array[index] = TextStringObject("")
                    inside_placeholder = "\u00ab" in joined_text and "\u00bb" not in joined_text
                    continue
                for index, value in enumerate(text_array):
                    text_array[index], inside_placeholder = _strip_placeholder_fragment(
                        value,
                        inside_placeholder,
                        strip_known_tokens=strip_known_tokens,
                    )

    page[NameObject("/Contents")] = content_stream


def _strip_placeholder_fragment(
    value: Any,
    inside_placeholder: bool,
    *,
    strip_known_tokens: bool = False,
) -> tuple[Any, bool]:
    if not isinstance(value, str):
        return value, inside_placeholder

    if _should_strip_text_fragment(value, inside_placeholder, strip_known_tokens=strip_known_tokens):
        starts_placeholder = "\u00ab" in value or "Ť" in value
        ends_placeholder = "\u00bb" in value or "ť" in value
        return TextStringObject(""), starts_placeholder or (inside_placeholder and not ends_placeholder)

    return value, inside_placeholder


def _should_strip_text_fragment(
    value: str,
    inside_placeholder: bool,
    *,
    strip_known_tokens: bool,
) -> bool:
    starts_placeholder = "\u00ab" in value
    ends_placeholder = "\u00bb" in value
    starts_placeholder = starts_placeholder or "Ť" in value
    ends_placeholder = ends_placeholder or "ť" in value
    contains_known_token = strip_known_tokens and any(token in value for token in KNOWN_PLACEHOLDER_TOKENS)
    return (
        inside_placeholder
        or starts_placeholder
        or ends_placeholder
        or PLACEHOLDER_TEXT_PATTERN.search(value)
        or contains_known_token
    )


def _records_by_bol(records: list[Any]) -> dict[str, Any]:
    return {_safe_text(getattr(record, "bol_number", "")): record for record in records}


def _multistop_records_by_key(records: list[BolMultistopRecord]) -> dict[tuple[str, str], BolMultistopRecord]:
    return {
        (_safe_text(record.bol_number), _safe_text(record.load_number)): record
        for record in records
    }


def _config_for_mode(mode: str) -> PdfTemplateConfig:
    if mode == "Standard":
        return STANDARD_CONFIG
    if mode == "No Recourse":
        return NO_RECOURSE_CONFIG
    if mode == "Multistop":
        return MULTISTOP_CONFIG
    raise ValueError(f"Unsupported BOL PDF template mode: {mode}.")


def _template_unavailable_result(template_path: Path, output_root: Path) -> StandardPdfConversionResult:
    return StandardPdfConversionResult(
        output_dir=str(output_root.resolve()),
        converted_files=[],
        failed_conversions=[],
        converter_name="pdf-template-stamper",
        conversion_available=False,
        unavailable_reason=f"PDF template file not found: {template_path}",
        converter_path=str(template_path),
    )


def stamp_bol_pdf_set(
    records: list[Any],
    selected_facility: BolFacilityRecord | None,
    generated_docx_files: list[GeneratedDocxFile],
    *,
    mode: str,
    bol_type: str | None = None,
    qty_type: str = "PLT",
    batch_comment: str | None = None,
    render_pickup_number: bool = True,
    output_dir: Path | None = None,
    progress_callback: Callable[[int, int, GeneratedDocxFile], None] | None = None,
) -> StandardPdfConversionResult:
    """Stamp BOL PDFs by merging dynamic overlays onto static PDF templates."""

    if selected_facility is None:
        raise ValueError("No ship-from facility is selected. Select a facility before PDF generation.")
    selected_ship_from = facility_to_ship_from(selected_facility)
    for record in records:
        if hasattr(record, "ship_from"):
            record.ship_from = selected_ship_from
    if not generated_docx_files:
        raise ValueError("No generated DOCX files were provided for PDF naming.")

    config = _config_for_mode(mode)
    output_root = output_dir or Path(mkdtemp(prefix="kkg_bol_template_pdf_"))
    output_root.mkdir(parents=True, exist_ok=True)
    if not config.template_path.exists():
        return _template_unavailable_result(config.template_path, output_root)

    converted_files: list[ConvertedPdfFile] = []
    failed_conversions: list[FailedPdfConversion] = []
    total_files = len(generated_docx_files)

    standard_lookup = _records_by_bol(records)
    multistop_lookup = _multistop_records_by_key(records) if mode == "Multistop" else {}

    for index, generated_file in enumerate(generated_docx_files, start=1):
        if progress_callback is not None:
            progress_callback(index, total_files, generated_file)

        if mode == "Multistop" and getattr(generated_file, "document_type", "") == "stop":
            continue

        source_docx = Path(generated_file.file_path)
        destination_pdf = output_root / f"{source_docx.stem}.pdf"

        try:
            if mode == "Multistop":
                record = multistop_lookup.get(
                    (
                        _safe_text(generated_file.bol_number),
                        _safe_text(getattr(generated_file, "load_number", "")),
                    )
                )
                if record is None:
                    raise RuntimeError("Matching multistop BOL record was not found for PDF stamping.")
                _stamp_template_pdf(
                    template_path=config.template_path,
                    destination_pdf=destination_pdf,
                    strip_known_tokens=True,
                    draw_callback=lambda canv, record=record: _draw_multistop_overlay(
                        canv,
                        config,
                        record,
                        selected_facility,
                        batch_comment=batch_comment,
                    ),
                )
            else:
                record = standard_lookup.get(_safe_text(generated_file.bol_number))
                if record is None:
                    raise RuntimeError("Matching BOL record was not found for PDF stamping.")
                _stamp_template_pdf(
                    template_path=config.template_path,
                    destination_pdf=destination_pdf,
                    strip_known_tokens=True,
                    draw_callback=lambda canv, record=record: _draw_standard_overlay(
                        canv,
                        config,
                        record,
                        selected_facility,
                        bol_type=bol_type,
                        qty_type=qty_type,
                        batch_comment=batch_comment,
                        render_pickup_number=render_pickup_number,
                    ),
                )

            converted_files.append(
                ConvertedPdfFile(
                    bol_number=generated_file.bol_number,
                    file_name=destination_pdf.name,
                    file_path=str(destination_pdf.resolve()),
                    document_type=str(getattr(generated_file, "document_type", "") or ""),
                    load_number=str(getattr(generated_file, "load_number", "") or ""),
                    stop_number=getattr(generated_file, "stop_number", None),
                )
            )
        except Exception as exc:
            failed_conversions.append(
                FailedPdfConversion(
                    bol_number=generated_file.bol_number,
                    source_docx=str(source_docx),
                    error=str(exc),
                )
            )

    return StandardPdfConversionResult(
        output_dir=str(output_root.resolve()),
        converted_files=converted_files,
        failed_conversions=failed_conversions,
        converter_name="pdf-template-stamper",
        conversion_available=True,
        unavailable_reason=None,
        converter_path=str(config.template_path.resolve()),
    )
