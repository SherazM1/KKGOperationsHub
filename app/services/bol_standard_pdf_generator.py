"""Direct PDF generation for Standard-family BOL records.

The renderer is intentionally coordinate-based.  The coordinates mirror the
current Standard / No Recourse DOCX templates, whose main layout is a single
wide Word table with narrow grid columns and stacked left-side shipment blocks.
"""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
from tempfile import mkdtemp
from typing import Any, Callable
from zipfile import ZipFile

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph

from app.models.bol_standard_record import BolStandardItemLine, BolStandardRecord
from app.services.bol_standard_docx_generator import (
    GeneratedDocxFile,
    NO_RECOURSE_TEMPLATE_PATH,
    STANDARD_TEMPLATE_PATH,
    _format_number,
    _format_ship_date_for_template,
    _parse_numeric,
    _qty_type_header,
)
from app.services.bol_standard_pdf_converter import (
    ConvertedPdfFile,
    FailedPdfConversion,
    StandardPdfConversionResult,
)
from app.utils.bol_facilities import BolFacilityRecord


PAGE_WIDTH, PAGE_HEIGHT = letter
FONT_NAME = "Helvetica"
FONT_BOLD = "Helvetica-Bold"
LINE_COLOR = colors.black
GRAY_FILL = colors.HexColor("#D9D9D9")
YELLOW_FILL = colors.HexColor("#FFFF00")

# The DOCX table grid is 11,694 twips wide: 584.7 pt.  Centering it on letter
# paper lands very close to the Word template's narrow left/right margins.
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

# Row heights copied from the Standard template.  The No Recourse template uses
# the same visual grid even though some explicit heights are omitted in XML.
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


def _safe_text(value: Any) -> str:
    return str(value or "").strip()


def _normalize_bol_type(bol_type: str | None) -> str:
    normalized = (bol_type or "PLT").strip().upper()
    return normalized if normalized in {"PLT", "CASE"} else "PLT"


def _resolve_comment(record: BolStandardRecord, batch_comment: str | None) -> str:
    record_comment = _safe_text(record.comments)
    return record_comment if record_comment else _safe_text(batch_comment)


def _col_x(col_index: int) -> float:
    return TABLE_X + sum(COL_WIDTHS[:col_index])


def _col_width(start_col: int, end_col: int) -> float:
    return sum(COL_WIDTHS[start_col:end_col])


def _row_top(row_index: int) -> float:
    return TABLE_TOP - sum(ROW_HEIGHTS_PT[:row_index])


def _row_bottom(row_index: int) -> float:
    return _row_top(row_index) - ROW_HEIGHTS_PT[row_index]


def _row_span_top(row_start: int) -> float:
    return _row_top(row_start)


def _row_span_bottom(row_start: int, row_end: int) -> float:
    return TABLE_TOP - sum(ROW_HEIGHTS_PT[:row_end])


def _box(
    canv: canvas.Canvas,
    col_start: int,
    col_end: int,
    row_start: int,
    row_end: int,
    *,
    fill: colors.Color | None = None,
    stroke: bool = True,
    line_width: float = 0.55,
) -> tuple[float, float, float, float]:
    x = _col_x(col_start)
    y = _row_span_bottom(row_start, row_end)
    width = _col_width(col_start, col_end)
    height = _row_span_top(row_start) - y
    canv.setLineWidth(line_width)
    canv.setStrokeColor(LINE_COLOR)
    if fill is not None:
        canv.setFillColor(fill)
        canv.rect(x, y, width, height, stroke=1 if stroke else 0, fill=1)
        canv.setFillColor(colors.black)
    elif stroke:
        canv.rect(x, y, width, height, stroke=1, fill=0)
    return x, y, width, height


def _fit_font_size(
    canv: canvas.Canvas,
    text: str,
    font_name: str,
    base_size: float,
    max_width: float,
    min_size: float = 5.0,
) -> float:
    text = _safe_text(text)
    size = base_size
    while size > min_size and canv.stringWidth(text, font_name, size) > max_width:
        size -= 0.25
    return size


def _draw_text(
    canv: canvas.Canvas,
    x: float,
    y: float,
    text: str,
    width: float,
    *,
    font_name: str = FONT_NAME,
    font_size: float = 7.0,
    min_size: float = 5.0,
    align: str = "left",
) -> None:
    text = _safe_text(text)
    size = _fit_font_size(canv, text, font_name, font_size, width, min_size)
    canv.setFont(font_name, size)
    if align == "right":
        canv.drawRightString(x + width, y, text)
    elif align == "center":
        canv.drawCentredString(x + width / 2, y, text)
    else:
        canv.drawString(x, y, text)


def _style(
    name: str,
    *,
    font_name: str = FONT_NAME,
    font_size: float = 7.0,
    leading: float | None = None,
) -> ParagraphStyle:
    return ParagraphStyle(
        name=name,
        fontName=font_name,
        fontSize=font_size,
        leading=leading or font_size + 1.1,
        alignment=TA_LEFT,
        spaceAfter=0,
        spaceBefore=0,
    )


def _draw_paragraph(
    canv: canvas.Canvas,
    text: str,
    x: float,
    y_top: float,
    width: float,
    height: float,
    *,
    style: ParagraphStyle,
) -> None:
    paragraph = Paragraph(_safe_text(text).replace("\n", "<br/>"), style)
    paragraph.wrapOn(canv, width, height)
    paragraph.drawOn(canv, x, y_top - paragraph.height)


def _template_path_for_mode(mode: str) -> Path:
    return NO_RECOURSE_TEMPLATE_PATH if mode == "No Recourse" else STANDARD_TEMPLATE_PATH


def _draw_template_logo(canv: canvas.Canvas, mode: str) -> None:
    template_path = _template_path_for_mode(mode)
    try:
        with ZipFile(template_path, "r") as archive:
            image_bytes = archive.read("word/media/image1.png")
    except Exception:
        image_bytes = b""

    logo_x = _col_x(0) + 8
    logo_top = _row_top(0) - 2
    logo_w = _col_width(0, 7) - 14
    logo_h = 36
    if image_bytes:
        canv.drawImage(
            ImageReader(BytesIO(image_bytes)),
            logo_x,
            logo_top - logo_h,
            width=logo_w,
            height=logo_h,
            preserveAspectRatio=True,
            anchor="c",
            mask="auto",
        )
    else:
        _draw_text(
            canv,
            logo_x,
            logo_top - 20,
            "Kendal King",
            logo_w,
            font_name=FONT_BOLD,
            font_size=18,
            align="center",
        )

    _draw_text(
        canv,
        _col_x(0),
        _row_bottom(1) + 2,
        "609 SW 8th St - Ste 140 - Bentonville, AR 72712",
        _col_width(0, 7),
        font_size=7.0,
        align="center",
    )
    _draw_text(
        canv,
        _col_x(8),
        _row_bottom(0) + 3,
        "UNIFORM BILL OF LADING",
        _col_width(8, 18),
        font_name=FONT_BOLD,
        font_size=12.0,
        align="center",
    )


def _draw_right_field(
    canv: canvas.Canvas,
    row: int,
    label: str,
    value: str,
    *,
    label_cols: tuple[int, int] = (8, 9),
    value_cols: tuple[int, int] = (9, 14),
) -> None:
    y = _row_bottom(row) + 3.2
    _draw_text(
        canv,
        _col_x(label_cols[0]) + 2,
        y,
        label,
        _col_width(*label_cols) - 4,
        font_name=FONT_BOLD,
        font_size=7.0,
        align="right",
    )
    _draw_text(
        canv,
        _col_x(value_cols[0]) + 4,
        y,
        value,
        _col_width(*value_cols) - 6,
        font_size=7.2,
        min_size=5.3,
    )


def _draw_header_and_fields(
    canv: canvas.Canvas,
    record: BolStandardRecord,
    mode: str,
    *,
    resolved_comment: str,
) -> None:
    _draw_template_logo(canv, mode)
    _draw_right_field(canv, 1, "BOL #", record.bol_number)
    _draw_right_field(canv, 2, "Ship Date", _format_ship_date_for_template(record.ship_date))
    _draw_right_field(canv, 3, "Carrier", record.carrier, value_cols=(9, 15))
    _draw_right_field(canv, 4, "Carrier Pro #", record.carrier_pro_number)
    _draw_right_field(canv, 5, "PO #", record.po_number)

    if mode == "No Recourse":
        _draw_right_field(canv, 6, "KK PO #", record.kk_po_number)
        _draw_right_field(canv, 7, "KK Load #", record.kk_load_number)
        _draw_right_field(canv, 8, "Seal #", record.seal_number_blank)
        _draw_right_field(canv, 9, "Pick Up #", getattr(record, "pickup_number", ""), value_cols=(9, 15))
        if resolved_comment:
            _draw_right_field(canv, 11, "Comments", resolved_comment, value_cols=(9, 18))
    else:
        _draw_right_field(canv, 6, "Tracker #", "")
        _draw_right_field(canv, 7, "KK PO #", record.kk_po_number)
        _draw_right_field(canv, 8, "KKG Load #", record.kk_load_number, value_cols=(9, 15))
        _draw_right_field(canv, 9, "Delivery Appt.", getattr(record, "pickup_number", ""), value_cols=(9, 15))
        _draw_right_field(canv, 10, "APPT #", getattr(record, "pickup_number", ""), value_cols=(9, 15))
        _draw_right_field(canv, 12, "Seal #", record.seal_number_blank)
        if resolved_comment:
            _draw_right_field(canv, 11, "Comments", resolved_comment, value_cols=(9, 18))


def _draw_section_header(canv: canvas.Canvas, col_start: int, col_end: int, row: int, title: str) -> None:
    x, y, width, height = _box(canv, col_start, col_end, row, row + 1, fill=GRAY_FILL)
    _draw_text(
        canv,
        x + 4,
        y + height - 9,
        title,
        width - 8,
        font_name=FONT_BOLD,
        font_size=7.3,
    )


def _draw_left_label_value(
    canv: canvas.Canvas,
    row: int,
    label: str,
    value: str,
    *,
    value_row: int | None = None,
) -> None:
    label_y = _row_bottom(row) + 3.2
    value_y = _row_bottom(value_row if value_row is not None else row) + 3.2
    _draw_text(canv, _col_x(0) + 3, label_y, label, _col_width(0, 1) - 6, font_name=FONT_BOLD, font_size=6.6)
    _draw_text(canv, _col_x(1) + 3, value_y, value, _col_width(1, 7) - 6, font_size=7.0, min_size=5.2)


def _draw_stacked_shipper_consignee(
    canv: canvas.Canvas,
    record: BolStandardRecord,
    selected_facility: BolFacilityRecord,
    mode: str,
) -> None:
    _draw_section_header(canv, 0, 7, 6, "FROM (SHIPPER)")
    _box(canv, 0, 7, 7, 13)
    _draw_left_label_value(canv, 7, "COMPANY", selected_facility["facility_name"])
    _draw_left_label_value(canv, 9, "STREET", selected_facility["address"])
    _draw_text(
        canv,
        _col_x(1) + 3,
        _row_bottom(10) + 3.2,
        selected_facility["location"],
        _col_width(1, 7) - 6,
        font_size=7.0,
    )
    _draw_left_label_value(canv, 12, "ATTN", "")

    _draw_section_header(canv, 0, 7, 13, "TO (CONSIGNEE)")
    _box(canv, 0, 7, 14, 22)
    _draw_left_label_value(canv, 14, "COMPANY", record.consignee_company)
    _draw_left_label_value(canv, 16, "STREET", record.consignee_street)
    _draw_left_label_value(canv, 17, "CITY/ST/ZIP", record.consignee_city_state_zip)
    if mode == "Standard":
        _draw_left_label_value(canv, 18, "APPOINTMENT #", getattr(record, "pickup_number", ""))
        _draw_left_label_value(canv, 19, "DC:", record.dc_number)
    else:
        _draw_left_label_value(canv, 19, "DC#", record.dc_number)


def _draw_freight_billto_subject(canv: canvas.Canvas, record: BolStandardRecord) -> None:
    _draw_section_header(canv, 8, 18, 13, "FREIGHT CHARGE TERMS:")
    _box(canv, 8, 18, 14, 17)
    terms_x = _col_x(9) + 6
    _draw_text(canv, terms_x, _row_bottom(14) + 4, "X      FREIGHT PREPAID", 160, font_size=7.2)
    _draw_text(canv, terms_x, _row_bottom(15) + 4, "FREIGHT COLLECT", 160, font_size=7.2)
    _draw_text(canv, terms_x, _row_bottom(16) + 4, "FREIGHT THIRD PARTY", 160, font_size=7.2)

    _draw_section_header(canv, 8, 18, 17, "BILL TO:")
    _box(canv, 8, 18, 18, 22)
    bill_x = _col_x(8) + 12
    bill_w = _col_width(8, 18) - 24
    bill_lines = [
        record.bill_to.company,
        record.bill_to.street,
        record.bill_to.city_state_zip,
        "Attn:",
    ]
    y = _row_top(18) - 10
    for line in bill_lines:
        _draw_text(canv, bill_x, y, line, bill_w, font_size=7.2)
        y -= 9.0

    subject_x = _col_x(0)
    subject_top = _row_top(21) - 2
    subject_w = _col_width(0, 7)
    subject_h = _row_span_top(21) - _row_span_bottom(21, 23)
    canv.setLineWidth(0.55)
    canv.rect(subject_x, subject_top - subject_h, subject_w, subject_h, stroke=1, fill=0)
    subject = (
        "SUBJECT TO SECTION 7: Of the conditions if shipment is to be delivered to consignee "
        "without recourse on the consignor, the consignor shall sign the following statement: "
        "The Carrier shall not make delivery of the shipment without payment of the freight "
        "and all other lawful charges."
    )
    _draw_paragraph(
        canv,
        subject,
        subject_x + 5,
        subject_top - 5,
        subject_w - 10,
        subject_h - 10,
        style=_style("Subject7", font_size=5.9, leading=6.7),
    )


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


def _rendered_item_lines(
    item_lines: list[BolStandardItemLine],
    *,
    filter_blank_item_lines: bool,
) -> list[BolStandardItemLine]:
    if not filter_blank_item_lines:
        return item_lines
    return [line for line in item_lines if _line_has_data(line)]


def _calculate_totals(
    item_lines: list[BolStandardItemLine],
    total_qty: float,
    *,
    filter_blank_item_lines: bool,
) -> tuple[str, str, str]:
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

    total_qty_display = (
        _format_number(total_pallet_qty_value)
        if filter_blank_item_lines and has_pallet_qty_value
        else _format_number(total_qty)
    )
    return total_qty_display, _format_number(total_skids_value), _format_number(total_weight_value)


def _draw_table_cell_text(
    canv: canvas.Canvas,
    col_start: int,
    col_end: int,
    row_start: int,
    row_end: int,
    text: str,
    *,
    font_name: str = FONT_NAME,
    font_size: float = 6.5,
    align: str = "left",
) -> None:
    x = _col_x(col_start)
    y = _row_span_bottom(row_start, row_end)
    width = _col_width(col_start, col_end)
    height = _row_span_top(row_start) - y
    _draw_text(
        canv,
        x + 3,
        y + height - 9,
        text,
        width - 6,
        font_name=font_name,
        font_size=font_size,
        min_size=5.0,
        align=align,
    )


def _draw_item_description(canv: canvas.Canvas, row: int, line: BolStandardItemLine) -> None:
    x = _col_x(5) + 4
    y_top = _row_top(row) - 4
    width = _col_width(5, 11) - 8
    height = ROW_HEIGHTS_PT[row] - 6
    text = (
        f"{_safe_text(line.item_description)}<br/>"
        f"Item #: {_safe_text(line.item_number)}     UPC #: {_safe_text(line.upc)}"
    )
    _draw_paragraph(
        canv,
        text,
        x,
        y_top,
        width,
        height,
        style=_style("ItemDescription", font_size=6.2, leading=7.0),
    )


def _draw_item_table(
    canv: canvas.Canvas,
    record: BolStandardRecord,
    *,
    mode: str,
    bol_type: str | None,
    qty_type: str,
) -> None:
    rendered_type = _normalize_bol_type(bol_type)
    filter_blank = mode == "No Recourse"
    item_lines = _rendered_item_lines(record.item_lines, filter_blank_item_lines=filter_blank)
    total_qty, total_skids, total_weight = _calculate_totals(
        item_lines,
        record.total_skids,
        filter_blank_item_lines=filter_blank,
    )

    # Header and item table grid from the DOCX template.
    for row in range(23, 34):
        _box(canv, 0, 1, row, row + 1, fill=GRAY_FILL if row == 23 else None)
        _box(canv, 1, 3, row, row + 1, fill=GRAY_FILL if row == 23 else None)
        _box(canv, 3, 5, row, row + 1, fill=GRAY_FILL if row == 23 else None)
        _box(canv, 5, 11, row, row + 1, fill=GRAY_FILL if row == 23 else None)
        _box(canv, 11, 14, row, row + 1, fill=GRAY_FILL if row == 23 else None)
        _box(canv, 14, 19, row, row + 1, fill=GRAY_FILL if row == 23 else None)

    headers = [
        (0, 1, _qty_type_header(qty_type)),
        (1, 3, "Type"),
        (3, 5, "PO #"),
        (5, 11, "ITEM DESCRIPTION"),
        (11, 14, "# SKIDS"),
        (14, 19, "WEIGHT"),
    ]
    for col_start, col_end, label in headers:
        _draw_table_cell_text(
            canv,
            col_start,
            col_end,
            23,
            24,
            label,
            font_name=FONT_BOLD,
            font_size=6.5,
            align="center",
        )

    item_rows = list(range(24, 33))
    for row, line in zip(item_rows, item_lines):
        _draw_table_cell_text(canv, 0, 1, row, row + 1, line.pallet_qty, font_size=6.5)
        _draw_table_cell_text(canv, 1, 3, row, row + 1, rendered_type, font_size=6.0)
        _draw_table_cell_text(canv, 3, 5, row, row + 1, line.po_number, font_size=6.2)
        _draw_item_description(canv, row, line)
        _draw_table_cell_text(canv, 11, 14, row, row + 1, line.skids, font_size=6.5)
        _draw_table_cell_text(canv, 14, 19, row, row + 1, line.weight_each, font_size=6.5)

    _draw_table_cell_text(canv, 5, 11, 33, 34, "TOTALS", font_name=FONT_BOLD, font_size=7.0, align="center")
    _draw_table_cell_text(canv, 0, 1, 33, 34, total_qty, font_name=FONT_BOLD, font_size=7.0)
    _draw_table_cell_text(canv, 11, 14, 33, 34, total_skids, font_name=FONT_BOLD, font_size=7.0)
    _draw_table_cell_text(canv, 14, 19, 33, 34, total_weight, font_name=FONT_BOLD, font_size=7.0)


def _draw_standard_footer(canv: canvas.Canvas) -> None:
    y = _row_bottom(34) - 8
    _draw_text(canv, _col_x(0), y, "Shipper Signature:", _col_width(0, 6), font_name=FONT_BOLD, font_size=7)
    canv.line(_col_x(2), y - 2, _col_x(9), y - 2)
    _draw_text(canv, _col_x(10), y, "Driver Signature:", _col_width(10, 4), font_name=FONT_BOLD, font_size=7)
    canv.line(_col_x(13), y - 2, _col_x(19), y - 2)


def _draw_no_recourse_footer(canv: canvas.Canvas) -> None:
    row_top = _row_top(34)
    row_y = _row_bottom(34)
    _draw_table_cell_text(canv, 0, 6, 34, 35, "Shipper Signature:", font_name=FONT_BOLD, font_size=7.0)
    canv.line(_col_x(2), row_y + 5, _col_x(7), row_y + 5)

    note_top = row_top - 18
    note_h = 22
    canv.rect(_col_x(0), note_top - note_h, _col_width(0, 19), note_h, stroke=1, fill=0)
    note = (
        "CUSTOMER NOTE: Please check this shipment carefully. No returns accepted without prior approval. "
        "If shipment does not agree with this manifest, have carrier note discrepancy on receipt and notify "
        "shipper immediately."
    )
    _draw_paragraph(
        canv,
        note,
        _col_x(0) + 5,
        note_top - 4,
        _col_width(0, 19) - 10,
        note_h - 8,
        style=_style("CustomerNote", font_size=5.1, leading=5.7),
    )

    sig_top = note_top - note_h - 3
    sig_h = 16
    half_w = _col_width(0, 19) / 2
    canv.rect(_col_x(0), sig_top - sig_h, half_w, sig_h, stroke=1, fill=0)
    canv.setFillColor(YELLOW_FILL)
    canv.rect(_col_x(0) + half_w, sig_top - sig_h, half_w, sig_h, stroke=1, fill=1)
    canv.setFillColor(colors.black)
    _draw_text(canv, _col_x(0) + 5, sig_top - 11, "Receiver Signature:", half_w - 10, font_name=FONT_BOLD, font_size=6.6)
    _draw_text(
        canv,
        _col_x(0) + half_w + 5,
        sig_top - 13,
        "Driver Signature:",
        half_w - 10,
        font_name=FONT_BOLD,
        font_size=6.6,
    )

    notice_top = sig_top - sig_h - 3
    notice_h = 44
    canv.rect(_col_x(0), notice_top - notice_h, _col_width(0, 19), notice_h, stroke=1, fill=0)
    _draw_text(
        canv,
        _col_x(0) + 5,
        notice_top - 10,
        "BROKER PAYMENT & NO RECOURSE NOTICE",
        _col_width(0, 19) - 10,
        font_name=FONT_BOLD,
        font_size=6.8,
        align="center",
    )
    _draw_text(
        canv,
        _col_x(0) + 5,
        notice_top - 20,
        "Broker of Record: TRIDENT TRANSPORT, LLC",
        _col_width(0, 19) - 10,
        font_name=FONT_BOLD,
        font_size=5.8,
    )
    notice = (
        "Freight charges for this shipment are to be paid solely by the Broker of Record. Carrier agrees "
        "that it shall look exclusively to the Broker for payment of freight charges and hereby waives any "
        "right of recourse against Kendal King Group for unpaid freight charges. Payment by Kendal King Group "
        "to Broker shall constitute full satisfaction of Shipper's freight payment obligation. Carrier "
        "acknowledges and agrees that it has no lien, claim, or right to pursue Kendal King Group for unpaid "
        "freight charges. This provision shall survive delivery and any termination of transportation services."
    )
    _draw_paragraph(
        canv,
        notice,
        _col_x(0) + 5,
        notice_top - 25,
        _col_width(0, 19) - 10,
        notice_h - 27,
        style=_style("NoRecourse", font_size=4.25, leading=4.75),
    )

    legal_top = notice_top - notice_h - 3
    legal = (
        "The transportation of the property described herein is tendered and accepted subject to applicable "
        "federal transportation law and the written agreements between the Parties. In the event of any conflict "
        "between this Bill of Lading and any other document, the applicable written contract between Kendal King "
        "Group and the Broker of Record shall control. Nothing herein shall be construed to impose payment "
        "responsibility on Kendal King Group beyond its obligation to pay the Broker of Record in accordance with "
        "the governing agreement."
    )
    _draw_paragraph(
        canv,
        legal,
        _col_x(0) + 5,
        legal_top,
        _col_width(0, 19) - 10,
        28,
        style=_style("FinalLegal", font_size=4.15, leading=4.65),
    )


def _draw_bol_pdf(
    destination_pdf: Path,
    record: BolStandardRecord,
    selected_facility: BolFacilityRecord,
    *,
    mode: str,
    bol_type: str | None,
    qty_type: str,
    batch_comment: str | None,
) -> None:
    destination_pdf.parent.mkdir(parents=True, exist_ok=True)
    canv = canvas.Canvas(str(destination_pdf), pagesize=letter)
    canv.setTitle(f"{mode} BOL {record.bol_number}")

    resolved_comment = _resolve_comment(record, batch_comment)
    _draw_header_and_fields(canv, record, mode, resolved_comment=resolved_comment)
    _draw_stacked_shipper_consignee(canv, record, selected_facility, mode)
    _draw_freight_billto_subject(canv, record)
    _draw_item_table(canv, record, mode=mode, bol_type=bol_type, qty_type=qty_type)
    if mode == "No Recourse":
        _draw_no_recourse_footer(canv)
    else:
        _draw_standard_footer(canv)

    canv.showPage()
    canv.save()


def _records_by_bol(records: list[BolStandardRecord]) -> dict[str, BolStandardRecord]:
    return {_safe_text(record.bol_number): record for record in records}


def generate_standard_pdf_set(
    records: list[BolStandardRecord],
    selected_facility: BolFacilityRecord | None,
    generated_docx_files: list[GeneratedDocxFile],
    *,
    mode: str,
    bol_type: str | None = None,
    qty_type: str = "PLT",
    batch_comment: str | None = None,
    output_dir: Path | None = None,
    progress_callback: Callable[[int, int, GeneratedDocxFile], None] | None = None,
) -> StandardPdfConversionResult:
    """Generate Standard or No Recourse BOL PDFs without DOCX-to-PDF conversion."""

    if mode not in {"Standard", "No Recourse"}:
        raise ValueError("Direct Standard PDF generation supports only Standard and No Recourse modes.")
    if selected_facility is None:
        raise ValueError("No ship-from facility is selected. Select a facility before PDF generation.")
    if not generated_docx_files:
        raise ValueError("No generated DOCX files were provided for PDF naming.")

    output_root = output_dir or Path(mkdtemp(prefix="kkg_standard_bol_pdf_"))
    output_root.mkdir(parents=True, exist_ok=True)
    record_lookup = _records_by_bol(records)

    converted_files: list[ConvertedPdfFile] = []
    failed_conversions: list[FailedPdfConversion] = []
    total_files = len(generated_docx_files)

    for index, generated_file in enumerate(generated_docx_files, start=1):
        if progress_callback is not None:
            progress_callback(index, total_files, generated_file)

        source_docx = Path(generated_file.file_path)
        destination_pdf = output_root / f"{source_docx.stem}.pdf"
        record = record_lookup.get(_safe_text(generated_file.bol_number))
        if record is None:
            failed_conversions.append(
                FailedPdfConversion(
                    bol_number=generated_file.bol_number,
                    source_docx=str(source_docx),
                    error="Matching BOL record was not found for direct PDF generation.",
                )
            )
            continue

        try:
            _draw_bol_pdf(
                destination_pdf,
                record,
                selected_facility,
                mode=mode,
                bol_type=bol_type,
                qty_type=qty_type,
                batch_comment=batch_comment,
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
        converter_name="reportlab-direct",
        conversion_available=True,
        unavailable_reason=None,
        converter_path=None,
    )
