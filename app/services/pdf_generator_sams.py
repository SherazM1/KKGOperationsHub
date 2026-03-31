"""PDF generator for Sam's warehouse 4x6 labels."""

from __future__ import annotations

from io import BytesIO

from reportlab.graphics import renderPDF
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas

from app.models.sams_label import SamsLabel
from app.services.barcode_service import generate_code128_barcode
from app.utils.formatting import sanitize_text


PAGE_WIDTH = 4 * inch
PAGE_HEIGHT = 6 * inch

LEFT_MARGIN = 0.20 * inch
RIGHT_MARGIN = 0.20 * inch
TOP_MARGIN = 0.20 * inch
BOTTOM_MARGIN = 0.20 * inch
PRINT_WIDTH = PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN
CENTER_X = PAGE_WIDTH / 2


def _draw_wrapped(
    c: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    max_width: float,
    *,
    font_name: str = "Helvetica",
    font_size: float = 9,
    line_height: float = 10.5,
    max_lines: int = 2,
) -> float:
    clean = sanitize_text(text)
    if not clean:
        return y

    words = clean.split()
    line = ""
    lines: list[str] = []

    for word in words:
        candidate = f"{line} {word}".strip()
        if c.stringWidth(candidate, font_name, font_size) <= max_width:
            line = candidate
        else:
            if line:
                lines.append(line)
            line = word
            if len(lines) >= max_lines:
                break

    if line and len(lines) < max_lines:
        lines.append(line)

    c.setFont(font_name, font_size)
    for value in lines:
        c.drawString(x, y, value)
        y -= line_height

    return y


def _create_fitted_barcode(
    data: str,
    *,
    target_width: float,
    bar_height: float,
    max_bar_width: float,
    min_bar_width: float,
    step: float = 0.02,
):
    bar_width = max_bar_width
    best = generate_code128_barcode(data, bar_height=bar_height, bar_width=bar_width)

    while bar_width >= min_bar_width:
        candidate = generate_code128_barcode(
            data,
            bar_height=bar_height,
            bar_width=bar_width,
        )
        if candidate.width <= target_width:
            return candidate
        best = candidate
        bar_width -= step

    return best


def _draw_label_page(c: canvas.Canvas, label: SamsLabel) -> None:
    top_y = PAGE_HEIGHT - TOP_MARGIN
    col_gap = 0.14 * inch
    col_width = (PRINT_WIDTH - col_gap) / 2
    left_x = LEFT_MARGIN
    right_x = LEFT_MARGIN + col_width + col_gap
    divider_x = LEFT_MARGIN + col_width + (col_gap / 2)

    c.setStrokeColorRGB(0, 0, 0)
    c.setFillColorRGB(0, 0, 0)

    c.setFont("Helvetica-Bold", 9)
    c.drawString(left_x, top_y - 8, "SHIP FROM")
    c.drawString(right_x, top_y - 8, "SHIP TO")

    line_y = top_y - 20
    c.setFont("Helvetica", 9)
    c.drawString(left_x, line_y, sanitize_text(label.shipper_name))
    c.drawString(right_x, line_y, sanitize_text(label.ship_to_name))

    line_y -= 11
    left_end_y = _draw_wrapped(
        c,
        label.shipper_address,
        left_x,
        line_y,
        col_width,
        font_name="Helvetica",
        font_size=9,
        line_height=10.5,
        max_lines=2,
    )
    right_end_y = _draw_wrapped(
        c,
        label.ship_to_address,
        right_x,
        line_y,
        col_width,
        font_name="Helvetica",
        font_size=9,
        line_height=10.5,
        max_lines=2,
    )

    c.setFont("Helvetica", 9)
    c.drawString(
        left_x,
        left_end_y,
        f"{sanitize_text(label.shipper_city)}, {sanitize_text(label.shipper_state)} {label.shipper_zip}",
    )
    c.drawString(
        right_x,
        right_end_y,
        f"{sanitize_text(label.ship_to_city)}, {sanitize_text(label.ship_to_state)} {label.ship_to_zip}",
    )

    top_block_bottom_y = min(left_end_y, right_end_y) - 7
    c.setLineWidth(1.0)
    c.line(divider_x, top_y - 12, divider_x, top_block_bottom_y + 4)
    c.line(LEFT_MARGIN, top_block_bottom_y, PAGE_WIDTH - RIGHT_MARGIN, top_block_bottom_y)

    postal_barcode_value = "420" + label.ship_to_zip.replace("-", "")
    postal_barcode = _create_fitted_barcode(
        postal_barcode_value,
        target_width=PRINT_WIDTH * 0.90,
        bar_height=0.66 * inch,
        max_bar_width=1.55,
        min_bar_width=0.90,
    )
    postal_x = (PAGE_WIDTH - postal_barcode.width) / 2
    postal_bottom = top_block_bottom_y - 0.71 * inch
    renderPDF.draw(postal_barcode, c, postal_x, postal_bottom)

    postal_text = f"(420){label.ship_to_zip}"
    c.setFont("Helvetica", 8.5)
    postal_text_width = c.stringWidth(postal_text, "Helvetica", 8.5)
    c.drawString((PAGE_WIDTH - postal_text_width) / 2, postal_bottom - 8, postal_text)

    static_x = PAGE_WIDTH - RIGHT_MARGIN - 0.62 * inch
    static_y = postal_bottom + 0.58 * inch
    c.setFont("Helvetica-Bold", 9)
    c.drawString(static_x, static_y, "CLUB")
    c.drawString(static_x, static_y - 11, "PRO")
    c.drawString(static_x, static_y - 22, "B/L")

    middle_divider_y = postal_bottom - 14
    c.setLineWidth(0.95)
    c.line(LEFT_MARGIN, middle_divider_y, PAGE_WIDTH - RIGHT_MARGIN, middle_divider_y)

    middle_y = middle_divider_y - 10
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(LEFT_MARGIN, middle_y, "DC#")
    c.drawString(LEFT_MARGIN + 20, middle_y, sanitize_text(label.whse))
    c.drawString(LEFT_MARGIN + 52, middle_y, "TYPE")
    c.drawString(LEFT_MARGIN + 80, middle_y, sanitize_text(label.type_code))
    c.drawString(LEFT_MARGIN + 104, middle_y, "DEPT")
    c.drawString(LEFT_MARGIN + 135, middle_y, sanitize_text(label.dept))
    c.drawRightString(PAGE_WIDTH - RIGHT_MARGIN, middle_y, f"ORDER# {sanitize_text(label.po_number)}")

    middle_y -= 13
    c.setFont("Helvetica-Bold", 10)
    c.drawString(LEFT_MARGIN, middle_y, "WMIT:")
    c.setFont("Helvetica", 10)
    c.drawString(LEFT_MARGIN + 31, middle_y, sanitize_text(label.item_number))

    middle_y -= 11
    c.setFont("Helvetica", 9)
    desc = f"DESC: {sanitize_text(label.description)}"
    c.drawString(LEFT_MARGIN, middle_y, desc)
    c.setFont("Helvetica-Bold", 10)
    c.drawRightString(
        PAGE_WIDTH - RIGHT_MARGIN,
        middle_y,
        f"Qty {sanitize_text(label.quantity)}",
    )

    upc_barcode = _create_fitted_barcode(
        label.upc,
        target_width=PRINT_WIDTH * 0.94,
        bar_height=1.02 * inch,
        max_bar_width=1.50,
        min_bar_width=0.90,
    )
    upc_x = (PAGE_WIDTH - upc_barcode.width) / 2
    upc_bottom = BOTTOM_MARGIN + 0.34 * inch
    renderPDF.draw(upc_barcode, c, upc_x, upc_bottom)

    c.setFont("Helvetica", 8.5)
    upc_text_width = c.stringWidth(label.upc, "Helvetica", 8.5)
    c.drawString((PAGE_WIDTH - upc_text_width) / 2, upc_bottom - 11, label.upc)


def generate_sams_pdf(labels: list[SamsLabel]) -> bytes:
    if not labels:
        raise ValueError("No labels provided for PDF generation.")

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))

    for label in labels:
        for _ in range(2):
            _draw_label_page(c, label)
            c.showPage()

    c.save()
    buffer.seek(0)
    return buffer.read()
