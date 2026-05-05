"""PDF generator for Andersons labels."""

from __future__ import annotations

from io import BytesIO

from reportlab.graphics import renderPDF
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas

from app.models.andersons_label import AndersonsLabel
from app.services.barcode_service import generate_code128_barcode
from app.utils.formatting import sanitize_text


PAGE_WIDTH = 4 * inch
PAGE_HEIGHT = 6 * inch
LEFT_MARGIN = 14
RIGHT_MARGIN = 14
CENTER_X = PAGE_WIDTH / 2
PRINT_WIDTH = PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN


def _draw_line(c: canvas.Canvas, y: float, line_width: float = 1.0) -> None:
    c.setLineWidth(line_width)
    c.line(LEFT_MARGIN, y, PAGE_WIDTH - RIGHT_MARGIN, y)


def _draw_underlined_string(
    c: canvas.Canvas,
    x: float,
    y: float,
    text: str,
    *,
    font_name: str = "Helvetica-Bold",
    font_size: float = 10,
    underline_gap: float = 2,
) -> None:
    c.setFont(font_name, font_size)
    clean = sanitize_text(text)
    c.drawString(x, y, clean)
    width = c.stringWidth(clean, font_name, font_size)
    c.setLineWidth(0.8)
    c.line(x, y - underline_gap, x + width, y - underline_gap)


def _draw_centered_underlined_string(
    c: canvas.Canvas,
    center_x: float,
    y: float,
    text: str,
    *,
    font_name: str = "Helvetica-Bold",
    font_size: float = 11,
    underline_gap: float = 2,
) -> None:
    clean = sanitize_text(text)
    width = c.stringWidth(clean, font_name, font_size)
    _draw_underlined_string(
        c,
        center_x - (width / 2),
        y,
        clean,
        font_name=font_name,
        font_size=font_size,
        underline_gap=underline_gap,
    )


def _draw_wrapped_text(
    c: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    max_width: float,
    *,
    font_name: str = "Helvetica",
    font_size: float = 8.5,
    line_height: float = 10,
    max_lines: int = 2,
) -> float:
    clean = sanitize_text(text)
    if not clean:
        return y

    words = clean.split()
    lines: list[str] = []
    current = ""

    for word in words:
        candidate = f"{current} {word}".strip()
        if c.stringWidth(candidate, font_name, font_size) <= max_width:
            current = candidate
        else:
            if current:
                lines.append(current)
            current = word
            if len(lines) >= max_lines:
                break

    if current and len(lines) < max_lines:
        lines.append(current)

    c.setFont(font_name, font_size)
    for line in lines:
        c.drawString(x, y, line)
        y -= line_height

    return y


def _draw_fitted_string(
    c: canvas.Canvas,
    x: float,
    y: float,
    text: str,
    max_width: float,
    *,
    font_name: str = "Helvetica",
    font_size: float = 9,
    min_font_size: float = 6,
) -> None:
    clean = sanitize_text(text)
    current_size = font_size
    while current_size > min_font_size and c.stringWidth(clean, font_name, current_size) > max_width:
        current_size -= 0.5

    c.setFont(font_name, current_size)
    c.drawString(x, y, clean)


def _create_fitted_barcode(
    data: str,
    *,
    target_width: float,
    bar_height: float,
    max_bar_width: float,
    min_bar_width: float,
):
    bar_width = max_bar_width
    best = generate_code128_barcode(data, bar_height=bar_height, bar_width=bar_width)

    while bar_width >= min_bar_width:
        candidate = generate_code128_barcode(data, bar_height=bar_height, bar_width=bar_width)
        if candidate.width <= target_width:
            return candidate
        best = candidate
        bar_width -= 0.02

    return best


def _draw_label_value(
    c: canvas.Canvas,
    label_text: str,
    value: str,
    x: float,
    y: float,
    value_x: float,
    max_width: float,
) -> None:
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(x, y, label_text)
    _draw_fitted_string(c, value_x, y, value, max_width, font_size=9, min_font_size=6)


def _draw_label_page(
    c: canvas.Canvas,
    label: AndersonsLabel,
    ship_from: dict[str, str],
) -> None:
    c.setFillColorRGB(0, 0, 0)
    c.setStrokeColorRGB(0, 0, 0)

    top_y = PAGE_HEIGHT - 19
    c.setFont("Helvetica-Bold", 10)
    c.drawString(LEFT_MARGIN, top_y, "SHIP FROM: KENDAL KING")

    client_x = PAGE_WIDTH - RIGHT_MARGIN - 78
    _draw_underlined_string(c, client_x, top_y, "CLIENT", font_size=10)
    _draw_fitted_string(
        c,
        client_x,
        top_y - 15,
        label.client,
        PAGE_WIDTH - RIGHT_MARGIN - client_x,
        font_size=8.5,
        min_font_size=6,
    )

    c.setFont("Helvetica", 8.8)
    c.drawString(LEFT_MARGIN, top_y - 19, f"C/O: {sanitize_text(ship_from['care_of'])}")
    c.drawString(LEFT_MARGIN, top_y - 33, sanitize_text(ship_from["address"]))
    c.drawString(
        LEFT_MARGIN,
        top_y - 47,
        (
            f"{sanitize_text(ship_from['city'])}, "
            f"{sanitize_text(ship_from['state'])} "
            f"{sanitize_text(ship_from['zip_code'])}"
        ),
    )

    _draw_line(c, top_y - 62, 0.9)

    field_y = top_y - 83
    _draw_label_value(c, "BRAND", label.brand, LEFT_MARGIN, field_y, LEFT_MARGIN + 48, 102)

    desc_y = field_y - 27
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(LEFT_MARGIN, desc_y, "DESC")
    _draw_wrapped_text(
        c,
        label.description,
        LEFT_MARGIN + 42,
        desc_y,
        PRINT_WIDTH - 42,
        font_size=8.5,
        line_height=9.8,
        max_lines=2,
    )

    qty_y = desc_y - 43
    _draw_label_value(
        c,
        "ORDER QTY",
        label.ordered_quantity,
        LEFT_MARGIN,
        qty_y,
        LEFT_MARGIN + 64,
        54,
    )
    _draw_label_value(
        c,
        "UOM",
        label.unit_of_measure,
        LEFT_MARGIN + 154,
        qty_y,
        LEFT_MARGIN + 185,
        72,
    )

    _draw_line(c, qty_y - 16, 0.9)

    po_top_y = qty_y - 36
    c.setFont("Helvetica-Bold", 10)
    c.drawString(LEFT_MARGIN, po_top_y, "PO NAME")
    _draw_wrapped_text(
        c,
        label.po_name,
        LEFT_MARGIN,
        po_top_y - 14,
        122,
        font_size=8.3,
        line_height=9.5,
        max_lines=2,
    )

    po_number_center_x = LEFT_MARGIN + 197
    po_number_label = "PO NUMBER"
    po_number_width = c.stringWidth(po_number_label, "Helvetica-Bold", 10)
    _draw_underlined_string(
        c,
        po_number_center_x - (po_number_width / 2),
        po_top_y,
        po_number_label,
        font_size=10,
    )

    po_barcode = _create_fitted_barcode(
        label.po_number,
        target_width=116,
        bar_height=34,
        max_bar_width=0.86,
        min_bar_width=0.44,
    )
    po_barcode_x = po_number_center_x - (po_barcode.width / 2)
    po_barcode_y = po_top_y - 50
    renderPDF.draw(po_barcode, c, po_barcode_x, po_barcode_y)

    c.setFont("Helvetica", 7.8)
    po_text = sanitize_text(label.po_number)
    po_text_width = c.stringWidth(po_text, "Helvetica", 7.8)
    c.drawString(po_number_center_x - (po_text_width / 2), po_barcode_y - 10, po_text)

    upc_divider_y = 133
    _draw_line(c, upc_divider_y, 0.9)

    upc_label_y = upc_divider_y - 21
    _draw_centered_underlined_string(c, CENTER_X, upc_label_y, "UPC", font_size=11)

    upc_barcode = _create_fitted_barcode(
        label.upc,
        target_width=PRINT_WIDTH * 0.96,
        bar_height=55,
        max_bar_width=1.28,
        min_bar_width=0.58,
    )
    upc_barcode_x = (PAGE_WIDTH - upc_barcode.width) / 2
    upc_barcode_y = 44
    renderPDF.draw(upc_barcode, c, upc_barcode_x, upc_barcode_y)

    c.setFont("Helvetica", 9)
    upc_text = sanitize_text(label.upc)
    upc_text_width = c.stringWidth(upc_text, "Helvetica", 9)
    c.drawString(CENTER_X - (upc_text_width / 2), upc_barcode_y - 13, upc_text)


def generate_andersons_pdf(
    labels: list[AndersonsLabel],
    ship_from: dict[str, str],
) -> bytes:
    if not labels:
        raise ValueError("No labels provided for PDF generation.")

    required_ship_from = {"care_of", "address", "city", "state", "zip_code"}
    missing_ship_from = sorted(required_ship_from - set(ship_from))
    if missing_ship_from:
        raise ValueError("Ship From is missing: " + ", ".join(missing_ship_from))

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))

    for label in labels:
        _draw_label_page(c, label, ship_from)
        c.showPage()

    c.save()
    buffer.seek(0)
    return buffer.read()
