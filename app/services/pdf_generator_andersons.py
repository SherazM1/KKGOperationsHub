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
LEFT_MARGIN = 11.735
RIGHT_MARGIN = 14.135
CENTER_X = PAGE_WIDTH / 2
PRINT_WIDTH = PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN
LABEL_RIGHT_X = 273.865


def _draw_line(c: canvas.Canvas, y: float, line_width: float = 1.0) -> None:
    c.setLineWidth(line_width)
    c.line(LEFT_MARGIN, y, LABEL_RIGHT_X, y)


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


def _draw_centered_fitted_string(
    c: canvas.Canvas,
    center_x: float,
    y: float,
    text: str,
    max_width: float,
    *,
    font_name: str = "Helvetica",
    font_size: float = 12,
    min_font_size: float = 6,
) -> None:
    clean = sanitize_text(text)
    current_size = font_size
    while (
        current_size > min_font_size
        and c.stringWidth(clean, font_name, current_size) > max_width
    ):
        current_size -= 0.5

    c.setFont(font_name, current_size)
    width = c.stringWidth(clean, font_name, current_size)
    c.drawString(center_x - (width / 2), y, clean)


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

    c.setFont("Helvetica-Bold", 11)
    c.drawString(12.358, 408.10, "SHIP FROM: KENDAL KING")

    _draw_underlined_string(c, 180.00, 408.10, "CLIENT", font_size=12, underline_gap=1.35)
    _draw_fitted_string(
        c,
        180.00,
        393.00,
        label.client,
        93.5,
        font_size=12,
        min_font_size=6,
    )

    c.setFont("Helvetica", 12)
    c.drawString(12.358, 390.09, "C/O:")
    _draw_fitted_string(c, 37.553, 390.00, ship_from["care_of"], 127.00, font_size=12)
    _draw_fitted_string(c, 12.454, 371.97, ship_from["address"], 151.00, font_size=12)
    c.drawString(
        12.454,
        353.97,
        (
            f"{sanitize_text(ship_from['city'])}, "
            f"{sanitize_text(ship_from['state'])} "
            f"{sanitize_text(ship_from['zip_code'])}"
        ),
    )

    c.setFont("Helvetica-Bold", 14)
    c.drawString(15.951, 325.16, "BRAND")
    _draw_fitted_string(c, 70.0, 325.16, label.brand, 196.0, font_size=12, min_font_size=6)

    c.setFont("Helvetica-Bold", 14)
    c.drawString(15.951, 301.13, "DESC")
    _draw_fitted_string(
        c,
        59.372,
        302.93,
        label.description,
        206.69,
        font_size=12,
        min_font_size=6,
    )

    c.setFont("Helvetica-Bold", 12)
    c.drawString(14.945, 279.02, "ORDER QTY")
    _draw_fitted_string(c, 87.273, 278.92, label.ordered_quantity, 31.76, font_size=12, min_font_size=6)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(124.97, 279.02, "UOM")
    _draw_fitted_string(c, 155.60, 278.92, label.unit_of_measure, 106.87, font_size=12, min_font_size=6)

    c.setFont("Helvetica-Bold", 12)
    c.drawString(12.646, 249.01, "PO NAME")
    _draw_wrapped_text(
        c,
        label.po_name,
        12.454,
        230.89,
        114.05,
        font_size=12,
        line_height=14.65,
        max_lines=2,
    )

    po_number_center_x = 205.575
    po_number_label = "PO NUMBER"
    _draw_underlined_string(c, 174.33, 249.01, po_number_label, font_size=12, underline_gap=1.37)

    po_barcode = _create_fitted_barcode(
        label.po_number,
        target_width=118.33,
        bar_height=49.87,
        max_bar_width=0.94,
        min_bar_width=0.40,
    )
    po_barcode_x = po_number_center_x - (po_barcode.width / 2)
    po_barcode_y = 193.37
    renderPDF.draw(po_barcode, c, po_barcode_x, po_barcode_y)

    c.setFillColorRGB(1, 1, 1)
    c.rect(144.63, 190.03, 121.74, 17.212, stroke=0, fill=1)
    c.setFillColorRGB(0, 0, 0)
    _draw_centered_fitted_string(
        c,
        po_number_center_x,
        194.88,
        label.po_number,
        82,
        font_name="Helvetica-Bold",
        font_size=12,
        min_font_size=6,
    )

    _draw_centered_underlined_string(
        c,
        CENTER_X,
        173.97,
        "UPC",
        font_size=12,
        underline_gap=1.37,
    )

    c.setLineWidth(0.7)
    c.line(11.855, 165.16, 268.905, 165.16)

    upc_barcode = _create_fitted_barcode(
        label.upc,
        target_width=253.51,
        bar_height=106.97,
        max_bar_width=2.12,
        min_bar_width=0.60,
    )
    upc_barcode_x = 15.927 + ((253.51 - upc_barcode.width) / 2)
    upc_barcode_y = 55.33
    renderPDF.draw(upc_barcode, c, upc_barcode_x, upc_barcode_y)

    c.setFillColorRGB(1, 1, 1)
    c.rect(11.951, 51.95, 261.58, 17.212, stroke=0, fill=1)
    c.setFillColorRGB(0, 0, 0)
    _draw_centered_fitted_string(
        c,
        CENTER_X,
        56.80,
        label.upc,
        100,
        font_name="Helvetica-Bold",
        font_size=12,
        min_font_size=6,
    )


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
