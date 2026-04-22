"""PDF generator for Sam's GCI 4x6 labels."""

from __future__ import annotations

from io import BytesIO
from typing import Any

from reportlab.graphics import renderPDF
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas

from app.models.sams_gci_label import SamsGciBottomRow, SamsGciPayload, SamsGciTopLabelRow
from app.services.barcode_service import generate_code128_barcode
from app.utils.formatting import sanitize_text


PAGE_WIDTH = 4 * inch
PAGE_HEIGHT = 6 * inch

LEFT_MARGIN = 0.16 * inch
RIGHT_MARGIN = 0.16 * inch
TOP_MARGIN = 0.15 * inch
BOTTOM_MARGIN = 0.15 * inch
PRINT_WIDTH = PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN


def _round_key(value: float) -> float:
    return round(value, 3)


def _normalize_bottom_rows(bottom_rows: list[SamsGciBottomRow]) -> list[dict[str, str]]:
    normalized_rows: list[dict[str, str]] = []
    for row in bottom_rows:
        normalized_rows.append(
            {
                "program_name": sanitize_text(row.program_name),
                "item_number": sanitize_text(row.item_number),
                "quantity": sanitize_text(row.quantity),
                "barcode_value": sanitize_text(row.barcode_value),
                "description": sanitize_text(row.description),
            }
        )
    return normalized_rows


def _get_wrapped_lines(
    c: canvas.Canvas,
    text: str,
    max_width: float,
    *,
    font_name: str = "Helvetica",
    font_size: float = 8.0,
    max_lines: int = 2,
    wrap_cache: dict[tuple[Any, ...], list[str]],
) -> list[str]:
    clean = sanitize_text(text)
    if not clean:
        return []

    cache_key = (clean, _round_key(max_width), font_name, _round_key(font_size), max_lines)
    cached = wrap_cache.get(cache_key)
    if cached is not None:
        return cached

    words = clean.split()
    lines: list[str] = []
    line = ""

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

    wrap_cache[cache_key] = lines
    return lines


def _draw_wrapped(
    c: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    max_width: float,
    *,
    font_name: str = "Helvetica",
    font_size: float = 8.0,
    line_height: float = 9.0,
    max_lines: int = 2,
    wrap_cache: dict[tuple[Any, ...], list[str]],
) -> float:
    lines = _get_wrapped_lines(
        c,
        text,
        max_width,
        font_name=font_name,
        font_size=font_size,
        max_lines=max_lines,
        wrap_cache=wrap_cache,
    )
    if not lines:
        return y

    c.setFont(font_name, font_size)
    for line in lines:
        c.drawString(x, y, line)
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
    barcode_cache: dict[tuple[Any, ...], Any],
):
    cache_key = (
        data,
        _round_key(target_width),
        _round_key(bar_height),
        _round_key(max_bar_width),
        _round_key(min_bar_width),
        _round_key(step),
    )
    cached = barcode_cache.get(cache_key)
    if cached is not None:
        return cached

    bar_width = max_bar_width
    best = None
    while bar_width >= min_bar_width - 1e-9:
        candidate = generate_code128_barcode(
            data,
            bar_height=bar_height,
            bar_width=bar_width,
        )
        if candidate.width <= target_width:
            barcode_cache[cache_key] = candidate
            return candidate
        best = candidate
        bar_width -= step

    if best is None:
        best = generate_code128_barcode(
            data,
            bar_height=bar_height,
            bar_width=min_bar_width,
        )
    barcode_cache[cache_key] = best
    return best


def _draw_top_section(
    c: canvas.Canvas,
    top_label: SamsGciTopLabelRow,
    *,
    barcode_cache: dict[tuple[Any, ...], Any],
    wrap_cache: dict[tuple[Any, ...], list[str]],
) -> float:
    top_y = PAGE_HEIGHT - TOP_MARGIN - 6

    col_gap = 0.12 * inch
    col_width = (PRINT_WIDTH - col_gap) / 2
    left_x = LEFT_MARGIN
    right_x = LEFT_MARGIN + col_width + col_gap
    divider_x = LEFT_MARGIN + col_width + (col_gap / 2)

    c.setStrokeColorRGB(0, 0, 0)
    c.setFillColorRGB(0, 0, 0)

    c.setFont("Helvetica-Bold", 8.8)
    c.drawString(left_x, top_y, "SHIP FROM")
    c.drawString(right_x, top_y, "SHIP TO")

    row_y = top_y - 11
    c.setFont("Helvetica", 8.2)
    c.drawString(left_x, row_y, sanitize_text(top_label.shipper_name))
    c.drawString(right_x, row_y, sanitize_text(top_label.ship_to_name))

    row_y -= 10
    left_end_y = _draw_wrapped(
        c,
        top_label.shipper_address,
        left_x,
        row_y,
        col_width,
        font_name="Helvetica",
        font_size=8.0,
        line_height=8.8,
        max_lines=2,
        wrap_cache=wrap_cache,
    )
    right_end_y = _draw_wrapped(
        c,
        top_label.ship_to_address,
        right_x,
        row_y,
        col_width,
        font_name="Helvetica",
        font_size=8.0,
        line_height=8.8,
        max_lines=2,
        wrap_cache=wrap_cache,
    )

    c.setFont("Helvetica", 8.2)
    c.drawString(
        left_x,
        left_end_y,
        f"{sanitize_text(top_label.shipper_city)}, "
        f"{sanitize_text(top_label.shipper_state)} {sanitize_text(top_label.shipper_zip)}",
    )
    c.drawString(
        right_x,
        right_end_y,
        f"{sanitize_text(top_label.ship_to_city)}, "
        f"{sanitize_text(top_label.ship_to_state)} {sanitize_text(top_label.ship_to_zip)}",
    )

    top_block_bottom_y = min(left_end_y, right_end_y) - 7
    c.setLineWidth(0.9)
    c.line(divider_x, top_y + 2, divider_x, top_block_bottom_y + 3)
    c.line(LEFT_MARGIN, top_block_bottom_y, PAGE_WIDTH - RIGHT_MARGIN, top_block_bottom_y)

    info_y = top_block_bottom_y - 8
    c.setFont("Helvetica-Bold", 8.4)
    c.drawString(LEFT_MARGIN, info_y, f"PO#: {sanitize_text(top_label.po_number)}")
    c.drawRightString(
        PAGE_WIDTH - RIGHT_MARGIN,
        info_y,
        f"CLUB#: {sanitize_text(top_label.club_display)}",
    )

    info_y -= 10
    c.setFont("Helvetica-Bold", 8.4)
    c.drawString(LEFT_MARGIN, info_y, "ITEM#:")
    c.setFont("Helvetica", 8.4)
    c.drawString(LEFT_MARGIN + 30, info_y, sanitize_text(top_label.item_number))
    c.setFont("Helvetica-Bold", 8.4)
    c.drawRightString(
        PAGE_WIDTH - RIGHT_MARGIN,
        info_y,
        f"QTY: {sanitize_text(top_label.quantity)}",
    )

    info_y -= 10
    c.setFont("Helvetica", 7.8)
    desc_label = f"DESC: {sanitize_text(top_label.description)}"
    _draw_wrapped(
        c,
        desc_label,
        LEFT_MARGIN,
        info_y,
        PRINT_WIDTH,
        font_name="Helvetica",
        font_size=7.8,
        line_height=8.6,
        max_lines=2,
        wrap_cache=wrap_cache,
    )

    barcode_value = sanitize_text(top_label.top_barcode_value)
    if barcode_value:
        top_barcode = _create_fitted_barcode(
            barcode_value,
            target_width=PRINT_WIDTH * 0.94,
            bar_height=0.46 * inch,
            max_bar_width=1.24,
            min_bar_width=0.64,
            barcode_cache=barcode_cache,
        )
        top_barcode_x = (PAGE_WIDTH - top_barcode.width) / 2
        top_barcode_bottom = info_y - 39
        renderPDF.draw(top_barcode, c, top_barcode_x, top_barcode_bottom)
        c.setFont("Helvetica-Bold", 7.8)
        c.drawCentredString(PAGE_WIDTH / 2, top_barcode_bottom + (0.46 * inch) + 2.5, "HOLGCPLT")
        c.setFont("Helvetica", 8.8)
        c.drawCentredString(PAGE_WIDTH / 2, top_barcode_bottom - 11, barcode_value)
        return top_barcode_bottom - 14

    return info_y - 6


def _draw_bottom_rows(
    c: canvas.Canvas,
    bottom_rows: list[dict[str, str]],
    start_y: float,
    *,
    barcode_cache: dict[tuple[Any, ...], Any],
    wrap_cache: dict[tuple[Any, ...], list[str]],
) -> None:
    c.setLineWidth(0.8)
    c.line(LEFT_MARGIN, start_y, PAGE_WIDTH - RIGHT_MARGIN, start_y)

    row_count = len(bottom_rows)
    if row_count <= 0:
        return

    available_height = max(18.0, start_y - BOTTOM_MARGIN - 2.0)
    row_block_height = available_height / row_count
    # Keep repeated rows compact while still fitting high row counts.
    row_block_height = max(13.0, min(row_block_height, 27.0))

    inner_pad_x = 2.0
    row_gap = 4.5
    text_region_width = PRINT_WIDTH * 0.44
    text_left_x = LEFT_MARGIN + inner_pad_x
    text_right_x = LEFT_MARGIN + text_region_width - inner_pad_x
    barcode_left_x = text_right_x + row_gap
    barcode_right_x = PAGE_WIDTH - RIGHT_MARGIN - inner_pad_x

    y = start_y
    for row in bottom_rows:
        row_top = y
        row_bottom = _draw_bottom_row_box(
            c,
            row,
            row_top=row_top,
            row_height=row_block_height,
            text_left_x=text_left_x,
            text_right_x=text_right_x,
            barcode_left_x=barcode_left_x,
            barcode_right_x=barcode_right_x,
            barcode_cache=barcode_cache,
            wrap_cache=wrap_cache,
        )
        c.setLineWidth(0.55)
        c.line(LEFT_MARGIN, row_bottom, PAGE_WIDTH - RIGHT_MARGIN, row_bottom)
        y = row_bottom


def _draw_bottom_row_box(
    c: canvas.Canvas,
    row: dict[str, str],
    *,
    row_top: float,
    row_height: float,
    text_left_x: float,
    text_right_x: float,
    barcode_left_x: float,
    barcode_right_x: float,
    barcode_cache: dict[tuple[Any, ...], Any],
    wrap_cache: dict[tuple[Any, ...], list[str]],
) -> float:
    row_bottom = row_top - row_height

    # Explicit row container bounds and internal padding.
    pad_top = 2.0
    pad_bottom = 2.0
    content_top = row_top - pad_top
    content_bottom = row_bottom + pad_bottom

    text_width = max(20.0, text_right_x - text_left_x)
    barcode_width = max(20.0, barcode_right_x - barcode_left_x)

    item_value = row["item_number"]
    quantity_value = row["quantity"]
    desc_value = row["description"]
    barcode_value = row["barcode_value"]

    # Left text block: ITEM/QTY line plus compact description lines.
    title_y = content_top - 0.3
    c.setFont("Helvetica-Bold", 7.0)
    c.drawString(text_left_x, title_y, f"ITEM#: {item_value}")
    c.drawRightString(text_right_x, title_y, f"QTY: {quantity_value}")

    desc_y = title_y - 6.6
    desc_line_height = 6.5
    # Keep description inside row bounds using available vertical space.
    desc_max_lines = 1
    if row_height >= 24:
        desc_max_lines = 2
    c.setFont("Helvetica", 6.9)
    _draw_wrapped(
        c,
        desc_value,
        text_left_x,
        desc_y,
        text_width,
        font_name="Helvetica",
        font_size=6.9,
        line_height=desc_line_height,
        max_lines=desc_max_lines,
        wrap_cache=wrap_cache,
    )

    # Right barcode block: barcode and human-readable text fully inside row container.
    if barcode_value:
        human_text_height = 6.0
        barcode_gap = 1.1
        barcode_target_width = max(36.0, barcode_width - 1.0)

        human_text_y = content_bottom + 0.5
        barcode_bottom = human_text_y + human_text_height + barcode_gap
        max_barcode_height = max(4.5, content_top - barcode_bottom - 0.3)
        barcode_height = min(0.30 * inch, max_barcode_height)

        row_barcode = _create_fitted_barcode(
            barcode_value,
            target_width=barcode_target_width,
            bar_height=barcode_height,
            max_bar_width=1.20,
            min_bar_width=0.42,
            barcode_cache=barcode_cache,
        )
        row_barcode_x = barcode_left_x + (barcode_width - row_barcode.width) / 2
        row_barcode_bottom = barcode_bottom
        renderPDF.draw(row_barcode, c, row_barcode_x, row_barcode_bottom)

        c.setFont("Helvetica", 6.3)
        c.drawCentredString(
            barcode_left_x + (barcode_width / 2),
            human_text_y,
            barcode_value,
        )

    return row_bottom


def _draw_gci_label_page(
    c: canvas.Canvas,
    top_label: SamsGciTopLabelRow,
    bottom_rows: list[dict[str, str]],
    *,
    barcode_cache: dict[tuple[Any, ...], Any],
    wrap_cache: dict[tuple[Any, ...], list[str]],
) -> None:
    bottom_start_y = _draw_top_section(
        c,
        top_label,
        barcode_cache=barcode_cache,
        wrap_cache=wrap_cache,
    )
    _draw_bottom_rows(
        c,
        bottom_rows,
        bottom_start_y,
        barcode_cache=barcode_cache,
        wrap_cache=wrap_cache,
    )


def _top_label_form_key(top_label: SamsGciTopLabelRow) -> tuple[str, ...]:
    return (
        top_label.shipper_name,
        top_label.shipper_address,
        top_label.shipper_city,
        top_label.shipper_state,
        top_label.shipper_zip,
        top_label.ship_to_name,
        top_label.ship_to_address,
        top_label.ship_to_city,
        top_label.ship_to_state,
        top_label.ship_to_zip,
        top_label.po_number,
        top_label.club_number,
        top_label.whse,
        top_label.item_number,
        top_label.description,
        top_label.quantity,
    )


def generate_sams_gci_pdf(payload: SamsGciPayload) -> bytes:
    """Generate Sam's GCI label PDF bytes with 2 pages per MDG row."""
    if not payload.mdg_labels:
        raise ValueError("No MDG labels provided for GCI PDF generation.")
    if not payload.bottom_rows:
        raise ValueError("No GCI bottom rows provided for GCI PDF generation.")

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))
    barcode_cache: dict[tuple[Any, ...], Any] = {}
    wrap_cache: dict[tuple[Any, ...], list[str]] = {}
    bottom_rows = _normalize_bottom_rows(payload.bottom_rows)
    form_name_by_label: dict[tuple[str, ...], str] = {}

    for top_label in payload.mdg_labels:
        form_key = _top_label_form_key(top_label)
        form_name = form_name_by_label.get(form_key)
        if form_name is None:
            form_name = f"gci_label_form_{len(form_name_by_label) + 1}"
            c.beginForm(form_name, 0, 0, PAGE_WIDTH, PAGE_HEIGHT)
            _draw_gci_label_page(
                c,
                top_label,
                bottom_rows,
                barcode_cache=barcode_cache,
                wrap_cache=wrap_cache,
            )
            c.endForm()
            form_name_by_label[form_key] = form_name

        for _ in range(2):
            c.doForm(form_name)
            c.showPage()

    c.save()
    buffer.seek(0)
    return buffer.read()
