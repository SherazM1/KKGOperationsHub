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
        c.setFont("Helvetica", 8.8)
        c.drawCentredString(PAGE_WIDTH / 2, top_barcode_bottom - 11, barcode_value)
        return top_barcode_bottom - 20

    return info_y - 6


def _draw_bottom_rows(
    c: canvas.Canvas,
    bottom_rows: list[dict[str, str]],
    start_y: float,
    *,
    barcode_cache: dict[tuple[Any, ...], Any],
    wrap_cache: dict[tuple[Any, ...], list[str]],
) -> None:
    c.setFont("Helvetica-Bold", 7.8)
    c.drawCentredString(PAGE_WIDTH / 2, start_y + 4.2, "HOLGCPLT")

    c.setLineWidth(0.8)
    c.line(LEFT_MARGIN, start_y, PAGE_WIDTH - RIGHT_MARGIN, start_y)

    row_count = len(bottom_rows)
    if row_count <= 0:
        return

    lower_bottom = BOTTOM_MARGIN
    available_height = max(10.0, start_y - lower_bottom)
    row_height = available_height / row_count

    inner_pad_x = 2.2
    column_gap = 4.2
    usable_left = LEFT_MARGIN + inner_pad_x
    usable_right = PAGE_WIDTH - RIGHT_MARGIN - inner_pad_x
    usable_width = max(40.0, usable_right - usable_left)
    text_width = usable_width * 0.36
    text_left_x = usable_left
    text_right_x = text_left_x + text_width
    barcode_left_x = text_right_x + column_gap
    barcode_right_x = usable_right

    for index, row in enumerate(bottom_rows):
        row_top = start_y - (index * row_height)
        row_bottom = start_y - ((index + 1) * row_height)
        _draw_bottom_row_box(
            c,
            row,
            row_top=row_top,
            row_bottom=row_bottom,
            text_left_x=text_left_x,
            text_right_x=text_right_x,
            barcode_left_x=barcode_left_x,
            barcode_right_x=barcode_right_x,
            barcode_cache=barcode_cache,
            wrap_cache=wrap_cache,
        )
        c.setLineWidth(0.55)
        c.line(LEFT_MARGIN, row_bottom, PAGE_WIDTH - RIGHT_MARGIN, row_bottom)


def _draw_bottom_row_box(
    c: canvas.Canvas,
    row: dict[str, str],
    *,
    row_top: float,
    row_bottom: float,
    text_left_x: float,
    text_right_x: float,
    barcode_left_x: float,
    barcode_right_x: float,
    barcode_cache: dict[tuple[Any, ...], Any],
    wrap_cache: dict[tuple[Any, ...], list[str]],
) -> float:
    row_height = max(8.0, row_top - row_bottom)
    row_pad_y = max(1.8, min(5.6, row_height * 0.14))
    inner_top = row_top - row_pad_y
    inner_bottom = row_bottom + row_pad_y
    if inner_top - inner_bottom < 5.0:
        inner_top = row_top - 1.2
        inner_bottom = row_bottom + 1.2

    inner_height = max(4.0, inner_top - inner_bottom)
    text_width = max(20.0, text_right_x - text_left_x)
    barcode_width = max(20.0, barcode_right_x - barcode_left_x)

    item_value = row["item_number"]
    quantity_value = row["quantity"]
    desc_value = row["description"]
    barcode_value = row["barcode_value"]

    # Left text block stays inside row bounds with top padding.
    title_font = 7.0 if row_height < 26 else 7.4
    title_y = inner_top - (title_font * 0.9)
    c.setFont("Helvetica-Bold", title_font)
    c.drawString(text_left_x, title_y, f"ITEM#: {item_value}")
    c.drawRightString(text_right_x, title_y, f"QTY: {quantity_value}")

    desc_font = 6.6 if row_height < 23 else 7.0
    desc_line_height = desc_font + 1.1
    desc_y = title_y - max(4.8, title_font * 0.92)
    desc_max_lines = 2 if (desc_y - inner_bottom) >= (desc_line_height * 1.65) else 1
    c.setFont("Helvetica", desc_font)
    _draw_wrapped(
        c,
        desc_value,
        text_left_x,
        desc_y,
        text_width,
        font_name="Helvetica",
        font_size=desc_font,
        line_height=desc_line_height,
        max_lines=desc_max_lines,
        wrap_cache=wrap_cache,
    )

    # Right barcode block is wider and uses more of each row.
    if barcode_value:
        barcode_target_width = max(42.0, barcode_width - 0.6)
        human_font = 6.5 if row_height < 25 else 6.8
        human_height = human_font + 1.1
        barcode_gap = 1.0
        human_text_y = inner_bottom + 0.35
        barcode_bottom = human_text_y + human_height + barcode_gap

        max_barcode_height = max(3.6, inner_top - barcode_bottom - 0.3)
        preferred_barcode_height = min(0.44 * inch, max(0.29 * inch, inner_height * 0.68))
        barcode_height = min(preferred_barcode_height, max_barcode_height)

        row_barcode = _create_fitted_barcode(
            barcode_value,
            target_width=barcode_target_width,
            bar_height=barcode_height,
            max_bar_width=1.48,
            min_bar_width=0.46,
            barcode_cache=barcode_cache,
        )
        row_barcode_x = barcode_left_x + (barcode_width - row_barcode.width) / 2
        row_barcode_bottom = barcode_bottom
        renderPDF.draw(row_barcode, c, row_barcode_x, row_barcode_bottom)

        c.setFont("Helvetica", human_font)
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
