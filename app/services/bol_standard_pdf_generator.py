"""Direct PDF generation for Standard-family BOL records."""

from __future__ import annotations

from pathlib import Path
from tempfile import mkdtemp
from typing import Any, Callable

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph

from app.models.bol_standard_record import BolStandardItemLine, BolStandardRecord
from app.services.bol_standard_docx_generator import (
    GeneratedDocxFile,
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
MARGIN = 0.32 * inch
FONT_NAME = "Helvetica"
FONT_BOLD = "Helvetica-Bold"
LINE_COLOR = colors.black
HEADER_FILL = colors.HexColor("#D9D9D9")
LIGHT_FILL = colors.HexColor("#F2F2F2")


def _safe_text(value: Any) -> str:
    return str(value or "").strip()


def _normalize_bol_type(bol_type: str | None) -> str:
    normalized = (bol_type or "PLT").strip().upper()
    return normalized if normalized in {"PLT", "CASE"} else "PLT"


def _resolve_comment(record: BolStandardRecord, batch_comment: str | None) -> str:
    record_comment = _safe_text(record.comments)
    return record_comment if record_comment else _safe_text(batch_comment)


def _display_number(value: float) -> str:
    return _format_number(value)


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
        _display_number(total_pallet_qty_value)
        if filter_blank_item_lines and has_pallet_qty_value
        else _display_number(total_qty)
    )
    return (
        total_qty_display,
        _display_number(total_skids_value),
        _display_number(total_weight_value),
    )


def _fit_font_size(
    canv: canvas.Canvas,
    text: str,
    font_name: str,
    base_size: float,
    max_width: float,
    min_size: float = 6.0,
) -> float:
    size = base_size
    while size > min_size and canv.stringWidth(text, font_name, size) > max_width:
        size -= 0.5
    return size


def _draw_string_fit(
    canv: canvas.Canvas,
    x: float,
    y: float,
    text: str,
    max_width: float,
    *,
    font_name: str = FONT_NAME,
    font_size: float = 8,
    align: str = "left",
    min_size: float = 6,
) -> None:
    text = _safe_text(text)
    size = _fit_font_size(canv, text, font_name, font_size, max_width, min_size)
    canv.setFont(font_name, size)
    if align == "right":
        canv.drawRightString(x + max_width, y, text)
    elif align == "center":
        canv.drawCentredString(x + max_width / 2, y, text)
    else:
        canv.drawString(x, y, text)


def _paragraph_style(
    name: str,
    *,
    font_name: str = FONT_NAME,
    font_size: float = 7.5,
    leading: float | None = None,
    alignment: int = TA_LEFT,
) -> ParagraphStyle:
    return ParagraphStyle(
        name=name,
        fontName=font_name,
        fontSize=font_size,
        leading=leading or font_size + 1.2,
        alignment=alignment,
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


def _draw_box(canv: canvas.Canvas, x: float, y: float, width: float, height: float) -> None:
    canv.setStrokeColor(LINE_COLOR)
    canv.setLineWidth(0.7)
    canv.rect(x, y, width, height, stroke=1, fill=0)


def _draw_label_cell(
    canv: canvas.Canvas,
    x: float,
    y: float,
    width: float,
    height: float,
    label: str,
    *,
    fill: bool = True,
    align: str = "left",
) -> None:
    if fill:
        canv.setFillColor(HEADER_FILL)
        canv.rect(x, y, width, height, stroke=0, fill=1)
        canv.setFillColor(colors.black)
    _draw_box(canv, x, y, width, height)
    _draw_string_fit(
        canv,
        x + 3,
        y + height - 9,
        label,
        width - 6,
        font_name=FONT_BOLD,
        font_size=7.2,
        align=align,
    )


def _draw_value_cell(
    canv: canvas.Canvas,
    x: float,
    y: float,
    width: float,
    height: float,
    value: str,
    *,
    font_size: float = 7.5,
    bold: bool = False,
) -> None:
    _draw_box(canv, x, y, width, height)
    _draw_string_fit(
        canv,
        x + 3,
        y + height - 9,
        value,
        width - 6,
        font_name=FONT_BOLD if bold else FONT_NAME,
        font_size=font_size,
    )


def _draw_checkbox(canv: canvas.Canvas, x: float, y: float, label: str, *, checked: bool = False) -> None:
    size = 7
    canv.rect(x, y, size, size, stroke=1, fill=0)
    if checked:
        canv.line(x + 1.5, y + 3, x + 3, y + 1.5)
        canv.line(x + 3, y + 1.5, x + 6, y + 6)
    canv.setFont(FONT_NAME, 6.5)
    canv.drawString(x + size + 3, y + 1, label)


def _draw_header(
    canv: canvas.Canvas,
    record: BolStandardRecord,
    mode: str,
    *,
    resolved_comment: str,
) -> None:
    top = PAGE_HEIGHT - MARGIN
    left_width = 4.2 * inch
    right_x = MARGIN + left_width + 0.18 * inch
    right_width = PAGE_WIDTH - MARGIN - right_x

    canv.setFont(FONT_NAME, 7.5)
    canv.drawString(MARGIN, top - 10, "609 SW 8th St • Ste 140 • Bentonville, AR 72712")
    canv.setFont(FONT_BOLD, 15)
    title = "UNIFORM BILL OF LADING" if mode == "No Recourse" else "STRAIGHT BILL OF LADING"
    canv.drawCentredString(MARGIN + left_width / 2, top - 33, title)
    canv.setFont(FONT_NAME, 7)
    canv.drawCentredString(MARGIN + left_width / 2, top - 46, "Original - Not Negotiable")

    row_h = 14
    label_w = 0.82 * inch
    value_w = right_width - label_w
    rows = [
        ("BOL #", record.bol_number),
        ("Ship Date", _format_ship_date_for_template(record.ship_date)),
        ("Carrier", record.carrier),
        ("Carrier Pro #", record.carrier_pro_number or record.kk_load_number),
        ("PO #", record.po_number),
        ("KK PO #", record.kk_po_number),
        ("KK Load #", record.kk_load_number),
        ("Seal #", record.seal_number_blank),
        ("Pick Up #", getattr(record, "pickup_number", "")),
        ("Comments", resolved_comment),
    ]
    y = top - row_h
    for label, value in rows:
        _draw_label_cell(canv, right_x, y, label_w, row_h, label, fill=False, align="right")
        _draw_value_cell(canv, right_x + label_w, y, value_w, row_h, value, font_size=7.5)
        y -= row_h


def _draw_party_blocks(
    canv: canvas.Canvas,
    record: BolStandardRecord,
    selected_facility: BolFacilityRecord,
) -> None:
    x = MARGIN
    y_top = PAGE_HEIGHT - MARGIN - 1.95 * inch
    width = PAGE_WIDTH - (2 * MARGIN)
    block_h = 0.96 * inch
    half_w = width / 2

    for offset, title in ((0, "FROM (SHIPPER)"), (half_w, "TO (CONSIGNEE)")):
        _draw_label_cell(canv, x + offset, y_top - 15, half_w, 15, title)
        _draw_box(canv, x + offset, y_top - block_h, half_w, block_h - 15)

    label_w = 0.72 * inch
    row_h = 13.5
    ship_values = [
        ("COMPANY", selected_facility["facility_name"]),
        ("STREET", selected_facility["address"]),
        ("CITY/ST/ZIP", selected_facility["location"]),
    ]
    consignee_values = [
        ("COMPANY", record.consignee_company),
        ("STREET", record.consignee_street),
        ("CITY/ST/ZIP", record.consignee_city_state_zip),
        ("DC #", record.dc_number),
    ]
    for col_x, values in ((x, ship_values), (x + half_w, consignee_values)):
        row_y = y_top - 30
        for label, value in values:
            canv.setFont(FONT_BOLD, 6.7)
            canv.drawString(col_x + 4, row_y, label)
            _draw_string_fit(canv, col_x + label_w, row_y, value, half_w - label_w - 8, font_size=7.5)
            row_y -= row_h


def _draw_terms_subject_billto(canv: canvas.Canvas, record: BolStandardRecord) -> None:
    x = MARGIN
    y = PAGE_HEIGHT - MARGIN - 3.08 * inch
    width = PAGE_WIDTH - (2 * MARGIN)
    terms_h = 0.44 * inch
    subject_h = 0.48 * inch
    bill_w = 2.3 * inch
    subject_w = width - bill_w - 0.08 * inch

    _draw_box(canv, x, y - terms_h, width, terms_h)
    canv.setFont(FONT_BOLD, 7)
    canv.drawString(x + 5, y - 11, "FREIGHT CHARGE TERMS")
    _draw_checkbox(canv, x + 130, y - 15, "Prepaid", checked=True)
    _draw_checkbox(canv, x + 195, y - 15, "Collect")
    _draw_checkbox(canv, x + 255, y - 15, "3rd Party")
    canv.setFont(FONT_NAME, 6.5)
    canv.drawString(x + 340, y - 13, "Freight charges are subject to all lawful tariffs and classifications.")

    subject_y = y - terms_h - 0.06 * inch
    _draw_box(canv, x, subject_y - subject_h, subject_w, subject_h)
    subject = (
        "SUBJECT TO SECTION 7: Of the conditions if shipment is to be delivered to consignee "
        "without recourse on the consignor, the consignor shall sign the following statement: "
        "The Carrier shall not make delivery of the shipment without payment of the freight "
        "and all other lawful charges."
    )
    _draw_paragraph(
        canv,
        subject,
        x + 5,
        subject_y - 4,
        subject_w - 10,
        subject_h - 8,
        style=_paragraph_style("Subject7", font_size=6.0, leading=6.8),
    )

    bill_x = x + subject_w + 0.08 * inch
    _draw_label_cell(canv, bill_x, subject_y - 14, bill_w, 14, "BILL TO:")
    _draw_box(canv, bill_x, subject_y - subject_h, bill_w, subject_h - 14)
    bill_lines = [
        record.bill_to.company,
        record.bill_to.street,
        record.bill_to.city_state_zip,
        "Attn:",
    ]
    line_y = subject_y - 27
    for line in bill_lines:
        _draw_string_fit(canv, bill_x + 6, line_y, line, bill_w - 12, font_size=7.2)
        line_y -= 8.5


def _draw_item_table(
    canv: canvas.Canvas,
    record: BolStandardRecord,
    *,
    mode: str,
    bol_type: str | None,
    qty_type: str,
) -> float:
    rendered_type = _normalize_bol_type(bol_type)
    filter_blank = mode == "No Recourse"
    item_lines = _rendered_item_lines(record.item_lines, filter_blank_item_lines=filter_blank)
    total_qty, total_skids, total_weight = _calculate_totals(
        item_lines,
        record.total_skids,
        filter_blank_item_lines=filter_blank,
    )

    x = MARGIN
    top = PAGE_HEIGHT - MARGIN - 4.15 * inch
    width = PAGE_WIDTH - (2 * MARGIN)
    header_h = 18
    row_h = 24 if mode == "No Recourse" else 27
    totals_h = 18
    min_rows = 4 if mode == "No Recourse" else 5
    rows_to_draw = max(min_rows, len(item_lines))
    table_h = header_h + (rows_to_draw * row_h) + totals_h

    col_widths = [
        0.62 * inch,
        0.42 * inch,
        1.0 * inch,
        2.05 * inch,
        0.72 * inch,
        1.0 * inch,
        0.55 * inch,
        width - (6.36 * inch),
    ]
    headers = [
        _qty_type_header(qty_type),
        "TYPE",
        "PO #",
        "ITEM DESCRIPTION",
        "ITEM #",
        "UPC #",
        "# SKIDS",
        "WEIGHT",
    ]

    y = top - header_h
    canv.setFillColor(HEADER_FILL)
    canv.rect(x, y, width, header_h, stroke=0, fill=1)
    canv.setFillColor(colors.black)
    cursor_x = x
    for header, col_w in zip(headers, col_widths):
        _draw_box(canv, cursor_x, y, col_w, header_h)
        _draw_string_fit(
            canv,
            cursor_x + 2,
            y + 6,
            header,
            col_w - 4,
            font_name=FONT_BOLD,
            font_size=6.7,
            align="center",
        )
        cursor_x += col_w

    item_style = _paragraph_style("ItemCell", font_size=6.5, leading=7.2)
    for row_index in range(rows_to_draw):
        y -= row_h
        line = item_lines[row_index] if row_index < len(item_lines) else None
        values = (
            [
                line.pallet_qty,
                rendered_type,
                line.po_number,
                line.item_description,
                line.item_number,
                line.upc,
                line.skids,
                line.weight_each,
            ]
            if line is not None
            else ["", "", "", "", "", "", "", ""]
        )
        cursor_x = x
        for col_index, (value, col_w) in enumerate(zip(values, col_widths)):
            _draw_box(canv, cursor_x, y, col_w, row_h)
            if col_index == 3:
                _draw_paragraph(
                    canv,
                    value,
                    cursor_x + 3,
                    y + row_h - 4,
                    col_w - 6,
                    row_h - 6,
                    style=item_style,
                )
            else:
                _draw_string_fit(
                    canv,
                    cursor_x + 3,
                    y + row_h - 12,
                    value,
                    col_w - 6,
                    font_size=6.5,
                    min_size=5.5,
                )
            cursor_x += col_w

    y -= totals_h
    cursor_x = x
    totals = [total_qty, "", "", "TOTALS", "", "", total_skids, total_weight]
    for col_index, (value, col_w) in enumerate(zip(totals, col_widths)):
        fill = col_index == 3
        if fill:
            canv.setFillColor(LIGHT_FILL)
            canv.rect(cursor_x, y, col_w, totals_h, stroke=0, fill=1)
            canv.setFillColor(colors.black)
        _draw_box(canv, cursor_x, y, col_w, totals_h)
        font = FONT_BOLD if col_index in {0, 3, 6, 7} else FONT_NAME
        align = "center" if col_index == 3 else "left"
        _draw_string_fit(
            canv,
            cursor_x + 3,
            y + 6,
            value,
            col_w - 6,
            font_name=font,
            font_size=7,
            align=align,
        )
        cursor_x += col_w

    return y


def _draw_footer(canv: canvas.Canvas, record: BolStandardRecord, table_bottom: float, *, mode: str) -> None:
    x = MARGIN
    width = PAGE_WIDTH - (2 * MARGIN)
    y = table_bottom - 0.08 * inch

    if mode == "No Recourse":
        notice_h = 0.42 * inch
        _draw_box(canv, x, y - notice_h, width, notice_h)
        notice = (
            "BROKER PAYMENT / NO RECOURSE NOTICE: Carrier agrees to seek payment solely from "
            "the broker and waives recourse against shipper, consignee, and Kendal King for "
            "freight charges unless otherwise required by law."
        )
        _draw_paragraph(
            canv,
            notice,
            x + 5,
            y - 5,
            width - 10,
            notice_h - 10,
            style=_paragraph_style("NoRecourseNotice", font_name=FONT_BOLD, font_size=6.6, leading=7.4),
        )
        y -= notice_h + 0.06 * inch

    legal_h = 0.44 * inch
    _draw_box(canv, x, y - legal_h, width, legal_h)
    legal = (
        "Received, subject to individually determined rates or contracts that have been agreed upon "
        "in writing between the carrier and shipper, if applicable, otherwise to the rates, "
        "classifications and rules that have been established by the carrier and are available "
        "to the shipper on request."
    )
    _draw_paragraph(
        canv,
        legal,
        x + 5,
        y - 5,
        width - 10,
        legal_h - 10,
        style=_paragraph_style("Legal", font_size=5.7, leading=6.5),
    )
    y -= legal_h + 0.08 * inch

    sig_h = 0.58 * inch
    half = width / 2
    _draw_box(canv, x, y - sig_h, half, sig_h)
    _draw_box(canv, x + half, y - sig_h, half, sig_h)
    canv.setFont(FONT_BOLD, 6.7)
    canv.drawString(x + 5, y - 10, "SHIPPER SIGNATURE / DATE")
    canv.drawString(x + half + 5, y - 10, "CARRIER SIGNATURE / PICKUP DATE")
    canv.line(x + 8, y - sig_h + 13, x + half - 8, y - sig_h + 13)
    canv.line(x + half + 8, y - sig_h + 13, x + width - 8, y - sig_h + 13)

    cod_y = y - sig_h - 0.06 * inch
    _draw_box(canv, x, cod_y - 0.34 * inch, width, 0.34 * inch)
    _draw_checkbox(canv, x + 6, cod_y - 16, "COD Amount: $")
    canv.line(x + 92, cod_y - 12, x + 190, cod_y - 12)
    _draw_checkbox(canv, x + 210, cod_y - 16, "Fee terms: Collect")
    _draw_checkbox(canv, x + 315, cod_y - 16, "Prepaid")
    _draw_checkbox(canv, x + 390, cod_y - 16, "Customer check acceptable")


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

    _draw_header(canv, record, mode, resolved_comment=_resolve_comment(record, batch_comment))
    _draw_party_blocks(canv, record, selected_facility)
    _draw_terms_subject_billto(canv, record)
    table_bottom = _draw_item_table(
        canv,
        record,
        mode=mode,
        bol_type=bol_type,
        qty_type=qty_type,
    )
    _draw_footer(canv, record, table_bottom, mode=mode)
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
