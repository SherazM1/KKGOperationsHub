"""Complete delivery schedules for loads beyond the master form's three slots."""

from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle


MASTER_DELIVERY_SLOTS = 3
HEADERS = ("Stop", "BOL / PO", "Delivery", "Load detail")
WIDTH_RATIOS = (0.075, 0.225, 0.525, 0.175)
TABLE_WIDTH = 540
TABLE_HEIGHT = 650
FONT_SIZES = (8.5, 8.0, 7.5, 7.0)


def _stop_rows(record):
    if len(record.stops) <= MASTER_DELIVERY_SLOTS:
        return
    # Include every stop: some Word master templates have only one visible
    # consignee block. The schedule must stand on its own for the entire load.
    for stop in record.stops:
        yield (
            str(stop.stop_number),
            f"BOL: {stop.bol_number or record.bol_number}\nPO: {stop.target_po_number}",
            "\n".join(part for part in (
                stop.delivery_dc, stop.delivery_address, stop.delivery_city_state_zip,
            ) if part),
            f"Cases: {stop.cases}\nPallets: {stop.total_pallets}\nWeight: {stop.weight}",
        )


def _delivery_table(rows, width, font_size):
    """Measure the same compact content used by both Word and PDF tables."""
    leading = font_size * 1.2
    padding = 2 if font_size >= 8 else 1
    style = ParagraphStyle("Delivery", fontName="Helvetica", fontSize=font_size,
                           leading=leading, wordWrap="CJK")
    header_style = ParagraphStyle("DeliveryHeader", parent=style, fontName="Helvetica-Bold")
    data = [[Paragraph(escape(value).replace("\n", "<br/>"),
                       header_style if index == 0 else style)
             for value in values] for index, values in enumerate((HEADERS, *rows))]
    table = Table(data, colWidths=[width * ratio for ratio in WIDTH_RATIOS],
                  repeatRows=1, hAlign="LEFT")
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EEEEEE")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), padding),
        ("BOTTOMPADDING", (0, 0), (-1, -1), padding),
    ]))
    return table, leading, padding


def _fit_delivery_table(rows, width=TABLE_WIDTH, height=TABLE_HEIGHT):
    # Prefer a single page, but retain readable type and every stop if a very
    # large schedule really needs to flow onto additional pages.
    for font_size in FONT_SIZES:
        table, leading, padding = _delivery_table(rows, width, font_size)
        if table.wrap(width, height)[1] <= height:
            break
    return table, font_size, leading, padding


def _identity(record):
    return f"Master BOL: {record.bol_number}    KK Load: {record.kk_load_number}"


def append_delivery_continuation_docx(doc, record):
    """Keep the existing form intact and append readable delivery address pages."""
    rows = list(_stop_rows(record))
    if not rows:
        return
    section = doc.sections[-1]
    width_pt = min(TABLE_WIDTH, (section.page_width - section.left_margin - section.right_margin) / Pt(1))
    height_pt = min(TABLE_HEIGHT, (section.page_height - section.top_margin - section.bottom_margin) / Pt(1) - 70)
    _, font_size, leading, padding = _fit_delivery_table(rows, width_pt, height_pt)
    width = Pt(width_pt)
    # One flowing table replaces the old forced page break after every eight stops.
    title = doc.add_paragraph()
    title.paragraph_format.page_break_before = True
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after = Pt(6)
    title.paragraph_format.line_spacing = Pt(15)
    title.paragraph_format.keep_with_next = True
    run = title.add_run("Stop Manifest")
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0, 0, 0)
    identity = doc.add_paragraph(_identity(record))
    identity.paragraph_format.space_before = Pt(0)
    identity.paragraph_format.space_after = Pt(10)
    identity.paragraph_format.line_spacing = Pt(11)
    identity.paragraph_format.keep_with_next = True
    for run in identity.runs:
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0, 0, 0)

    table = doc.add_table(rows=1, cols=len(HEADERS))
    table.autofit = False
    for column, ratio in zip(table.columns, WIDTH_RATIOS):
        column.width = int(width * ratio)
    table.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
    for values in (HEADERS, *rows):
        row = table.rows[0] if values is HEADERS else table.add_row()
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        for cell, value, ratio in zip(row.cells, values, WIDTH_RATIOS):
            cell.width = int(width * ratio)
            cell.text = value
            properties = cell._tc.get_or_add_tcPr()
            borders = OxmlElement("w:tcBorders")
            for edge in ("top", "bottom", "left", "right"):
                border = OxmlElement(f"w:{edge}")
                for attr, setting in (("val", "single"), ("sz", "4"), ("color", "D9D9D9")):
                    border.set(qn(f"w:{attr}"), setting)
                borders.append(border)
            properties.append(borders)
            margins = OxmlElement("w:tcMar")
            for edge, points in (("top", padding), ("bottom", padding), ("left", 6), ("right", 6)):
                margin = OxmlElement(f"w:{edge}")
                margin.set(qn("w:w"), str(int(points * 20)))
                margin.set(qn("w:type"), "dxa")
                margins.append(margin)
            properties.append(margins)
            if values is HEADERS:
                shading = OxmlElement("w:shd")
                shading.set(qn("w:fill"), "EEEEEE")
                properties.append(shading)
            for paragraph in cell.paragraphs:
                fmt = paragraph.paragraph_format
                fmt.space_before = fmt.space_after = Pt(0)
                fmt.line_spacing = Pt(leading)
                fmt.keep_with_next = values is HEADERS
                for run in paragraph.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(font_size)
                    run.font.color.rgb = RGBColor(0, 0, 0)
                    run.bold = values is HEADERS


def append_delivery_continuation_pdf(destination: Path, record):
    """Append actual PDF pages, avoiding hidden form text beneath a white cover."""
    rows = list(_stop_rows(record))
    if not rows:
        return
    buffer = BytesIO()
    table, _, _, _ = _fit_delivery_table(rows)
    style = ParagraphStyle("Identity", fontName="Helvetica", fontSize=9, leading=11)
    title_style = ParagraphStyle("Title", parent=style, fontName="Helvetica-Bold",
                                fontSize=12, leading=15, spaceAfter=6)

    def draw_heading(canv, _document):
        # Keep the load identity on every page if an unusually large schedule
        # cannot fit at the minimum readable font size.
        title = Paragraph("Delivery continuation", title_style)
        _, title_height = title.wrap(TABLE_WIDTH, 40)
        title.drawOn(canv, 36, letter[1] - 36 - title_height)
        identity = Paragraph(escape(_identity(record)), style)
        _, identity_height = identity.wrap(TABLE_WIDTH, 30)
        identity.drawOn(canv, 36, letter[1] - 42 - title_height - identity_height)

    story = [table]
    SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36,
                      topMargin=84, bottomMargin=36).build(
                          story, onFirstPage=draw_heading, onLaterPages=draw_heading,
                      )
    # Load the original into memory before replacing the destination on Windows.
    writer = PdfWriter()
    writer.append(PdfReader(BytesIO(destination.read_bytes())))
    buffer.seek(0)
    writer.append(PdfReader(buffer))
    with destination.open("wb") as output:
        writer.write(output)

