"""Complete delivery schedules for loads beyond the master form's three slots."""

from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


MASTER_DELIVERY_SLOTS = 3
CONTINUATION_STOPS_PER_PAGE = 8
HEADERS = ("Stop", "BOL / PO", "Delivery", "Load detail")
WIDTH_RATIOS = (0.075, 0.225, 0.525, 0.175)


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


def _chunks(record):
    rows = list(_stop_rows(record))
    for start in range(0, len(rows), CONTINUATION_STOPS_PER_PAGE):
        yield rows[start:start + CONTINUATION_STOPS_PER_PAGE]


def _identity(record):
    return f"Master BOL: {record.bol_number}    KK Load: {record.kk_load_number}"


def append_delivery_continuation_docx(doc, record):
    """Keep the existing form intact and append readable delivery address pages."""
    section = doc.sections[-1]
    width = section.page_width - section.left_margin - section.right_margin
    for rows in _chunks(record):
        title = doc.add_paragraph()
        title.paragraph_format.page_break_before = True
        title.paragraph_format.space_after = Pt(6)
        run = title.add_run("Delivery continuation")
        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(12)
        identity = doc.add_paragraph(_identity(record))
        identity.paragraph_format.space_after = Pt(10)
        for run in identity.runs:
            run.font.size = Pt(9)

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
                if values is HEADERS:
                    shading = OxmlElement("w:shd")
                    shading.set(qn("w:fill"), "EEEEEE")
                    properties.append(shading)
                for paragraph in cell.paragraphs:
                    fmt = paragraph.paragraph_format
                    fmt.space_before = fmt.space_after = Pt(5)
                    fmt.line_spacing = Pt(11)
                    fmt.keep_with_next = False
                    for run in paragraph.runs:
                        run.font.name = "Arial"
                        run.font.size = Pt(8.5)
                        run.bold = values is HEADERS


def append_delivery_continuation_pdf(destination: Path, record):
    """Append actual PDF pages, avoiding hidden form text beneath a white cover."""
    chunks = list(_chunks(record))
    if not chunks:
        return
    buffer = BytesIO()
    style = ParagraphStyle("Delivery", fontName="Helvetica", fontSize=8.5, leading=11,
                           wordWrap="CJK")
    title_style = ParagraphStyle("Title", parent=style, fontName="Helvetica-Bold",
                                fontSize=12, leading=15, spaceAfter=6)
    story = []
    for index, rows in enumerate(chunks):
        if index:
            story.append(PageBreak())
        story.extend([Paragraph("Delivery continuation", title_style),
                      Paragraph(escape(_identity(record)), style), Spacer(1, 10)])
        data = [[Paragraph(escape(value).replace("\n", "<br/>"), style)
                 for value in values] for values in (HEADERS, *rows)]
        table = Table(data, colWidths=[540 * ratio for ratio in WIDTH_RATIOS], repeatRows=1)
        table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EEEEEE")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(table)
    SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36,
                      topMargin=36, bottomMargin=36).build(story)
    # Load the original into memory before replacing the destination on Windows.
    writer = PdfWriter()
    writer.append(PdfReader(BytesIO(destination.read_bytes())))
    buffer.seek(0)
    writer.append(PdfReader(buffer))
    with destination.open("wb") as output:
        writer.write(output)

