"""Fixed 15-row pages for Word and PDF BOLs."""

from copy import deepcopy
from math import ceil

from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.table import Table, _Row
from reportlab.pdfbase.pdfmetrics import stringWidth


MAX_ITEM_ROWS = 15
WORD_ITEM_BODY_HEIGHT_PT = {"Standard": 132.0, "No Recourse": 120.0, "Multistop": 170.0}


def validate_item_row_count(item_row_count: int) -> None:
    if item_row_count != MAX_ITEM_ROWS:
        raise ValueError(f"BOL pages have a maximum of {MAX_ITEM_ROWS} item rows.")


def apply_docx_item_row_layout(doc, item_row_count: int, template_mode: str) -> None:
    """Repeat the full BOL form for each page of items, without clipping text."""
    validate_item_row_count(item_row_count)
    for table in doc.tables:
        rows = list(table.rows)
        header = next((i for i, row in enumerate(rows)
                       if "ITEM DESCRIPTION" in " ".join(c.text for c in row.cells).upper()
                       and "PO #" in " ".join(c.text for c in row.cells).upper()), None)
        if header is None:
            continue
        totals = next((i for i in range(header + 1, len(rows))
                       if "TOTALS" in " ".join(c.text for c in rows[i].cells).upper()), None)
        if totals is None or totals == header + 1:
            continue
        body_rows = rows[header + 1:totals]
        data_rows = [deepcopy(r._tr) for r in body_rows if any(c.text.strip() for c in r.cells)]
        blank_template = deepcopy(body_rows[0]._tr)
        for cell in _Row(blank_template, table).cells:
            cell.text = ""
        body = doc.element.body
        original = [deepcopy(element) for element in body if element.tag != qn("w:sectPr")]
        table_index = list(body).index(table._tbl)
        pages = []
        for start in range(0, max(1, len(data_rows)), item_row_count):
            elements = deepcopy(original)
            page_table = Table(elements[table_index], doc._body)
            page_rows = list(page_table.rows)
            anchor = page_rows[totals]._tr
            for row in page_rows[header + 1:totals]:
                page_table._tbl.remove(row._tr)
            chunk = data_rows[start:start + item_row_count]
            chunk = chunk + [blank_template] * (item_row_count - len(chunk))
            for row_xml in chunk:
                anchor.addprevious(deepcopy(row_xml))
            _format_docx_rows(list(page_table.rows)[header + 1:header + 1 + item_row_count], template_mode)
            # Remove the template's leading blank totals lines; they otherwise
            # push the customer notice onto a separate printed page.
            _compact_footer_rows(list(page_table.rows)[header + 1 + item_row_count:])
            if template_mode == "No Recourse":
                _compact_no_recourse_notice(elements[table_index + 1:])
            pages.append(elements)
        for element in list(body):
            if element.tag != qn("w:sectPr"):
                body.remove(element)
        for index, elements in enumerate(pages):
            if index:
                separator = OxmlElement("w:p")
                properties = OxmlElement("w:pPr")
                properties.append(OxmlElement("w:pageBreakBefore"))
                spacing = OxmlElement("w:spacing")
                for attr, value in (("before", "0"), ("after", "0"), ("line", "20"), ("lineRule", "exact")):
                    spacing.set(qn(f"w:{attr}"), value)
                properties.append(spacing)
                separator.append(properties)
                elements.insert(0, separator)
            for element in elements:
                body.insert(len(body) - (1 if body.sectPr is not None else 0), element)
        # Repeated drawing anchors must have unique IDs within the package.
        drawing_id = 1
        for element in body.iter():
            if element.tag == qn("wp:docPr"):
                element.set("id", str(drawing_id))
                drawing_id += 1
        return


def _compact_no_recourse_notice(elements):
    """Keep the existing signature drawings and fit the complete broker notice."""
    from docx.text.paragraph import Paragraph

    for element in elements:
        if element.tag == qn("w:tbl"):
            rows = list(element.findall(qn("w:tr")))
            while len(rows) > 1 and not any((node.text or "").strip() for node in rows[-1].iter(qn("w:t"))):
                if any(node.tag in {qn("w:drawing"), "{urn:schemas-microsoft-com:vml}shape"}
                       for node in rows[-1].iter()):
                    break
                element.remove(rows.pop())
        for node in element.iter(qn("w:p")):
            if any(child.tag in {qn("w:drawing"), "{urn:schemas-microsoft-com:vml}shape"}
                   for child in node.iter()):
                continue
            paragraph = Paragraph(node, None)
            text = paragraph.text.strip()
            if "BROKER PAYMENT" in text or text.startswith("Broker of Record:"):
                fmt = paragraph.paragraph_format
                fmt.space_before = fmt.space_after = Pt(0)
                fmt.line_spacing = Pt(12 if "BROKER PAYMENT" in text else 10)
                for run in paragraph.runs:
                    run.font.size = Pt(9 if "BROKER PAYMENT" in text else 8)
    # Word requires a final paragraph after a table. Keep it, but prevent its
    # inherited line height from creating an otherwise empty extra page.
    for element in reversed(elements):
        if element.tag != qn("w:p") or any((node.text or "").strip() for node in element.iter(qn("w:t"))):
            break
        if any(node.tag in {qn("w:drawing"), "{urn:schemas-microsoft-com:vml}shape"}
               for node in element.iter()):
            break
        fmt = Paragraph(element, None).paragraph_format
        fmt.space_before = fmt.space_after = Pt(0)
        fmt.line_spacing = Pt(1)


def _compact_footer_rows(rows):
    for index, row in enumerate(rows[:2]):
        row.height = Pt(20 if index == 0 else 14)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        for cell in {id(c._tc): c for c in row.cells}.values():
            alignment = cell.paragraphs[0].alignment
            cell.text = " ".join(cell.text.split())
            cell.paragraphs[0].alignment = alignment
            for paragraph in cell.paragraphs:
                fmt = paragraph.paragraph_format
                fmt.space_before = fmt.space_after = Pt(0)
                fmt.line_spacing = Pt(9)
                for run in paragraph.runs:
                    run.font.size = Pt(8 if index == 0 else 7)


def _format_docx_rows(data_rows, template_mode):
    cells_by_row = []
    for row in data_rows:
        cells = list({id(c._tc): c for c in row.cells}.values())
        cells_by_row.append(cells)
        properties = row._tr.get_or_add_trPr()
        if properties.find(qn("w:cantSplit")) is None:
            properties.append(OxmlElement("w:cantSplit"))
        for cell in cells:
            alignment = cell.paragraphs[0].alignment
            # The templates use separate paragraphs for description and item
            # codes. Flow them together so item rows do not inherit extra lines.
            cell.text = " ".join(cell.text.split())
            cell.paragraphs[0].alignment = alignment
            tc_pr = cell._tc.get_or_add_tcPr()
            for no_wrap in list(tc_pr.findall(qn("w:noWrap"))):
                tc_pr.remove(no_wrap)
            margins = tc_pr.find(qn("w:tcMar"))
            if margins is None:
                margins = OxmlElement("w:tcMar")
                tc_pr.append(margins)
            for edge, value in (("top", 0), ("bottom", 0), ("left", 60), ("right", 60)):
                node = margins.find(qn(f"w:{edge}"))
                if node is None:
                    node = OxmlElement(f"w:{edge}")
                    margins.append(node)
                node.set(qn("w:w"), str(value))
                node.set(qn("w:type"), "dxa")
            for paragraph in cell.paragraphs:
                # Match the compact PDF text flow and remove template indentation.
                paragraph.text = " ".join(paragraph.text.split())
                fmt = paragraph.paragraph_format
                fmt.space_before = fmt.space_after = Pt(0)
                fmt.left_indent = fmt.right_indent = fmt.first_line_indent = Pt(0)
                fmt.keep_with_next = False
                fmt.widow_control = False
                p_pr = paragraph._p.get_or_add_pPr()
                snap = OxmlElement("w:snapToGrid")
                snap.set(qn("w:val"), "0")
                p_pr.append(snap)

    def row_heights(size):
        heights = []
        for cells in cells_by_row:
            line_count = max((sum(max(1, ceil(stringWidth(line, "Helvetica", size) /
                              max(1, cell.width.pt - 6))) for line in cell.text.splitlines())
                              for cell in cells), default=1)
            heights.append(max(1, line_count) * size * 1.15 + 2)
        return heights

    budget = WORD_ITEM_BODY_HEIGHT_PT[template_mode]
    size = 8.5
    while sum(row_heights(size)) > budget and size > 1:
        size = round(size - 0.1, 1)
    heights = row_heights(size)
    spare = max(0, budget - sum(heights)) / len(data_rows)
    for row, cells, height in zip(data_rows, cells_by_row, heights):
        row.height = Pt(height + spare)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        for cell in cells:
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.line_spacing = Pt(size * 1.15)
                for run in paragraph.runs:
                    run.font.size = Pt(size)
