"""PDF generator service for creating print-ready label documents."""

from __future__ import annotations

from io import BytesIO

from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.graphics import renderPDF

from app.models.label import Label
from app.services.barcode_service import generate_code128_barcode
from app.utils.formatting import drop_leading_zeros, sanitize_text


PAGE_WIDTH = 4.25 * inch
PAGE_HEIGHT = 6.25 * inch

LEFT_MARGIN = 25.2
FONT_SIZE = 12
FOOTER_SIZE = 10
LINE_GAP = 16.1


def _draw_label_page(c: canvas.Canvas, label: Label) -> None:
    c.setFont("Helvetica", FONT_SIZE)

    # 1️⃣ Shipper
    y = PAGE_HEIGHT - 20.2 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, f"Shipper: {sanitize_text(label.supplier)}")

    # 2️⃣ ATTN
    y = PAGE_HEIGHT - 36.3 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, "ATTN: Dept. Mgr. Dept#: 5")

    # 3️⃣ ELECTRONICS
    y = PAGE_HEIGHT - 52.4 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, "ELECTRONICS DEPARTMENT")

    # 4️⃣ STORE
    y = PAGE_HEIGHT - 84.6 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, f"STORE #: {sanitize_text(label.store)}")

    # 5️⃣ CONTENTS
    y = PAGE_HEIGHT - 100.6 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, "CONTENTS: SIGNAGE KITS")

    # 6️⃣ PO
    po_display = drop_leading_zeros(label.po)
    y = PAGE_HEIGHT - 133.0 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, f"PO #: {po_display}")

    # 7️⃣ PO Barcode (24pt height)
    po_barcode = generate_code128_barcode(
        label.po,
        bar_height=32,
        bar_width=0.85,
    )
    barcode_y = PAGE_HEIGHT - 194.9
    renderPDF.draw(po_barcode, c, LEFT_MARGIN, barcode_y)

    # 8️⃣ Desc
    y = PAGE_HEIGHT - 197.1 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, "Desc:")

    y = PAGE_HEIGHT - 213.2 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, sanitize_text(label.description))

    # 9️⃣ SAP
    sap_display = drop_leading_zeros(label.sap)
    y = PAGE_HEIGHT - 263.8 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, f"SAP #: {sap_display}")

    # 🔟 SAP Barcode (28pt height)
    sap_barcode = generate_code128_barcode(
        label.sap,
        bar_height=28,
        bar_width=0.72,
    )
    barcode_y = PAGE_HEIGHT - 335.9
    renderPDF.draw(sap_barcode, c, LEFT_MARGIN, barcode_y)

    # 1️⃣1️⃣ CAT
    y = PAGE_HEIGHT - 338.2 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, "CAT: ELECTRONICS DEPT.")

    # 1️⃣2️⃣ QTY
    y = PAGE_HEIGHT - 354.3 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, "QTY: 1")

    # 1️⃣3️⃣ Footer (centered)
    c.setFont("Helvetica", FOOTER_SIZE)

    footer1 = "For questions or additional information, call"
    footer2 = "Tara Webb 501-454-6407"

    f1_w = c.stringWidth(footer1, "Helvetica", FOOTER_SIZE)
    f2_w = c.stringWidth(footer2, "Helvetica", FOOTER_SIZE)

    f1_x = (PAGE_WIDTH - f1_w) / 2
    f2_x = (PAGE_WIDTH - f2_w) / 2

    y = PAGE_HEIGHT - 386.5 - FOOTER_SIZE
    c.drawString(f1_x, y, footer1)

    y = PAGE_HEIGHT - 403.1 - FOOTER_SIZE
    c.drawString(f2_x, y, footer2)


def generate_label_pdf(labels: list[Label]) -> bytes:
    if not labels:
        raise ValueError("No labels provided for PDF generation.")

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))

    for label in labels:
        _draw_label_page(c, label)
        c.showPage()

    c.save()
    buffer.seek(0)
    return buffer.read()