"""PDF generator service for creating print-ready label documents."""

from __future__ import annotations

from io import BytesIO

from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.graphics import renderPDF

from app.models.label import Label
from app.services.barcode_service import generate_code128_barcode
from app.utils.formatting import drop_leading_zeros, sanitize_text


# ── Page size: half-letter label (4.25" × 6.25") ──────────────────────────────
PAGE_WIDTH  = 4.25 * inch   # 306 pt
PAGE_HEIGHT = 6.25 * inch   # 450 pt

# ── Layout constants derived from the model PDF ────────────────────────────────
LEFT_MARGIN   = 25.2        # pts (~0.35")
FONT_SIZE     = 12          # pt – body text
FOOTER_SIZE   = 10          # pt – footer lines
LINE_GAP      = 16.1        # pt – standard single line gap
DOUBLE_GAP    = 32.2        # pt – double gap (used after ELECTRONICS and CONTENTS)


def _draw_wrapped_text(
    c: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    max_width: float,
    line_height: float,
    max_lines: int = 2,
) -> float:
    """Draw text wrapped across up to `max_lines` lines; returns y after last line."""
    words = text.split()
    lines: list[str] = []
    current = ""

    for word in words:
        test = f"{current} {word}".strip()
        if c.stringWidth(test, "Helvetica", FONT_SIZE) <= max_width:
            current = test
        else:
            if current:
                lines.append(current)
            current = word
            if len(lines) >= max_lines:
                break

    if current and len(lines) < max_lines:
        lines.append(current)

    for line in lines:
        c.drawString(x, y, line)
        y -= line_height

    return y


def _draw_label_page(c: canvas.Canvas, label: Label) -> None:
    """Draw one label page pixel-matched to the Sony EOTF model PDF."""

    # ReportLab y=0 is the bottom of the page; the model measures from the top.
    # We anchor every element using: rl_y = PAGE_HEIGHT - top_pt - font_size
    # For 12pt text the baseline sits ~4pt above the bottom of the character box.
    # The model "top" values are the top of the text box, so:
    #   baseline_rl_y = PAGE_HEIGHT - top_pt - (font_size - 4)  ≈ PAGE_HEIGHT - top_pt - 8
    # After calibration the following offsets reproduce exact positions.

    c.setFont("Helvetica", FONT_SIZE)

    # ── 1. Shipper ─────────────────────────────────────────────── top = 20.2 ──
    y = PAGE_HEIGHT - 20.2 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, f"Shipper: {sanitize_text(label.supplier)}")

    # ── 2. ATTN ────────────────────────────────────────────────── top = 36.3 ──
    y = PAGE_HEIGHT - 36.3 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, "ATTN: Dept. Mgr. Dept#: 5")

    # ── 3. ELECTRONICS DEPARTMENT ──────────────────────────────── top = 52.4 ──
    y = PAGE_HEIGHT - 52.4 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, "ELECTRONICS DEPARTMENT")

    # ── 4. STORE (double gap below ELECTRONICS) ────────────────── top = 84.6 ──
    y = PAGE_HEIGHT - 84.6 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, f"STORE #: {sanitize_text(label.store)}")

    # ── 5. CONTENTS ────────────────────────────────────────────── top = 100.6 ─
    y = PAGE_HEIGHT - 100.6 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, "CONTENTS: SIGNAGE KITS")

    # ── 6. PO # (double gap below CONTENTS) ───────────────────── top = 133.0 ─
    po_display = drop_leading_zeros(label.po)
    y = PAGE_HEIGHT - 133.0 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, f"PO #: {po_display}")

    # ── 7. PO barcode ─────────────────────────────────────────── top = 170.9 ─
    # Barcode height in the model: 194.9 - 170.9 = 24 pt → use barWidth to match
    po_barcode = generate_code128_barcode(label.po)
    po_barcode.barHeight = 24         # pt – match model barcode height
    po_barcode.barWidth  = 0.72       # pt – controls overall width; tune if needed
    barcode_y = PAGE_HEIGHT - 194.9   # bottom of barcode box in RL coords
    renderPDF.draw(po_barcode, c, LEFT_MARGIN, barcode_y)

    # ── 8. Desc label ─────────────────────────────────────────── top = 197.1 ─
    y = PAGE_HEIGHT - 197.1 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, "Desc:")

    # ── 9. Description text (wraps at ~230 pt wide) ────────────── top = 213.2 ─
    y = PAGE_HEIGHT - 213.2 - FONT_SIZE
    _draw_wrapped_text(
        c,
        sanitize_text(label.description),
        LEFT_MARGIN,
        y,
        max_width=230,       # pt – matches model line-break point
        line_height=LINE_GAP,
        max_lines=2,
    )

    # ── 10. SAP # ─────────────────────────────────────────────── top = 263.8 ─
    sap_display = drop_leading_zeros(label.sap)
    y = PAGE_HEIGHT - 263.8 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, f"SAP #: {sap_display}")

    # ── 11. SAP barcode ───────────────────────────────────────── top = 307.8 ─
    # Barcode height in the model: 335.9 - 307.8 = 28.1 pt
    sap_barcode = generate_code128_barcode(label.sap)
    sap_barcode.barHeight = 28         # pt
    sap_barcode.barWidth  = 0.72
    barcode_y = PAGE_HEIGHT - 335.9
    renderPDF.draw(sap_barcode, c, LEFT_MARGIN, barcode_y)

    # ── 12. CAT ───────────────────────────────────────────────── top = 338.2 ─
    y = PAGE_HEIGHT - 338.2 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, "CAT: ELECTRONICS DEPT.")

    # ── 13. QTY ───────────────────────────────────────────────── top = 354.3 ─
    y = PAGE_HEIGHT - 354.3 - FONT_SIZE
    c.drawString(LEFT_MARGIN, y, "QTY: 1")

    # ── 14. Footer (centered) ─────────────────────────────────── top = 386.5 ─
    c.setFont("Helvetica", FOOTER_SIZE)
    footer1 = "For questions or additional information, call"
    footer2 = "Tara Webb 501-454-6407"

    # Center both footer lines within the page width
    f1_w = c.stringWidth(footer1, "Helvetica", FOOTER_SIZE)
    f2_w = c.stringWidth(footer2, "Helvetica", FOOTER_SIZE)
    f1_x = (PAGE_WIDTH - f1_w) / 2
    f2_x = (PAGE_WIDTH - f2_w) / 2

    y = PAGE_HEIGHT - 386.5 - FOOTER_SIZE
    c.drawString(f1_x, y, footer1)

    y = PAGE_HEIGHT - 403.1 - FOOTER_SIZE
    c.drawString(f2_x, y, footer2)


def generate_label_pdf(labels: list[Label]) -> bytes:
    """Generate a label PDF with one label per page (4.25" × 6.25" pages)."""

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