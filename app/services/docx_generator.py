"""DOCX generator service scaffold for label documents."""

from __future__ import annotations

from app.models.label import Label


def generate_label_docx(labels: list[Label]) -> bytes:
    """Generate a print-ready DOCX for labels.

    The active app currently generates label PDFs directly. This scaffold keeps
    the legacy service import available until DOCX label output is implemented.
    """
    raise NotImplementedError("DOCX label generation is not implemented.")
