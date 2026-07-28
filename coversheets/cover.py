"""Generate letter-sized exhibit cover sheet PDFs."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph

# ~1-inch margins on US Letter (612 x 792 pt)
DEFAULT_AVAIL_WIDTH = 468
DEFAULT_FONT_SIZE = 36
DEFAULT_LEADING = 42


def cover_label_from_filename(filename: str | Path) -> str:
    """Return default display text for a cover sheet (filename stem)."""
    return Path(filename).stem


def create_cover_sheet(
    label: str,
    *,
    font_size: int = DEFAULT_FONT_SIZE,
    avail_width: float = DEFAULT_AVAIL_WIDTH,
) -> BytesIO:
    """
    Create a letter-sized cover sheet PDF in memory.

    ``label`` is centered in Times Bold and wrapped if it exceeds
    ``avail_width``. Special characters are escaped for ReportLab Paragraph.
    """
    text = label.strip() if label.strip() else " "  # empty labels still get a page
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    page_width, page_height = letter

    base = getSampleStyleSheet()["Normal"]
    style = ParagraphStyle(
        "CoverTitle",
        parent=base,
        fontName="Times-Bold",
        fontSize=font_size,
        leading=max(
            font_size + 6,
            DEFAULT_LEADING if font_size == DEFAULT_FONT_SIZE else font_size + 6,
        ),
        alignment=TA_CENTER,
    )

    paragraph = Paragraph(escape(text), style)
    paragraph.wrap(avail_width, 1000)

    x = (page_width - avail_width) / 2
    y = (page_height / 2) + (paragraph.height / 2)
    paragraph.drawOn(c, x, y)

    c.save()
    buffer.seek(0)
    return buffer
