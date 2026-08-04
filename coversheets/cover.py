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
# Keep multi-line titles inside the page (~0.75")
DEFAULT_VERTICAL_MARGIN = 54

VERTICAL_POSITIONS = frozenset({"center", "top_third"})
DEFAULT_VERTICAL_POSITION = "center"


def normalize_vertical_position(value: str | None) -> str:
    """Return a valid vertical position; unknown values become center."""
    text = (value or "").strip().lower().replace("-", "_").replace(" ", "_")
    if text in VERTICAL_POSITIONS:
        return text
    return DEFAULT_VERTICAL_POSITION


def cover_text_bottom_y(
    page_height: float,
    text_height: float,
    vertical_position: str = DEFAULT_VERTICAL_POSITION,
    *,
    margin: float = DEFAULT_VERTICAL_MARGIN,
) -> float:
    """
    Bottom Y (ReportLab coords) so the text block's midpoint sits at the
    chosen vertical position, clamped to stay within top/bottom margins.
    """
    position = normalize_vertical_position(vertical_position)
    if position == "top_third":
        # Midpoint of the upper third of the page (band from 2/3 → top).
        target_center = page_height * (5.0 / 6.0)
    else:
        target_center = page_height * 0.5

    bottom = target_center - (text_height / 2.0)
    max_bottom = max(margin, page_height - margin - text_height)
    min_bottom = margin
    if max_bottom < min_bottom:
        # Extremely tall text: pin to bottom margin.
        return min_bottom
    return min(max(bottom, min_bottom), max_bottom)


def preview_rely_for_position(vertical_position: str) -> float:
    """
    Tk ``place`` rely for title with anchor=center on the cover preview.

    Tk y grows downward (unlike ReportLab). Mid of the top third is at
    1/6 from the top of the page.
    """
    position = normalize_vertical_position(vertical_position)
    if position == "top_third":
        return 1.0 / 6.0
    return 0.5


def cover_label_from_filename(filename: str | Path) -> str:
    """Return default display text for a cover sheet (filename stem)."""
    return Path(filename).stem


def create_cover_sheet(
    label: str,
    *,
    font_size: int = DEFAULT_FONT_SIZE,
    avail_width: float = DEFAULT_AVAIL_WIDTH,
    vertical_position: str = DEFAULT_VERTICAL_POSITION,
) -> BytesIO:
    """
    Create a letter-sized cover sheet PDF in memory.

    ``label`` is centered horizontally in Times Bold and wrapped if it
    exceeds ``avail_width``. Vertical placement is controlled by
    ``vertical_position`` (``center`` or ``top_third``). Special characters
    are escaped for ReportLab Paragraph.
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
    y = cover_text_bottom_y(
        page_height,
        paragraph.height,
        vertical_position,
    )
    paragraph.drawOn(c, x, y)

    c.save()
    buffer.seek(0)
    return buffer
