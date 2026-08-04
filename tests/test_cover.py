"""Tests for cover sheet generation."""

from __future__ import annotations

from pypdf import PdfReader
from reportlab.lib.pagesizes import letter

from coversheets.cover import (
    DEFAULT_VERTICAL_MARGIN,
    cover_label_from_filename,
    cover_text_bottom_y,
    create_cover_sheet,
    normalize_vertical_position,
    preview_rely_for_position,
)


def test_cover_label_from_filename_strips_extension() -> None:
    assert cover_label_from_filename("Exhibit A.pdf") == "Exhibit A"
    assert cover_label_from_filename("nested/path/Foo & Bar.pdf") == "Foo & Bar"


def test_normalize_vertical_position() -> None:
    assert normalize_vertical_position("center") == "center"
    assert normalize_vertical_position("top_third") == "top_third"
    assert normalize_vertical_position("TOP-THIRD") == "top_third"
    assert normalize_vertical_position("Top Third") == "top_third"
    assert normalize_vertical_position("nope") == "center"
    assert normalize_vertical_position(None) == "center"


def test_cover_text_bottom_y_center_and_top_third() -> None:
    page_h = float(letter[1])
    text_h = 42.0

    center_y = cover_text_bottom_y(page_h, text_h, "center")
    top_y = cover_text_bottom_y(page_h, text_h, "top_third")

    # Midpoint of the text block should sit at the intended target.
    assert abs((center_y + text_h / 2) - (page_h * 0.5)) < 1e-6
    assert abs((top_y + text_h / 2) - (page_h * 5.0 / 6.0)) < 1e-6
    assert top_y > center_y


def test_cover_text_bottom_y_clamps_to_margins() -> None:
    page_h = float(letter[1])
    margin = DEFAULT_VERTICAL_MARGIN
    # Tall enough that unclamped top_third would exceed the top margin.
    text_h = 200.0
    unclamped_target = page_h * (5.0 / 6.0) - text_h / 2
    assert unclamped_target + text_h > page_h - margin

    y = cover_text_bottom_y(page_h, text_h, "top_third", margin=margin)
    assert y + text_h <= page_h - margin + 1e-6
    assert y >= margin - 1e-6

    # Extremely tall: pin to bottom margin (top may still overflow).
    huge = page_h - margin
    y_huge = cover_text_bottom_y(page_h, huge, "top_third", margin=margin)
    assert y_huge == margin


def test_preview_rely_for_position() -> None:
    assert preview_rely_for_position("center") == 0.5
    assert abs(preview_rely_for_position("top_third") - (1.0 / 6.0)) < 1e-9


def test_create_cover_sheet_is_single_page_pdf() -> None:
    buf = create_cover_sheet("Exhibit A")
    reader = PdfReader(buf)
    assert len(reader.pages) == 1


def test_create_cover_sheet_top_third_is_single_page() -> None:
    buf = create_cover_sheet("Exhibit A", vertical_position="top_third")
    reader = PdfReader(buf)
    assert len(reader.pages) == 1


def test_create_cover_sheet_escapes_markup_specials() -> None:
    # Must not raise when label contains ReportLab Paragraph markup chars.
    buf = create_cover_sheet("A <B> & C")
    reader = PdfReader(buf)
    assert len(reader.pages) == 1
