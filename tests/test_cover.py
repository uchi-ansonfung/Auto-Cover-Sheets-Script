"""Tests for cover sheet generation."""

from __future__ import annotations

from pypdf import PdfReader

from coversheets.cover import cover_label_from_filename, create_cover_sheet


def test_cover_label_from_filename_strips_extension() -> None:
    assert cover_label_from_filename("Exhibit A.pdf") == "Exhibit A"
    assert cover_label_from_filename("nested/path/Foo & Bar.pdf") == "Foo & Bar"


def test_create_cover_sheet_is_single_page_pdf() -> None:
    buf = create_cover_sheet("Exhibit A")
    reader = PdfReader(buf)
    assert len(reader.pages) == 1


def test_create_cover_sheet_escapes_markup_specials() -> None:
    # Must not raise when label contains ReportLab Paragraph markup chars.
    buf = create_cover_sheet("A <B> & C")
    reader = PdfReader(buf)
    assert len(reader.pages) == 1
