"""Tests for PDF merge / metadata handling."""

from __future__ import annotations

from pathlib import Path

import pytest
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

from coversheets.cover import create_cover_sheet
from coversheets.merge import PARTIAL_SUFFIX, add_cover_to_pdf, partial_path_for
from coversheets.options import ProcessOptions


def test_add_cover_prepends_page(sample_pdf: Path, tmp_path: Path) -> None:
    cover = create_cover_sheet("Exhibit A")
    out = tmp_path / "+Exhibit A.pdf"
    add_cover_to_pdf(cover, sample_pdf, out, compress=False)

    result = PdfReader(str(out))
    assert len(result.pages) == 2  # cover + original body


def test_add_cover_with_compress_succeeds(sample_pdf: Path, tmp_path: Path) -> None:
    """compress_content_streams must run on writer pages, not reader pages."""
    cover = create_cover_sheet("Exhibit A")
    out = tmp_path / "+Exhibit A.pdf"
    add_cover_to_pdf(
        cover,
        sample_pdf,
        out,
        options=ProcessOptions(compress=True, strip_metadata=True, optimize=True),
    )

    result = PdfReader(str(out))
    assert len(result.pages) == 2
    assert out.stat().st_size > 0


def test_true_metadata_strip_removes_info_fields(tmp_path: Path) -> None:
    original = tmp_path / "meta.pdf"
    c = canvas.Canvas(str(original), pagesize=letter)
    c.drawString(72, 720, "body")
    c.save()

    reader = PdfReader(str(original))
    writer = PdfWriter()
    writer.add_page(reader.pages[0])
    writer.add_metadata(
        {
            "/Title": "Secret Title",
            "/Author": "Secret Author",
            "/Subject": "Secret Subject",
        }
    )
    with original.open("wb") as fh:
        writer.write(fh)

    cover = create_cover_sheet("meta")
    out = tmp_path / "+meta.pdf"
    add_cover_to_pdf(
        cover,
        original,
        out,
        options=ProcessOptions(compress=False, strip_metadata=True),
    )

    meta = PdfReader(str(out)).metadata
    # True strip sets metadata to None (no Info dict / no Producer).
    assert meta is None or (
        meta.get("/Title") in (None, "")
        and meta.get("/Author") in (None, "")
        and meta.get("/Subject") in (None, "")
    )


def test_atomic_write_leaves_no_partial_on_success(
    sample_pdf: Path, tmp_path: Path
) -> None:
    cover = create_cover_sheet("Exhibit A")
    out = tmp_path / "+Exhibit A.pdf"
    add_cover_to_pdf(
        cover,
        sample_pdf,
        out,
        options=ProcessOptions(compress=False, strip_metadata=True, optimize=False),
    )
    assert out.is_file()
    assert not partial_path_for(out).exists()
    assert not any(tmp_path.glob(f"*{PARTIAL_SUFFIX}"))


def test_atomic_write_cleans_partial_on_failure(
    sample_pdf: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    cover = create_cover_sheet("Exhibit A")
    out = tmp_path / "+Exhibit A.pdf"

    real_replace = Path.replace

    def boom(self: Path, target: Path) -> Path:  # type: ignore[override]
        if str(self).endswith(PARTIAL_SUFFIX):
            raise OSError("simulated replace failure")
        return real_replace(self, target)

    monkeypatch.setattr(Path, "replace", boom)
    with pytest.raises(OSError, match="simulated"):
        add_cover_to_pdf(
            cover,
            sample_pdf,
            out,
            options=ProcessOptions(compress=False, optimize=False),
        )
    assert not out.exists()
    assert not partial_path_for(out).exists()


def test_without_strip_still_avoids_source_title(tmp_path: Path) -> None:
    original = tmp_path / "meta.pdf"
    c = canvas.Canvas(str(original), pagesize=letter)
    c.drawString(72, 720, "body")
    c.save()

    reader = PdfReader(str(original))
    writer = PdfWriter()
    writer.add_page(reader.pages[0])
    writer.add_metadata({"/Title": "Secret Title"})
    with original.open("wb") as fh:
        writer.write(fh)

    cover = create_cover_sheet("meta")
    out = tmp_path / "+meta.pdf"
    add_cover_to_pdf(
        cover,
        original,
        out,
        options=ProcessOptions(compress=False, strip_metadata=False),
    )

    meta = PdfReader(str(out)).metadata
    assert meta is None or meta.get("/Title") in (None, "")
