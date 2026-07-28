"""Tests for PDF post-processing helpers."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

from coversheets.pdf_ops import (
    DependencyError,
    linearize_pdf,
    run_ocr,
    strip_metadata_file,
    strip_metadata_writer,
)


def _one_page_pdf(path: Path, *, title: str | None = None) -> Path:
    c = canvas.Canvas(str(path), pagesize=letter)
    c.drawString(72, 720, "body")
    c.save()
    if title:
        reader = PdfReader(str(path))
        writer = PdfWriter()
        writer.add_page(reader.pages[0])
        writer.add_metadata({"/Title": title, "/Author": "Tester"})
        with path.open("wb") as fh:
            writer.write(fh)
    return path


def test_strip_metadata_writer_clears_info(tmp_path: Path) -> None:
    src = _one_page_pdf(tmp_path / "a.pdf", title="KeepMe")
    reader = PdfReader(str(src))
    writer = PdfWriter()
    writer.add_page(reader.pages[0])
    writer.add_metadata({"/Title": "KeepMe"})
    strip_metadata_writer(writer)
    out = tmp_path / "out.pdf"
    with out.open("wb") as fh:
        writer.write(fh)
    assert PdfReader(str(out)).metadata is None


def test_strip_metadata_file(tmp_path: Path) -> None:
    path = _one_page_pdf(tmp_path / "doc.pdf", title="Secret")
    strip_metadata_file(path)
    meta = PdfReader(str(path)).metadata
    assert meta is None or meta.get("/Title") in (None, "")


def test_run_ocr_missing_dependency() -> None:
    with patch("coversheets.pdf_ops.ocr_available", return_value=False):
        with pytest.raises(DependencyError, match="ocrmypdf"):
            run_ocr("/tmp/x.pdf")


def test_linearize_missing_dependency() -> None:
    with patch("coversheets.pdf_ops.linearize_available", return_value=False):
        with pytest.raises(DependencyError, match="pikepdf|qpdf"):
            linearize_pdf("/tmp/x.pdf")


def test_run_ocr_calls_ocrmypdf(tmp_path: Path) -> None:
    path = _one_page_pdf(tmp_path / "scan.pdf")
    fake_mod = MagicMock()

    def fake_ocr(src: str, dest: str, **kwargs: object) -> None:
        Path(dest).write_bytes(Path(src).read_bytes())

    fake_mod.ocr.side_effect = fake_ocr

    with (
        patch("coversheets.pdf_ops.ocr_available", return_value=True),
        patch.dict("sys.modules", {"ocrmypdf": fake_mod}),
    ):
        run_ocr(path, language="eng", skip_text=True)

    fake_mod.ocr.assert_called_once()
    kwargs = fake_mod.ocr.call_args.kwargs
    assert kwargs["language"] == "eng"
    assert kwargs["skip_text"] is True
