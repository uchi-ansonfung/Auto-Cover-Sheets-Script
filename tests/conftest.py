"""Shared fixtures for coversheets tests."""

from __future__ import annotations

from pathlib import Path

import pytest
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas


@pytest.fixture
def sample_pdf(tmp_path: Path) -> Path:
    """Minimal one-page PDF for merge tests."""
    path = tmp_path / "Exhibit A.pdf"
    c = canvas.Canvas(str(path), pagesize=letter)
    c.drawString(72, 720, "Body page")
    c.save()
    return path


@pytest.fixture
def pdf_folder(tmp_path: Path) -> Path:
    """Folder with two input PDFs and one prior +output to ignore."""
    for name in ("Exhibit A.pdf", "Exhibit B.pdf"):
        path = tmp_path / name
        c = canvas.Canvas(str(path), pagesize=letter)
        c.drawString(72, 720, name)
        c.save()
    # Prior output (should never be treated as input)
    prior = tmp_path / "+Old.pdf"
    prior.write_bytes(b"%PDF-1.4 prior output placeholder\n")
    return tmp_path
