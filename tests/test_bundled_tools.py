"""Tests for bundled Tesseract/Ghostscript discovery."""

from __future__ import annotations

import os
import sys
from pathlib import Path

import coversheets.bundled_tools as bt


def _tess_name() -> str:
    return "tesseract.exe" if sys.platform == "win32" else "tesseract"


def _gs_name() -> str:
    return "gswin64c.exe" if sys.platform == "win32" else "gs"


def test_configure_bundled_tools_noop_without_root(monkeypatch) -> None:
    monkeypatch.delenv("COVERSHEETS_INSTALL_ROOT", raising=False)
    monkeypatch.setattr(bt, "_CONFIGURED", False)
    monkeypatch.setattr(bt, "install_root", lambda: None)
    found = bt.configure_bundled_tools(force=True)
    assert found == {}


def test_configure_bundled_tools_prepends_path(
    tmp_path: Path, monkeypatch
) -> None:
    tess = tmp_path / "tesseract"
    tess.mkdir()
    (tess / _tess_name()).write_bytes(b"fake")
    tessdata = tess / "tessdata"
    tessdata.mkdir()
    (tessdata / "eng.traineddata").write_bytes(b"fake")

    gs_bin = tmp_path / "ghostscript" / "bin"
    gs_bin.mkdir(parents=True)
    (gs_bin / _gs_name()).write_bytes(b"fake")

    monkeypatch.setenv("COVERSHEETS_INSTALL_ROOT", str(tmp_path))
    monkeypatch.setattr(bt, "_CONFIGURED", False)
    monkeypatch.setenv("PATH", "C:\\existing" if sys.platform == "win32" else "/existing")
    monkeypatch.delenv("TESSDATA_PREFIX", raising=False)

    found = bt.configure_bundled_tools(force=True)

    assert found["tesseract"] == str(tess)
    assert found["ghostscript"] == str(gs_bin)
    assert found["tessdata"] == str(tessdata)
    path = os.environ["PATH"]
    assert path.startswith(str(tess))
    assert str(gs_bin) in path
    assert os.environ["TESSDATA_PREFIX"] == str(tessdata)


def test_find_bundled_tesseract_missing(tmp_path: Path) -> None:
    assert bt.find_bundled_tesseract_dir(tmp_path) is None
