"""Tests for bundled Tesseract discovery (and optional legacy Ghostscript)."""

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


def test_configure_bundled_tools_tesseract_only(
    tmp_path: Path, monkeypatch
) -> None:
    """Primary full-installer layout: Tesseract only (no Ghostscript)."""
    tess = tmp_path / "tesseract"
    tess.mkdir()
    (tess / _tess_name()).write_bytes(b"fake")
    tessdata = tess / "tessdata"
    tessdata.mkdir()
    (tessdata / "eng.traineddata").write_bytes(b"fake")

    monkeypatch.setenv("COVERSHEETS_INSTALL_ROOT", str(tmp_path))
    monkeypatch.setattr(bt, "_CONFIGURED", False)
    monkeypatch.setenv("PATH", "C:\\existing" if sys.platform == "win32" else "/existing")
    monkeypatch.delenv("TESSDATA_PREFIX", raising=False)

    found = bt.configure_bundled_tools(force=True)

    assert found["tesseract"] == str(tess)
    assert found["tessdata"] == str(tessdata)
    assert "ghostscript" not in found
    path = os.environ["PATH"]
    assert path.startswith(str(tess))
    assert os.environ["TESSDATA_PREFIX"] == str(tessdata)


def test_configure_bundled_tools_optional_legacy_ghostscript(
    tmp_path: Path, monkeypatch
) -> None:
    """Legacy ghostscript\\ next to the app is still activated if present."""
    tess = tmp_path / "tesseract"
    tess.mkdir()
    (tess / _tess_name()).write_bytes(b"fake")
    tessdata = tess / "tessdata"
    tessdata.mkdir()
    (tessdata / "eng.traineddata").write_bytes(b"fake")

    gs_bin = tmp_path / "ghostscript" / "bin"
    gs_bin.mkdir(parents=True)
    (gs_bin / _gs_name()).write_bytes(b"fake")
    (tmp_path / "ghostscript" / "lib").mkdir()
    (tmp_path / "ghostscript" / "Resource" / "Init").mkdir(parents=True)
    dll_name = "gsdll64.dll" if sys.platform == "win32" else None
    if dll_name:
        (gs_bin / dll_name).write_bytes(b"fake")

    monkeypatch.setenv("COVERSHEETS_INSTALL_ROOT", str(tmp_path))
    monkeypatch.setattr(bt, "_CONFIGURED", False)
    monkeypatch.setenv("PATH", "C:\\existing" if sys.platform == "win32" else "/existing")
    monkeypatch.delenv("TESSDATA_PREFIX", raising=False)
    monkeypatch.delenv("GS_LIB", raising=False)
    monkeypatch.delenv("GS_DLL", raising=False)

    found = bt.configure_bundled_tools(force=True)

    assert found["tesseract"] == str(tess)
    assert found["ghostscript"] == str(gs_bin)
    assert found["tessdata"] == str(tessdata)
    assert found["ghostscript_root"] == str(tmp_path / "ghostscript")
    path = os.environ["PATH"]
    assert path.startswith(str(tess))
    assert str(gs_bin) in path
    assert os.environ["TESSDATA_PREFIX"] == str(tessdata)
    gs_lib = os.environ["GS_LIB"]
    assert str(tmp_path / "ghostscript" / "lib") in gs_lib
    assert str(tmp_path / "ghostscript" / "Resource" / "Init") in gs_lib
    if dll_name:
        assert os.environ["GS_DLL"] == str(gs_bin / dll_name)


def test_find_bundled_ghostscript_root(tmp_path: Path) -> None:
    gs_bin = tmp_path / "ghostscript" / "bin"
    gs_bin.mkdir(parents=True)
    (gs_bin / _gs_name()).write_bytes(b"fake")
    (tmp_path / "ghostscript" / "lib").mkdir()
    assert bt.find_bundled_ghostscript_root(tmp_path) == tmp_path / "ghostscript"


def test_find_bundled_tesseract_missing(tmp_path: Path) -> None:
    assert bt.find_bundled_tesseract_dir(tmp_path) is None


def test_rasterizer_available_prefers_pypdfium(monkeypatch) -> None:
    monkeypatch.setattr(bt, "pypdfium_available", lambda: True)
    monkeypatch.setattr(bt, "ghostscript_on_path", lambda: False)
    assert bt.rasterizer_available() is True


def test_rasterizer_available_falls_back_to_ghostscript(monkeypatch) -> None:
    monkeypatch.setattr(bt, "pypdfium_available", lambda: False)
    monkeypatch.setattr(bt, "ghostscript_on_path", lambda: True)
    assert bt.rasterizer_available() is True


def test_rasterizer_available_false_when_neither(monkeypatch) -> None:
    monkeypatch.setattr(bt, "pypdfium_available", lambda: False)
    monkeypatch.setattr(bt, "ghostscript_on_path", lambda: False)
    assert bt.rasterizer_available() is False
