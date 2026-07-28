"""Locate and activate tools shipped next to a frozen/installed binary.

Windows full installer layout (next to coversheets.exe)::

    coversheets.exe
    tesseract/
      tesseract.exe
      tessdata/eng.traineddata
      ...
    ghostscript/
      bin/gswin64c.exe
      ...

Call :func:`configure_bundled_tools` early at process start so OCR helpers can
find Tesseract and Ghostscript without a system-wide install.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from shutil import which

_CONFIGURED = False


def install_root() -> Path | None:
    """
    Directory that contains the app binary (and optional tool folders).

    - Frozen (PyInstaller): parent of ``sys.executable``
    - Override: ``COVERSHEETS_INSTALL_ROOT`` (useful in tests)
    - Dev/source: ``None`` unless the env override is set
    """
    override = (os.environ.get("COVERSHEETS_INSTALL_ROOT") or "").strip()
    if override:
        path = Path(override).expanduser()
        try:
            return path.resolve()
        except OSError:
            return path
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return None


def _tesseract_names() -> tuple[str, ...]:
    if sys.platform == "win32":
        return ("tesseract.exe", "tesseract")
    return ("tesseract",)


def _ghostscript_names() -> tuple[str, ...]:
    if sys.platform == "win32":
        return ("gswin64c.exe", "gswin32c.exe", "gs.exe", "gs")
    return ("gs",)


def find_bundled_tesseract_dir(root: Path | None = None) -> Path | None:
    """Return the directory containing a bundled ``tesseract`` binary, if any."""
    base = root if root is not None else install_root()
    if base is None:
        return None
    candidates = (
        base / "tesseract",
        base / "tools" / "tesseract",
    )
    names = _tesseract_names()
    for directory in candidates:
        if any((directory / name).is_file() for name in names):
            return directory
    return None


def find_bundled_ghostscript_bin(root: Path | None = None) -> Path | None:
    """Return the directory containing a bundled Ghostscript CLI, if any."""
    base = root if root is not None else install_root()
    if base is None:
        return None
    candidates = (
        base / "ghostscript" / "bin",
        base / "ghostscript",
        base / "tools" / "ghostscript" / "bin",
        base / "tools" / "ghostscript",
    )
    names = _ghostscript_names()
    for directory in candidates:
        if any((directory / name).is_file() for name in names):
            return directory
    return None


def configure_bundled_tools(*, force: bool = False) -> dict[str, str]:
    """
    Prepend bundled Tesseract/Ghostscript dirs to ``PATH`` and set tessdata.

    Safe to call multiple times; subsequent calls are no-ops unless ``force``.

    Returns paths that were activated, e.g.
    ``{"tesseract": "...", "ghostscript": "...", "tessdata": "..."}``.
    """
    global _CONFIGURED
    found: dict[str, str] = {}
    if _CONFIGURED and not force:
        return found

    root = install_root()
    path_parts: list[str] = []

    tess_dir = find_bundled_tesseract_dir(root)
    if tess_dir is not None:
        path_parts.append(str(tess_dir))
        found["tesseract"] = str(tess_dir)
        tessdata = tess_dir / "tessdata"
        if tessdata.is_dir():
            # Point at the tessdata directory for predictable frozen layouts.
            os.environ.setdefault("TESSDATA_PREFIX", str(tessdata))
            found["tessdata"] = str(tessdata)

    gs_bin = find_bundled_ghostscript_bin(root)
    if gs_bin is not None:
        path_parts.append(str(gs_bin))
        found["ghostscript"] = str(gs_bin)

    if path_parts:
        existing = os.environ.get("PATH", "")
        os.environ["PATH"] = os.pathsep.join(
            path_parts + ([existing] if existing else [])
        )

    _CONFIGURED = True
    return found


def tesseract_on_path() -> bool:
    """True if a ``tesseract`` executable is discoverable after configuration."""
    configure_bundled_tools()
    for name in _tesseract_names():
        if which(name) is not None:
            return True
    return False


def ghostscript_on_path() -> bool:
    """True if a Ghostscript CLI is discoverable after configuration."""
    configure_bundled_tools()
    for name in _ghostscript_names():
        if which(name) is not None:
            return True
    return False
