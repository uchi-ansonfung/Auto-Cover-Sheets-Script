"""Locate and activate tools shipped next to a frozen/installed binary.

Windows full installer layout (next to coversheets.exe)::

    coversheets.exe
    tesseract/
      tesseract.exe
      tessdata/eng.traineddata
      ...
    ghostscript/
      bin/gswin64c.exe
      bin/gsdll64.dll
      lib/
      Resource/
      ...

Call :func:`configure_bundled_tools` early at process start so OCR helpers can
find Tesseract and Ghostscript without a system-wide install.

Portable Ghostscript also needs ``GS_LIB`` / ``GS_DLL`` so it can load its
runtime files (including JBIG2 decode support via jbig2dec). Without those,
scanned PDFs that use ``/JBIG2Decode`` often fail OCR with a jbig2dec error
even though ``gswin64c.exe`` is on ``PATH``.
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


def _ghostscript_dll_names() -> tuple[str, ...]:
    if sys.platform == "win32":
        return ("gsdll64.dll", "gsdll32.dll")
    return ()


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


def find_bundled_ghostscript_root(root: Path | None = None) -> Path | None:
    """
    Return the Ghostscript install root (parent of ``bin`` when present).

    Layout expected by Artifex/Chocolatey packages::

        ghostscript/
          bin/gswin64c.exe
          lib/
          Resource/
    """
    gs_bin = find_bundled_ghostscript_bin(root)
    if gs_bin is None:
        return None
    # Prefer the version root that owns lib/Resource next to bin/.
    if gs_bin.name.lower() == "bin":
        parent = gs_bin.parent
        if (parent / "lib").is_dir() or (parent / "Resource").is_dir():
            return parent
        return parent
    return gs_bin


def _ghostscript_lib_dirs(gs_root: Path) -> list[Path]:
    """Candidate GS_LIB directories under a Ghostscript root (existing only)."""
    candidates = (
        gs_root / "lib",
        gs_root / "Resource" / "Init",
        gs_root / "Resource",
        gs_root / "fonts",
        gs_root / "Resource" / "Font",
    )
    return [p for p in candidates if p.is_dir()]


def _configure_ghostscript_env(gs_bin: Path, found: dict[str, str]) -> None:
    """
    Point portable Ghostscript at its own lib/DLL tree.

    System installs usually get this from the registry; the Windows full
    installer copies Ghostscript next to the app, so registry entries are
    absent and GS_LIB/GS_DLL must be set explicitly. That is what enables
    JBIG2 decoding (jbig2dec linked into gsdll) for scanned exhibit PDFs.
    """
    # Artifex layout: <root>/bin/gswin64c.exe, <root>/lib, <root>/Resource
    gs_root = gs_bin.parent if gs_bin.name.lower() == "bin" else gs_bin
    found["ghostscript_root"] = str(gs_root)

    lib_dirs = _ghostscript_lib_dirs(gs_root)
    if lib_dirs:
        # Windows uses ';' in GS_LIB; POSIX uses ':'.
        sep = ";" if sys.platform == "win32" else os.pathsep
        gs_lib = sep.join(str(p) for p in lib_dirs)
        # Always prefer the bundled tree over any stale user/system value.
        os.environ["GS_LIB"] = gs_lib
        found["gs_lib"] = gs_lib

    for dll_name in _ghostscript_dll_names():
        dll_path = gs_bin / dll_name
        if dll_path.is_file():
            os.environ["GS_DLL"] = str(dll_path)
            found["gs_dll"] = str(dll_path)
            break


def configure_bundled_tools(*, force: bool = False) -> dict[str, str]:
    """
    Prepend bundled Tesseract/Ghostscript dirs to ``PATH`` and set tessdata.

    Also sets ``GS_LIB`` / ``GS_DLL`` for portable Ghostscript so JBIG2
    decode (jbig2dec) works without a system install.

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
        _configure_ghostscript_env(gs_bin, found)

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
