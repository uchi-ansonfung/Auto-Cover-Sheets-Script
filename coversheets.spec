# -*- mode: python ; coding: utf-8 -*-
# Build (slim):     pip install -e ".[dev]" && pyinstaller coversheets.spec
# Build (full):     pip install -e ".[full,dev]" && pyinstaller coversheets.spec
#
# Full builds collect pikepdf + ocrmypdf when installed. Native Tesseract and
# Ghostscript are staged next to the exe by scripts/build_windows_full.ps1
# (not embedded in this Analysis).

from PyInstaller.utils.hooks import collect_all

hiddenimports = [
    "coversheets",
    "coversheets.cli",
    "coversheets.cover",
    "coversheets.merge",
    "coversheets.process",
    "coversheets.gui",
    "coversheets.util",
    "coversheets.options",
    "coversheets.pdf_ops",
    "coversheets.prefs",
    "coversheets.bundled_tools",
]

datas = []
binaries = []


def _try_collect(package: str) -> None:
    """Pull in package data/binaries/imports when the extra is installed."""
    global datas, binaries, hiddenimports
    try:
        pkg_datas, pkg_binaries, pkg_hidden = collect_all(package)
    except Exception:
        return
    datas += pkg_datas
    binaries += pkg_binaries
    hiddenimports += pkg_hidden


# Optional extras — only present for full installer / local .[full] builds.
_try_collect("pikepdf")
_try_collect("ocrmypdf")
# Common ocrmypdf runtime pieces that are sometimes missed by analysis.
for _pkg in (
    "pdfminer",
    "pdfminer.six",
    "img2pdf",
    "pi_heif",
    "PIL",
    "packaging",
    "deprecation",
    "rich",
):
    _try_collect(_pkg)

a = Analysis(
    ["coversheets/__main__.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="coversheets",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
