# -*- mode: python ; coding: utf-8 -*-
# Build: pyinstaller coversheets.spec

a = Analysis(
    ['coversheets/__main__.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['coversheets', 'coversheets.cli', 'coversheets.cover', 'coversheets.merge', 'coversheets.process', 'coversheets.gui', 'coversheets.util', 'coversheets.options', 'coversheets.pdf_ops', 'coversheets.prefs'],
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
    name='coversheets',
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
