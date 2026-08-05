# Windows full installer

Produces a per-user Setup wizard that installs:

- `coversheets.exe` (PyInstaller one-file with `pikepdf` + `ocrmypdf` + `pypdfium2`)
- Bundled `tesseract\` (English OCR data)

Ghostscript is **not** required. OCRmyPDF 17+ rasterizes pages with pypdfium2
(PDFium), which is collected into the frozen exe.

## Build

From the repo root on Windows:

```powershell
pip install -e ".[full,dev]"
powershell -ExecutionPolicy Bypass -File scripts\build_windows_full.ps1
```

Output:

- `dist\coversheets-<version>-windows-x64-setup.exe`
- `dist\windows-full\` (payload folder compiled by Inno Setup)

## Layout expectations

`coversheets.iss` reads `{#PayloadDir}` (default `dist\windows-full`):

```text
windows-full/
  coversheets.exe
  README-INSTALLED.txt
  tesseract/
    tesseract.exe
    tessdata/eng.traineddata
    ...
```

At runtime, `coversheets.bundled_tools` prepends `tesseract\` to `PATH` and
sets `TESSDATA_PREFIX` when a bundled tessdata tree is present.

**JBIG2:** Scanned exhibit PDFs often use `/JBIG2Decode`. OCR rasterizes those
pages through **pypdfium2** (bundled in the exe), not Ghostscript/jbig2dec.
A legacy `ghostscript\` folder next to the exe is still detected if present,
but current builds do not ship it.
