# Windows full installer

Produces a per-user Setup wizard that installs:

- `coversheets.exe` (PyInstaller one-file with `pikepdf` + `ocrmypdf`)
- Bundled `tesseract\` (English OCR data)
- Bundled `ghostscript\` (required by ocrmypdf)

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
  ghostscript/
    bin/gswin64c.exe
    ...
```

At runtime, `coversheets.bundled_tools` prepends those folders to `PATH` and sets
`GS_LIB` / `GS_DLL` for the portable Ghostscript tree.

**JBIG2 / jbig2dec:** Many scanned exhibit PDFs compress pages with
`/JBIG2Decode`. OCR rasterizes those pages through Ghostscript, which needs its
full install tree (`lib\`, `Resource\`, `gsdll64.dll`) — not just `gswin64c.exe`.
That is jbig2dec (decode), not the optional jbig2enc encoder used for output
optimization. Stage the whole Ghostscript folder from Program Files / Chocolatey;
do not strip it down to `bin\` only.
