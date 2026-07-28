# Automatic Exhibit Cover Sheets

Add letter-sized cover sheets to PDFs. Primary UX is a **GUI list** (edit labels, generate with progress). Headless batch mode remains available for scripts.

**v0.8** — PDF options: true metadata strip, optional OCR, optimize, and linearize.

## Install (from source)

```bash
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -e ".[dev]"
```

Optional extras:

```bash
pip install -e ".[ocr]"        # ocrmypdf (also install system Tesseract)
pip install -e ".[optimize]"   # pikepdf for linearize + thorough metadata strip
pip install -e ".[full]"       # both
```

Or with the pinned runtime list only:

```bash
pip install -r requirements.txt
```

## First run (prebuilt binaries)

1. Download the binary for your OS from the latest [release](https://github.com/uchi-ansonfung/Auto-Cover-Sheets-Script/releases).
   - Windows: `coversheets-<version>-windows-x64.exe`
   - macOS Apple Silicon: `coversheets-<version>-macos-arm64`
   - macOS Intel: `coversheets-<version>-macos-x64`
2. Windows: dismiss SmartScreen if needed (More info → Run anyway).  
   macOS: right-click → Open the first time (unsigned builds are blocked by Gatekeeper).

## Usage (GUI list — default)

```bash
coversheets
# or
python -m coversheets
# optional: preload a folder into the list
coversheets /path/to/pdfs
```

1. **Open Folder…** or **Add PDFs…** to populate the list.
2. Edit each **Cover Label** by double-clicking the cell (defaults to the filename without `.pdf`).
3. Click **Include** to toggle rows; remove rows you don’t want.
4. Optionally set an **output folder** (empty = write next to each source file).
5. Choose options, then **Generate Cover Sheets**.
6. A **progress window** shows status/log and a **Cancel** button (stops after the current file). A **success dialog** can open the output folder.
7. Outputs are written **atomically** (temp `.partial` then rename) so crashes don’t leave truncated finals. Default name: `+OriginalName.pdf` (or `+CoverLabel.pdf` with rename). Originals are not modified.

The GUI **remembers preferences** between runs (checkboxes, OCR language, window size, output folder, and last input folder). On launch it reloads the last folder if it still exists. Settings are stored under your user config directory (`~/Library/Application Support/coversheets/` on macOS, `%APPDATA%\\coversheets\\` on Windows, `~/.config/coversheets/` on Linux).

### PDF options (GUI + CLI)

| Option | Default | Notes |
|--------|---------|--------|
| Compress page streams | on | Lossless content-stream compression |
| **Strip metadata** | on | Removes document Info + XMP (true strip, not only “don’t copy”) |
| **Optimize** | on | Dedupe identical objects / drop orphans (pypdf) |
| **Linearize** | off | Web-optimize; needs `coversheets[optimize]` (pikepdf) or `qpdf` |
| **OCR** | off | Makes text searchable via ocrmypdf; needs `coversheets[ocr]` + Tesseract |
| Name output after label | off | `+Label.pdf` instead of `+OriginalName.pdf` |
| Open folder when done | on | GUI only |

## Usage (headless batch)

```bash
coversheets --batch /path/to/pdfs
coversheets /path/to/pdfs --dry-run
coversheets --batch /path/to/pdfs --force --no-compress -o /path/to/output
coversheets --batch /path/to/pdfs --rename-to-label
coversheets --batch /path/to/pdfs --strip-metadata --linearize
coversheets --batch /path/to/pdfs --no-optimize
coversheets --batch /path/to/pdfs --ocr --ocr-language eng
coversheets --batch /path/to/pdfs --no-strip-metadata
coversheets --version
```

| Flag | Meaning |
|------|---------|
| `folder` | Optional folder (GUI preload, or required with `--batch`) |
| `--gui` | Force GUI (default when not batching) |
| `--batch` | Headless process all PDFs in folder |
| `-o`, `--output-dir DIR` | Write `+Name.pdf` here (batch) |
| `--no-compress` | Skip lossless page-stream compression |
| `--force` | Overwrite existing `+Name.pdf` |
| `--rename-to-label` | Name outputs `+Label.pdf` |
| `--strip-metadata` / `--no-strip-metadata` | True metadata strip (default on) |
| `--optimize` / `--no-optimize` | Dedupe PDF objects (default on) |
| `--linearize` | Linearize output (pikepdf or qpdf) |
| `--ocr` | OCR with ocrmypdf |
| `--ocr-language LANG` | Tesseract language(s), default `eng` |
| `--ocr-force` | OCR even when text already exists |
| `-n`, `--dry-run` | Show plan only; implies batch |
| `-V`, `--version` | Print version and exit |

## What this program does

* Adds a letter-sized coversheet with centered text (36-pt Times Bold; wraps if needed).
* Cover text comes from the **editable label** in the GUI (or the filename stem in batch mode).
* Optional **true metadata strip** (document Info dictionary + XMP).
* Optional **OCR** path (ocrmypdf) so scanned bodies become text-searchable.
* Optional **optimize** (object dedupe) and **linearize** (fast web view).
* Optional **name output after label**.
* Continues after per-file errors and reports a summary.

## What this program doesn't do

* Does not modify the original PDFs.
* OCR/linearize extras are optional and disabled in the UI when dependencies are missing.

## Development

```bash
pip install -e ".[dev]"
pytest
```

### Layout

```
coversheets/
  __main__.py         # python -m coversheets
  gui.py              # tkinter list UI
  cover.py            # cover sheet generation (ReportLab)
  merge.py            # prepend cover + write
  pdf_ops.py          # metadata strip, OCR, optimize, linearize
  options.py          # ProcessOptions
  prefs.py            # GUI preference persistence
  process.py          # JobItem + batch processing
  cli.py              # argparse entry point
tests/
pyproject.toml
```

## Release builds (CI)

Publishing a GitHub Release runs [`.github/workflows/release.yml`](.github/workflows/release.yml), which:

1. Runs tests on Windows + macOS (arm64 and x64)
2. Builds one-file PyInstaller binaries
3. Attaches them to the release as:
   - `coversheets-<version>-windows-x64.exe`
   - `coversheets-<version>-macos-arm64`
   - `coversheets-<version>-macos-x64`

You can also run the workflow manually (**Actions → Release builds → Run workflow**) to produce downloadable workflow artifacts without creating a release.

```bash
# Tag and publish a release (example)
git tag v0.8.0
git push origin v0.8.0
# Then create a Release from that tag in the GitHub UI (or: gh release create v0.8.0)
```

## Building locally (optional)

```bash
pip install -e ".[dev]"
pyinstaller coversheets.spec
# → dist/coversheets.exe  (Windows)
# → dist/coversheets      (macOS / Linux)
```
