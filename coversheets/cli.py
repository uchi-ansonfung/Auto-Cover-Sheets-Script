"""Command-line interface and entry points."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from coversheets import __author__, __version__
from coversheets.bundled_tools import configure_bundled_tools
from coversheets.cover import DEFAULT_VERTICAL_POSITION, VERTICAL_POSITIONS
from coversheets.options import ProcessOptions
from coversheets.process import process_folder


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="coversheets",
        description=(
            "Add letter-sized cover sheets to PDFs. Default: open a GUI list "
            "where you can edit cover labels. Use --batch for headless "
            "processing (filename stems as labels)."
        ),
    )
    parser.add_argument(
        "folder",
        nargs="?",
        default=None,
        type=Path,
        help=(
            "Optional folder of PDFs. With --batch/--dry-run, process headlessly. "
            "Otherwise open the GUI with this folder loaded."
        ),
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Open the GUI list (default when not using --batch/--dry-run).",
    )
    parser.add_argument(
        "--batch",
        action="store_true",
        help=(
            "Headless mode: process all PDFs in folder using filename stems "
            "as cover labels (no GUI)."
        ),
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default=None,
        type=Path,
        metavar="DIR",
        help="Write +Name.pdf files here (batch mode; default: input folder).",
    )
    parser.add_argument(
        "--no-compress",
        action="store_true",
        help="Skip lossless page-stream compression (batch mode).",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Overwrite existing +Name.pdf outputs (batch mode).",
    )
    parser.add_argument(
        "--rename-to-label",
        action="store_true",
        help=(
            "Name outputs after the cover label (+Label.pdf) instead of "
            "+OriginalName.pdf (batch mode; also available as a GUI toggle)."
        ),
    )
    parser.add_argument(
        "--strip-metadata",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Remove document Info and XMP metadata (default: on).",
    )
    parser.add_argument(
        "--ocr",
        action="store_true",
        help=(
            "OCR output PDFs with ocrmypdf (requires: pip install "
            "'coversheets[ocr]' and system Tesseract)."
        ),
    )
    parser.add_argument(
        "--ocr-language",
        default="eng",
        metavar="LANG",
        help="Tesseract language code(s) for --ocr (default: eng).",
    )
    parser.add_argument(
        "--ocr-force",
        action="store_true",
        help="With --ocr, OCR pages even if they already contain text.",
    )
    parser.add_argument(
        "--optimize",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Dedupe identical PDF objects (structure optimize; default: on).",
    )
    parser.add_argument(
        "--linearize",
        action="store_true",
        help=(
            "Linearize (web-optimize) outputs (requires pikepdf or qpdf; "
            "pip install 'coversheets[optimize]')."
        ),
    )
    parser.add_argument(
        "--vertical-position",
        choices=sorted(VERTICAL_POSITIONS),
        default=DEFAULT_VERTICAL_POSITION,
        metavar="POS",
        help=(
            "Vertical placement of cover title text: center (default) or "
            "top_third (batch mode)."
        ),
    )
    parser.add_argument(
        "-n",
        "--dry-run",
        action="store_true",
        help="List what would be done without writing files (implies --batch).",
    )
    parser.add_argument(
        "-V",
        "--version",
        action="version",
        version=f"Automatic Exhibit Cover Sheets v{__version__} ({__author__})",
    )
    return parser


def options_from_args(args: argparse.Namespace) -> ProcessOptions:
    return ProcessOptions(
        compress=not args.no_compress,
        force=args.force,
        dry_run=args.dry_run,
        rename_to_label=args.rename_to_label,
        strip_metadata=args.strip_metadata,
        ocr=args.ocr,
        ocr_language=args.ocr_language,
        ocr_skip_text=not args.ocr_force,
        optimize=args.optimize,
        linearize=args.linearize,
        vertical_position=args.vertical_position,
    )


def main(argv: list[str] | None = None) -> int:
    # Frozen/installer builds ship Tesseract next to the exe (pypdfium2 is baked in).
    configure_bundled_tools()

    parser = build_parser()
    args = parser.parse_args(argv)

    if args.gui and (args.batch or args.dry_run):
        parser.error("Cannot combine --gui with --batch or --dry-run")

    print(f"Automatic Exhibit Cover Sheets v{__version__}.  {__author__}")

    use_batch = bool(args.batch or args.dry_run)

    if not use_batch:
        from coversheets.gui import run_app

        initial: Path | None = None
        if args.folder is not None:
            initial = args.folder.expanduser().resolve()
            if not initial.is_dir():
                print(f"Not a valid directory: {initial}", file=sys.stderr)
                return 1
        return run_app(initial_folder=initial)

    # --- Headless batch --------------------------------------------------
    if args.folder is None:
        print(
            "Batch mode requires a folder path.\n"
            "  coversheets --batch /path/to/pdfs\n"
            "Or omit --batch to open the GUI.",
            file=sys.stderr,
        )
        return 1

    folder = args.folder.expanduser().resolve()
    if not folder.is_dir():
        print(f"Not a valid directory: {folder}", file=sys.stderr)
        return 1

    output_dir: Path | None = None
    if args.output_dir is not None:
        output_dir = args.output_dir.expanduser().resolve()
        if not args.dry_run:
            output_dir.mkdir(parents=True, exist_ok=True)

    result = process_folder(
        folder,
        output_dir=output_dir,
        options=options_from_args(args),
    )
    return 0 if result.ok else 1


def run() -> None:
    """Console-script entry point (``coversheets`` command)."""
    raise SystemExit(main())


if __name__ == "__main__":
    run()
