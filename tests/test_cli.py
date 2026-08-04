"""Tests for CLI argument parsing and main()."""

from __future__ import annotations

from pathlib import Path

import pytest

from coversheets.cli import build_parser, main, options_from_args


def test_parser_defaults() -> None:
    args = build_parser().parse_args([])
    assert args.folder is None
    assert args.output_dir is None
    assert args.no_compress is False
    assert args.force is False
    assert args.dry_run is False
    assert args.batch is False
    assert args.gui is False
    assert args.rename_to_label is False
    assert args.strip_metadata is True
    assert args.ocr is False
    assert args.optimize is True
    assert args.linearize is False


def test_parser_flags(tmp_path: Path) -> None:
    args = build_parser().parse_args(
        [
            str(tmp_path),
            "--batch",
            "-o",
            str(tmp_path / "out"),
            "--force",
            "--no-compress",
            "--rename-to-label",
            "--no-strip-metadata",
            "--ocr",
            "--ocr-language",
            "eng+spa",
            "--ocr-force",
            "--no-optimize",
            "--linearize",
            "-n",
        ]
    )
    assert args.folder == tmp_path
    assert args.output_dir == tmp_path / "out"
    assert args.force is True
    assert args.no_compress is True
    assert args.dry_run is True
    assert args.batch is True
    assert args.rename_to_label is True
    assert args.strip_metadata is False
    assert args.ocr is True
    assert args.ocr_language == "eng+spa"
    assert args.ocr_force is True
    assert args.optimize is False
    assert args.linearize is True
    assert args.vertical_position == "center"

    opts = options_from_args(args)
    assert opts.compress is False
    assert opts.strip_metadata is False
    assert opts.ocr is True
    assert opts.ocr_skip_text is False
    assert opts.optimize is False
    assert opts.linearize is True
    assert opts.vertical_position == "center"


def test_parser_vertical_position() -> None:
    args = build_parser().parse_args(["--vertical-position", "top_third"])
    assert args.vertical_position == "top_third"
    opts = options_from_args(args)
    assert opts.vertical_position == "top_third"


def test_main_rejects_missing_dir_in_batch(tmp_path: Path) -> None:
    missing = tmp_path / "nope"
    assert main(["--batch", str(missing)]) == 1


def test_main_dry_run_success(pdf_folder: Path) -> None:
    assert main([str(pdf_folder), "--dry-run"]) == 0


def test_main_batch_success(pdf_folder: Path) -> None:
    assert main(["--batch", str(pdf_folder), "--no-compress"]) == 0
    assert (pdf_folder / "+Exhibit A.pdf").is_file()


def test_main_rejects_gui_with_batch(pdf_folder: Path) -> None:
    with pytest.raises(SystemExit):
        main(["--gui", "--batch", str(pdf_folder)])
