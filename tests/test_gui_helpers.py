"""Tests for GUI helpers that do not require a display."""

from __future__ import annotations

from pathlib import Path

from coversheets.process import BatchResult, JobItem
from coversheets.util import format_result_summary, resolve_app_asset, resolve_result_folders


def test_resolve_app_asset_finds_icon() -> None:
    png = resolve_app_asset("app-icon.png")
    ico = resolve_app_asset("app-icon.ico")
    assert png is not None and png.is_file()
    assert ico is not None and ico.is_file()
    assert resolve_app_asset("definitely-missing-icon-xyz.xyz") is None


def test_format_result_summary() -> None:
    result = BatchResult(total=3, succeeded=2, skipped=1, failed=0)
    text = format_result_summary(result)
    assert "2 of 3" in text
    assert "skipped 1" in text
    assert "failed 0" in text


def test_format_result_summary_cancelled() -> None:
    result = BatchResult(
        total=5, succeeded=2, skipped=0, failed=0, cancelled=3, was_cancelled=True
    )
    text = format_result_summary(result)
    assert "Cancelled" in text
    assert "cancelled 3" in text


def test_resolve_result_folders_explicit_output(tmp_path: Path) -> None:
    out = tmp_path / "out"
    out.mkdir()
    job = JobItem(source=tmp_path / "a.pdf", label="a", include=True)
    folders = resolve_result_folders([job], out)
    assert folders == [out.resolve()]


def test_resolve_result_folders_source_parents(tmp_path: Path) -> None:
    a = tmp_path / "a"
    b = tmp_path / "b"
    a.mkdir()
    b.mkdir()
    j1 = JobItem(source=a / "1.pdf", label="1", include=True)
    j2 = JobItem(source=b / "2.pdf", label="2", include=True)
    j3 = JobItem(source=a / "3.pdf", label="3", include=False)
    folders = resolve_result_folders([j1, j2, j3], None)
    assert folders == sorted(
        [a.resolve(), b.resolve()], key=lambda p: str(p).casefold()
    )
