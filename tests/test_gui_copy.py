"""Tests for plain-language GUI helpers (no display required)."""

from __future__ import annotations

from coversheets.gui.copy import (
    done_headline,
    output_example,
    plain_option_lines,
    preview_target_index,
    status_for_jobs,
    truncate_middle,
)
from coversheets.gui.dnd import parse_drop_paths
from coversheets.options import ProcessOptions
from coversheets.process import BatchResult


def test_truncate_middle() -> None:
    assert truncate_middle("short", 40) == "short"
    long = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    out = truncate_middle(long, 12)
    assert "…" in out
    assert len(out) == 12


def test_preview_target_index() -> None:
    assert preview_target_index(set(), anchor=None, job_count=0) is None
    assert preview_target_index(set(), anchor=None, job_count=1) == 0
    assert preview_target_index({2, 0}, anchor=2, job_count=3) == 2
    assert preview_target_index({1}, anchor=None, job_count=3) == 1


def test_status_for_jobs() -> None:
    assert "begin" in status_for_jobs(0, 0).lower()
    assert "will get cover sheets" in status_for_jobs(3, 2)
    assert "none selected" in status_for_jobs(3, 0)


def test_output_example_beside() -> None:
    text = output_example(mode="beside", folder=None, rename_to_label=False)
    assert "+Contract.pdf" in text
    assert "next to" in text.lower()


def test_plain_option_lines() -> None:
    lines = plain_option_lines(
        ProcessOptions(strip_metadata=True, ocr=True, ocr_language="eng")
    )
    assert any("searchable" in line.lower() for line in lines)
    assert any("document info" in line.lower() for line in lines)


def test_done_headline_success() -> None:
    title, headline = done_headline(
        BatchResult(total=2, succeeded=2, skipped=0, failed=0)
    )
    assert title == "Done"
    assert "+" in headline


def test_parse_drop_paths_braces() -> None:
    paths = parse_drop_paths("{/tmp/My File.pdf} /tmp/other.pdf")
    assert len(paths) == 2
    assert paths[0].name == "My File.pdf"
    assert paths[1].name == "other.pdf"
