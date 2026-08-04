"""Tests for batch job processing rules."""

from __future__ import annotations

from pathlib import Path

from coversheets.options import ProcessOptions
from coversheets.process import (
    JobItem,
    collect_pdfs,
    is_output_filename,
    jobs_from_folder,
    output_path_for,
    process_folder,
    process_jobs,
    sanitize_filename_stem,
)


def test_is_output_filename() -> None:
    assert is_output_filename("+Exhibit A.pdf") is True
    assert is_output_filename("Exhibit A.pdf") is False


def test_collect_pdfs_skips_plus_prefix(pdf_folder: Path) -> None:
    names = collect_pdfs(pdf_folder)
    assert names == ["Exhibit A.pdf", "Exhibit B.pdf"]
    assert all(not n.startswith("+") for n in names)


def test_output_path_for() -> None:
    dest = output_path_for("Exhibit A.pdf", Path("/tmp/out"))
    assert dest == Path("/tmp/out/+Exhibit A.pdf")


def test_sanitize_filename_stem() -> None:
    assert sanitize_filename_stem("Exhibit A") == "Exhibit A"
    assert sanitize_filename_stem("A/B:C*?.pdf") == "ABC.pdf"
    assert sanitize_filename_stem("  foo   bar  ") == "foo bar"
    assert sanitize_filename_stem("///") == "untitled"
    assert sanitize_filename_stem("con") == "con_file"


def test_output_path_rename_to_label() -> None:
    dest = output_path_for(
        "scan_001.pdf",
        Path("/tmp/out"),
        label="Exhibit A — Contract",
        rename_to_label=True,
    )
    assert dest == Path("/tmp/out/+Exhibit A — Contract.pdf")


def test_output_path_rename_disambiguates_duplicates() -> None:
    used: set[str] = set()
    a = output_path_for(
        "a.pdf", Path("/tmp"), label="Same", rename_to_label=True, used_basenames=used
    )
    b = output_path_for(
        "b.pdf", Path("/tmp"), label="Same", rename_to_label=True, used_basenames=used
    )
    assert a.name == "+Same.pdf"
    assert b.name == "+Same (2).pdf"


def test_job_item_default_label(sample_pdf: Path) -> None:
    job = JobItem.from_path(sample_pdf)
    assert job.label == "Exhibit A"
    assert job.include is True


def test_process_jobs_uses_custom_label(sample_pdf: Path, tmp_path: Path) -> None:
    job = JobItem.from_path(sample_pdf, label="EXHIBIT 1 — Contract")
    out_dir = tmp_path / "out"
    result = process_jobs(
        [job],
        output_dir=out_dir,
        options=ProcessOptions(compress=False),
    )
    assert result.ok
    assert result.succeeded == 1
    assert (out_dir / f"+{sample_pdf.name}").is_file()


def test_process_jobs_top_third_vertical_position(
    sample_pdf: Path, tmp_path: Path
) -> None:
    job = JobItem.from_path(sample_pdf, label="Top Third Title")
    out_dir = tmp_path / "out"
    result = process_jobs(
        [job],
        output_dir=out_dir,
        options=ProcessOptions(compress=False, vertical_position="top_third"),
    )
    assert result.ok
    assert result.succeeded == 1
    assert (out_dir / f"+{sample_pdf.name}").is_file()


def test_process_jobs_rename_to_label(sample_pdf: Path, tmp_path: Path) -> None:
    job = JobItem.from_path(sample_pdf, label="Exhibit Z")
    out_dir = tmp_path / "out"
    result = process_jobs(
        [job],
        output_dir=out_dir,
        options=ProcessOptions(compress=False, rename_to_label=True),
    )
    assert result.ok
    assert (out_dir / "+Exhibit Z.pdf").is_file()
    assert not (out_dir / f"+{sample_pdf.name}").exists()


def test_process_jobs_skips_excluded(sample_pdf: Path, tmp_path: Path) -> None:
    job = JobItem.from_path(sample_pdf)
    job.include = False
    result = process_jobs(
        [job], output_dir=tmp_path, options=ProcessOptions(compress=False)
    )
    assert result.total == 0
    assert result.succeeded == 0


def test_process_folder_dry_run_writes_nothing(pdf_folder: Path) -> None:
    result = process_folder(pdf_folder, options=ProcessOptions(dry_run=True))
    assert result.ok
    assert result.succeeded == 2
    assert result.failed == 0
    assert not list(pdf_folder.glob("+Exhibit*.pdf"))


def test_process_folder_writes_outputs_and_skips_on_rerun(pdf_folder: Path) -> None:
    first = process_folder(pdf_folder, options=ProcessOptions(compress=False))
    assert first.ok
    assert first.succeeded == 2
    assert (pdf_folder / "+Exhibit A.pdf").is_file()
    assert (pdf_folder / "+Exhibit B.pdf").is_file()

    second = process_folder(pdf_folder, options=ProcessOptions(compress=False))
    assert second.ok
    assert second.skipped == 2
    assert second.succeeded == 0


def test_process_folder_force_overwrites(pdf_folder: Path) -> None:
    process_folder(pdf_folder, options=ProcessOptions(compress=False))
    again = process_folder(
        pdf_folder, options=ProcessOptions(compress=False, force=True)
    )
    assert again.ok
    assert again.succeeded == 2
    assert again.skipped == 0


def test_process_folder_output_dir(pdf_folder: Path, tmp_path: Path) -> None:
    out = tmp_path / "out"
    result = process_folder(
        pdf_folder, output_dir=out, options=ProcessOptions(compress=False)
    )
    assert result.ok
    assert (out / "+Exhibit A.pdf").is_file()
    assert not (pdf_folder / "+Exhibit A.pdf").exists()


def test_jobs_from_folder(pdf_folder: Path) -> None:
    jobs = jobs_from_folder(pdf_folder)
    assert len(jobs) == 2
    assert {j.label for j in jobs} == {"Exhibit A", "Exhibit B"}


def test_process_jobs_optimize_flag(sample_pdf: Path, tmp_path: Path) -> None:
    job = JobItem.from_path(sample_pdf)
    result = process_jobs(
        [job],
        output_dir=tmp_path,
        options=ProcessOptions(compress=False, optimize=True),
    )
    assert result.ok
    assert (tmp_path / f"+{sample_pdf.name}").is_file()


def test_process_jobs_cancel_after_first(pdf_folder: Path) -> None:
    jobs = jobs_from_folder(pdf_folder)
    assert len(jobs) >= 2
    calls = {"n": 0}

    def cancel_after_first() -> bool:
        # Allow first file to start/complete; cancel before second.
        calls["n"] += 1
        # Checked at the start of each iteration: 1st call → False, 2nd → True
        return calls["n"] > 1

    result = process_jobs(
        jobs,
        options=ProcessOptions(compress=False),
        cancel_check=cancel_after_first,
    )
    assert result.was_cancelled
    assert result.succeeded == 1
    assert result.cancelled == len(jobs) - 1
    assert (pdf_folder / "+Exhibit A.pdf").is_file()
    assert not (pdf_folder / "+Exhibit B.pdf").exists()
