"""User-facing strings and plain-language helpers for the GUI."""

from __future__ import annotations

import sys
from pathlib import Path

from coversheets import OUTPUT_PREFIX
from coversheets.options import ProcessOptions

# --- Short labels -----------------------------------------------------------

APP_SHORT_NAME = "Exhibit Cover Sheets"

STEP_LABELS = (
    "1 · Add PDFs",
    "2 · Review titles",
    "3 · Generate",
)

PRESET_LABELS = {
    "recommended": "Recommended",
    "searchable": "Searchable exhibits (OCR)",
    "custom": "Custom",
}

VERTICAL_POSITION_LABELS = {
    "center": "Center",
    "top_third": "Top third",
}

# --- Messages ---------------------------------------------------------------


def ocr_unavailable_message() -> str:
    """Explain missing OCR without pip jargon when running frozen."""
    if getattr(sys, "frozen", False):
        return (
            "Searchable text (OCR) isn’t available in this install.\n\n"
            "On Windows, install the full setup package:\n"
            "coversheets-…-windows-x64-setup.exe\n\n"
            "That build includes everything needed for OCR."
        )
    return (
        "Searchable text (OCR) isn’t available.\n\n"
        "Install the Windows full setup package, or for developers:\n"
        "pip install 'coversheets[ocr]' and install Tesseract."
    )


def linearize_unavailable_message() -> str:
    if getattr(sys, "frozen", False):
        return (
            "Faster web viewing isn’t available in this install.\n\n"
            "Use the full Windows setup package for this option."
        )
    return (
        "Faster web viewing needs an extra package.\n\n"
        "Install with: pip install 'coversheets[optimize]'"
    )


def empty_list_blurb(*, dnd_available: bool) -> str:
    base = (
        "We add a cover page to each PDF. Your originals are never changed.\n"
        "New files are named with a “+” so they’re easy to find."
    )
    if dnd_available:
        return base + "\nYou can also drag PDFs or a folder onto this window."
    return base


def output_example(
    *,
    mode: str,
    folder: str | None,
    rename_to_label: bool,
    sample_name: str = "Contract.pdf",
    sample_label: str = "Contract",
) -> str:
    """One-line explanation of where outputs go."""
    if rename_to_label:
        out_name = f"{OUTPUT_PREFIX}{sample_label}.pdf"
    else:
        stem = Path(sample_name).stem
        out_name = f"{OUTPUT_PREFIX}{stem}.pdf"
    if mode == "folder" and folder:
        return f"Saves as {out_name} into:\n{folder}"
    if mode == "folder":
        return f"Saves as {out_name} into your chosen folder (pick one below)."
    return f"Saves as {out_name} next to each original PDF."


def status_for_jobs(total: int, included: int) -> str:
    if total == 0:
        return "Add PDFs or open a folder to begin."
    if included == 0:
        return f"{total} PDF(s) loaded · none selected to process"
    return f"{total} PDF(s) loaded · {included} will get cover sheets"


def plain_option_lines(options: ProcessOptions) -> list[str]:
    """Short plain-English lines for the progress log."""
    lines: list[str] = []
    position = (options.vertical_position or "center").strip().lower()
    if position == "top_third":
        lines.append("Title position: top third")
    else:
        lines.append("Title position: center")
    if options.strip_metadata:
        lines.append("Remove hidden document info")
    if options.ocr:
        lines.append(f"Make text searchable (OCR, {options.ocr_language})")
    if options.compress:
        lines.append("Compress PDF data")
    if options.optimize:
        lines.append("Shrink duplicate PDF data")
    if options.linearize:
        lines.append("Faster web viewing")
    if options.rename_to_label:
        lines.append("Name output after cover title")
    if options.force:
        lines.append("Replace existing +files")
    return lines


def truncate_middle(text: str, max_len: int = 42) -> str:
    """Truncate a long path for list display."""
    text = text or ""
    if max_len < 8 or len(text) <= max_len:
        return text
    keep = max_len - 1
    left = keep // 2
    right = keep - left
    return text[:left] + "…" + text[-right:]


def preview_target_index(
    selected: set[int],
    *,
    anchor: int | None,
    job_count: int,
) -> int | None:
    """Which job index the cover preview should show."""
    if job_count <= 0:
        return None
    if anchor is not None and 0 <= anchor < job_count:
        return anchor
    if selected:
        return max(i for i in selected if 0 <= i < job_count)
    return 0 if job_count == 1 else None


def done_headline(result) -> tuple[str, str]:
    """Return (window title, headline) for the completion dialog."""
    ok = result.ok and result.failed == 0 and not result.was_cancelled
    if result.was_cancelled:
        return "Cancelled", "Stopped before everything finished."
    if ok and result.succeeded == 0 and result.skipped:
        return "Nothing new written", "Those files were skipped (outputs already exist)."
    if ok:
        return (
            "Done",
            f"Cover sheets are ready. New PDFs start with “{OUTPUT_PREFIX}” "
            "so they’re easy to spot.",
        )
    return "Finished with problems", "Some files need attention."
