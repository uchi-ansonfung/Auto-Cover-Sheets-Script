"""Shared helpers (no GUI toolkit dependency)."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path
from typing import Sequence

from coversheets.process import BatchResult, JobItem


def resolve_app_asset(name: str) -> Path | None:
    """
    Locate a packaged asset under assets/ for dev installs and frozen builds.

    Returns None when the file is missing (callers should fail soft).
    """
    candidates: list[Path] = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "assets" / name)
    # coversheets/util.py → repo root is parents[1] when running from source tree.
    package_root = Path(__file__).resolve().parent
    candidates.append(package_root.parent / "assets" / name)
    # Fallback: cwd-relative (e.g. running from repo root in odd layouts).
    candidates.append(Path.cwd() / "assets" / name)
    for path in candidates:
        try:
            if path.is_file():
                return path
        except OSError:
            continue
    return None


def open_in_file_manager(path: Path) -> None:
    """Open a folder (or its parent if a file) in the OS file manager."""
    target = path.expanduser().resolve()
    if target.is_file():
        target = target.parent
    if not target.is_dir():
        raise FileNotFoundError(f"Not a directory: {target}")

    if sys.platform == "darwin":
        subprocess.run(["open", str(target)], check=False)
    elif sys.platform == "win32":
        os.startfile(str(target))  # type: ignore[attr-defined]
    else:
        subprocess.run(["xdg-open", str(target)], check=False)


def resolve_result_folders(
    jobs: Sequence[JobItem],
    output_dir: Path | None,
) -> list[Path]:
    """
    Folders that contain outputs for this run.

    Prefer the explicit output directory; otherwise unique source parents of
    included jobs (sorted).
    """
    if output_dir is not None:
        return [output_dir.expanduser().resolve()]
    parents = sorted(
        {job.source.parent.resolve() for job in jobs if job.include},
        key=lambda p: str(p).casefold(),
    )
    return parents


def format_result_summary(result: BatchResult) -> str:
    """Human-readable one-line summary of a batch run."""
    parts = [
        f"Processed {result.succeeded} of {result.total}",
        f"skipped {result.skipped}",
        f"failed {result.failed}",
    ]
    if result.cancelled or result.was_cancelled:
        parts.append(f"cancelled {result.cancelled}")
    text = "  ·  ".join(parts)
    if result.was_cancelled:
        text = f"Cancelled — {text}"
    return text
