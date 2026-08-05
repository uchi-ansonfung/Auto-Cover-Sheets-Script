"""Batch processing of PDF jobs (paths + cover labels)."""

from __future__ import annotations

import re
import sys
import traceback
from collections.abc import Callable, Iterable, Sequence
from dataclasses import dataclass, field
from pathlib import Path

from coversheets import OUTPUT_PREFIX
from coversheets.cover import cover_label_from_filename, create_cover_sheet
from coversheets.merge import add_cover_to_pdf
from coversheets.options import ProcessOptions
from coversheets.pdf_ops import linearize_available, ocr_available

ProgressCallback = Callable[[int, int, str], None]
CancelCheck = Callable[[], bool]

# Characters illegal or awkward in file names on Windows/macOS/Linux.
_UNSAFE_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


@dataclass
class JobItem:
    """One PDF to process with an editable cover label."""

    source: Path
    label: str
    include: bool = True

    @classmethod
    def from_path(cls, path: Path | str, *, label: str | None = None) -> JobItem:
        source = Path(path).expanduser().resolve()
        return cls(
            source=source,
            label=label if label is not None else cover_label_from_filename(source),
            include=True,
        )


@dataclass(frozen=True)
class BatchResult:
    """Summary of a processing run."""

    total: int = 0
    succeeded: int = 0
    skipped: int = 0
    failed: int = 0
    cancelled: int = 0
    errors: tuple[tuple[str, str], ...] = ()
    dry_run: bool = False
    was_cancelled: bool = False

    @property
    def ok(self) -> bool:
        return self.failed == 0


@dataclass
class _Counters:
    succeeded: int = 0
    skipped: int = 0
    failed: int = 0
    cancelled: int = 0
    errors: list[tuple[str, str]] = field(default_factory=list)


def is_output_filename(filename: str) -> bool:
    """Return True if this looks like a previous run's output (+Name.pdf)."""
    return filename.startswith(OUTPUT_PREFIX)


def sanitize_filename_stem(label: str) -> str:
    """
    Turn a cover label into a safe file-name stem (no extension).

    Strips path separators and reserved characters; collapses whitespace.
    Falls back to ``untitled`` if nothing usable remains.
    """
    text = label.strip()
    text = _UNSAFE_FILENAME_CHARS.sub("", text)
    text = re.sub(r"\s+", " ", text).strip()
    # Windows does not like trailing dots/spaces in basenames.
    text = text.rstrip(". ")
    # Avoid reserved device names on Windows when used as the whole stem.
    if text.casefold() in {
        "con",
        "prn",
        "aux",
        "nul",
        *(f"com{i}" for i in range(1, 10)),
        *(f"lpt{i}" for i in range(1, 10)),
    }:
        text = f"{text}_file"
    return text or "untitled"


def output_path_for(
    source: Path | str,
    output_dir: Path | None = None,
    *,
    label: str | None = None,
    rename_to_label: bool = False,
    used_basenames: set[str] | None = None,
) -> Path:
    """
    Destination path for a source PDF.

    By default writes ``+OriginalName.pdf`` next to the source (or in
    ``output_dir``). With ``rename_to_label=True``, writes
    ``+{sanitized label}.pdf`` instead. Duplicate basenames in one run get
    `` (2)``, `` (3)``, … suffixes.
    """
    source_path = Path(source)
    parent = output_dir if output_dir is not None else source_path.parent

    if rename_to_label:
        stem = sanitize_filename_stem(label if label is not None else source_path.stem)
        basename = f"{OUTPUT_PREFIX}{stem}.pdf"
    else:
        basename = f"{OUTPUT_PREFIX}{source_path.name}"

    if used_basenames is not None:
        key = basename.casefold()
        if key in used_basenames:
            # Disambiguate collisions within the same batch.
            if rename_to_label:
                stem = sanitize_filename_stem(
                    label if label is not None else source_path.stem
                )
                n = 2
                while f"{OUTPUT_PREFIX}{stem} ({n}).pdf".casefold() in used_basenames:
                    n += 1
                basename = f"{OUTPUT_PREFIX}{stem} ({n}).pdf"
            else:
                # Same source name from different folders into one output dir.
                stem = source_path.stem
                suffix = source_path.suffix  # typically .pdf
                n = 2
                while (
                    f"{OUTPUT_PREFIX}{stem} ({n}){suffix}".casefold() in used_basenames
                ):
                    n += 1
                basename = f"{OUTPUT_PREFIX}{stem} ({n}){suffix}"
        used_basenames.add(basename.casefold())

    return parent / basename


def collect_pdf_paths(folder: Path | str) -> list[Path]:
    """List PDF paths in folder, sorted, excluding prior +outputs."""
    folder_path = Path(folder)
    paths = [
        path
        for path in folder_path.iterdir()
        if path.is_file()
        and path.suffix.lower() == ".pdf"
        and not is_output_filename(path.name)
    ]
    paths.sort(key=lambda p: p.name.casefold())
    return paths


def collect_pdfs(folder: Path | str) -> list[str]:
    """List PDF basenames in folder (legacy helper)."""
    return [path.name for path in collect_pdf_paths(folder)]


def jobs_from_folder(folder: Path | str) -> list[JobItem]:
    """Build job items for every input PDF in a folder."""
    return [JobItem.from_path(path) for path in collect_pdf_paths(folder)]


def jobs_from_paths(paths: Iterable[Path | str]) -> list[JobItem]:
    """Build job items from an explicit list of PDF paths."""
    jobs: list[JobItem] = []
    seen: set[Path] = set()
    for raw in paths:
        path = Path(raw).expanduser().resolve()
        if path in seen:
            continue
        if not path.is_file() or path.suffix.lower() != ".pdf":
            continue
        if is_output_filename(path.name):
            continue
        seen.add(path)
        jobs.append(JobItem.from_path(path))
    jobs.sort(key=lambda j: j.source.name.casefold())
    return jobs


def validate_options(options: ProcessOptions) -> None:
    """Raise ValueError if enabled options cannot run in this environment."""
    if options.ocr and not ocr_available():
        raise ValueError(
            "OCR is enabled but ocrmypdf/pypdfium2/Tesseract is not available. "
            "Use the Windows full installer, or: pip install 'coversheets[ocr]' "
            "and install Tesseract."
        )
    if options.linearize and not linearize_available():
        raise ValueError(
            "Linearize is enabled but neither pikepdf nor qpdf is available. "
            "Install with: pip install 'coversheets[optimize]' (or install qpdf)."
        )


def process_jobs(
    jobs: Sequence[JobItem],
    *,
    output_dir: Path | str | None = None,
    options: ProcessOptions | None = None,
    compress: bool | None = None,
    force: bool | None = None,
    dry_run: bool | None = None,
    rename_to_label: bool | None = None,
    progress: ProgressCallback | None = None,
    log: Callable[[str], None] | None = None,
    cancel_check: CancelCheck | None = None,
) -> BatchResult:
    """
    Process included jobs.

    Prefer passing a :class:`ProcessOptions` instance. Individual keyword
    arguments still work and override matching fields on ``options``.

    ``cancel_check`` is polled between files; when it returns True, the
    current file is not started and remaining jobs are counted as cancelled.
    """
    base = options or ProcessOptions()
    opts = ProcessOptions(
        compress=base.compress if compress is None else compress,
        force=base.force if force is None else force,
        dry_run=base.dry_run if dry_run is None else dry_run,
        rename_to_label=base.rename_to_label
        if rename_to_label is None
        else rename_to_label,
        strip_metadata=base.strip_metadata,
        ocr=base.ocr,
        ocr_language=base.ocr_language,
        ocr_skip_text=base.ocr_skip_text,
        optimize=base.optimize,
        linearize=base.linearize,
        vertical_position=base.vertical_position,
    )

    write = log if log is not None else print
    out_root = Path(output_dir).resolve() if output_dir else None
    selected = [job for job in jobs if job.include]
    counters = _Counters()
    used_basenames: set[str] = set()
    was_cancelled = False

    if not selected:
        write("No jobs selected.")
        return BatchResult(dry_run=opts.dry_run)

    if not opts.dry_run:
        try:
            validate_options(opts)
        except ValueError as exc:
            write(f"ERROR: {exc}")
            return BatchResult(
                total=len(selected),
                failed=len(selected),
                errors=(("<options>", str(exc)),),
                dry_run=False,
            )

    total = len(selected)
    write(f"Processing {total} PDF(s)")
    if out_root is not None:
        write(f"Output directory: {out_root}")
    for line in opts.describe():
        write(f"Option: {line}")
    if opts.dry_run:
        write("Dry run — no files will be written.")
    write("")

    for index, job in enumerate(selected, start=1):
        if cancel_check is not None and cancel_check():
            remaining = total - index + 1
            counters.cancelled += remaining
            was_cancelled = True
            write(
                f"Cancelled after completing {index - 1} of {total} "
                f"({remaining} left unprocessed)."
            )
            break

        source = job.source
        dest = output_path_for(
            source,
            out_root,
            label=job.label,
            rename_to_label=opts.rename_to_label,
            used_basenames=used_basenames,
        )
        display = source.name
        prefix = f"[{index}/{total}]"
        if progress:
            progress(index, total, display)

        if not source.is_file():
            counters.failed += 1
            msg = f"File not found: {source}"
            counters.errors.append((display, msg))
            write(f"{prefix} ERROR {display}: {msg}")
            continue

        if dest.exists() and not opts.force:
            write(
                f"{prefix} Skip {display} "
                f"(output exists: {dest.name}; use force to overwrite)"
            )
            counters.skipped += 1
            continue

        if opts.dry_run:
            action = "overwrite" if dest.exists() else "write"
            write(f"{prefix} Would {action} {dest}  (label={job.label!r})")
            counters.succeeded += 1
            continue

        write(f"{prefix} Processing {display}  (label={job.label!r})...")
        try:
            cover_buffer = create_cover_sheet(
                job.label,
                vertical_position=opts.vertical_position,
            )
            add_cover_to_pdf(
                cover_buffer,
                source,
                dest,
                options=opts,
            )
            write(f"{prefix} Saved {dest}")
            counters.succeeded += 1
        except Exception as exc:
            counters.failed += 1
            counters.errors.append((display, str(exc)))
            write(f"{prefix} ERROR {display}: {exc}")
            traceback.print_exc(file=sys.stderr)

    write("")
    status = "Cancelled" if was_cancelled else "Done"
    write(
        f"{status}. processed={counters.succeeded} skipped={counters.skipped} "
        f"failed={counters.failed} cancelled={counters.cancelled} total={total}"
        + (" (dry run)" if opts.dry_run else "")
    )
    if counters.errors:
        write("Failures:")
        for name, msg in counters.errors:
            write(f"  - {name}: {msg}")

    return BatchResult(
        total=total,
        succeeded=counters.succeeded,
        skipped=counters.skipped,
        failed=counters.failed,
        cancelled=counters.cancelled,
        errors=tuple(counters.errors),
        dry_run=opts.dry_run,
        was_cancelled=was_cancelled,
    )


def process_folder(
    folder: Path | str,
    *,
    output_dir: Path | str | None = None,
    options: ProcessOptions | None = None,
    progress: ProgressCallback | None = None,
) -> BatchResult:
    """
    Process all input PDFs in ``folder`` using filename stems as labels.

    Convenience wrapper around :func:`process_jobs` for headless/CLI use.
    """
    folder_path = Path(folder).resolve()
    jobs = jobs_from_folder(folder_path)
    opts = options or ProcessOptions()
    if not jobs:
        print("No input PDFs found (skipping files that already start with '+').")
        return BatchResult(dry_run=opts.dry_run)

    print(f"Found {len(jobs)} PDF(s) in {folder_path}")
    return process_jobs(
        jobs,
        output_dir=output_dir,
        options=opts,
        progress=progress,
    )
