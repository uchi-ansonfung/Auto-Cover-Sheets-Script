"""Merge cover sheets with original PDFs."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import BinaryIO

from pypdf import PdfReader, PdfWriter

from coversheets.options import ProcessOptions
from coversheets.pdf_ops import (
    linearize_pdf,
    optimize_writer,
    run_ocr,
    strip_metadata_file,
    strip_metadata_writer,
)

PARTIAL_SUFFIX = ".partial"


def partial_path_for(output: Path) -> Path:
    """Temporary path used for atomic writes (final name + ``.partial``)."""
    return Path(str(output) + PARTIAL_SUFFIX)


def add_cover_to_pdf(
    cover: BytesIO | BinaryIO | Path | str,
    original: Path | str,
    output: Path | str,
    *,
    compress: bool = True,
    strip_metadata: bool = True,
    optimize: bool = True,
    ocr: bool = False,
    ocr_language: str = "eng",
    ocr_skip_text: bool = True,
    linearize: bool = False,
    options: ProcessOptions | None = None,
) -> Path:
    """
    Prepend a cover sheet to ``original`` and write the result to ``output``.

    Writes to a ``.partial`` temp file first, then renames into place so a
    crash mid-write cannot leave a truncated final PDF. Optional post-steps
    (OCR, linearize, metadata strip) run on the temp file before the rename.

    Returns the resolved output path.
    """
    if options is not None:
        compress = options.compress
        strip_metadata = options.strip_metadata
        optimize = options.optimize
        ocr = options.ocr
        ocr_language = options.ocr_language
        ocr_skip_text = options.ocr_skip_text
        linearize = options.linearize

    original_path = Path(original)
    output_path = Path(output)
    tmp_path = partial_path_for(output_path)

    if isinstance(cover, (str, Path)):
        cover_reader = PdfReader(str(cover))
    else:
        cover.seek(0)
        cover_reader = PdfReader(cover)

    original_reader = PdfReader(str(original_path))
    writer = PdfWriter()

    if not cover_reader.pages:
        raise ValueError("Cover PDF has no pages")

    writer.add_page(cover_reader.pages[0])

    for page in original_reader.pages:
        # Add first: compress_content_streams requires a PdfWriter-owned page.
        # Calling it on a PdfReader page raises ValueError ("Page must be part
        # of a PdfWriter") and can surface AttributeError on ContentStream.
        writer.add_page(page)
        if compress:
            # Lossless; can be CPU-intensive on large/complex pages.
            writer.pages[-1].compress_content_streams()

    if optimize:
        optimize_writer(writer)

    if strip_metadata:
        strip_metadata_writer(writer)
    else:
        # Still avoid copying source Info; leave a minimal writer default.
        writer.add_metadata({})

    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Remove any leftover partial from a previous crash.
    if tmp_path.exists():
        tmp_path.unlink()

    try:
        with tmp_path.open("wb") as fh:
            writer.write(fh)

        # Post-write steps may reintroduce producer/XMP (especially OCR).
        if ocr:
            run_ocr(
                tmp_path,
                language=ocr_language,
                skip_text=ocr_skip_text,
            )
            if strip_metadata:
                strip_metadata_file(tmp_path)

        if linearize:
            linearize_pdf(tmp_path)
            if strip_metadata:
                strip_metadata_file(tmp_path)

        # Atomic replace of the final destination.
        tmp_path.replace(output_path)
    except Exception:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass
        raise

    return output_path.resolve()
