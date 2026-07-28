"""PDF post-processing: metadata strip, optimize, OCR, linearize."""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from pypdf import PdfReader, PdfWriter

from coversheets.bundled_tools import configure_bundled_tools, tesseract_on_path


class DependencyError(RuntimeError):
    """Raised when an optional tool/package is required but missing."""


def _closed_temp_pdf(prefix: str) -> Path:
    """
    Create an empty temp PDF path with the handle closed.

    ``mkstemp`` returns an open FD; leaving it open breaks replace/unlink on
    Windows (WinError 32). Always close before returning the path.
    """
    fd, name = tempfile.mkstemp(prefix=prefix, suffix=".pdf")
    os.close(fd)
    return Path(name)


def strip_metadata_writer(writer: PdfWriter) -> None:
    """
    Remove document-level metadata from a PdfWriter before save.

    Clears the Info dictionary (including Producer) and any XMP packet
    attached to the writer. Does not attempt to scrub hidden data inside
    page content streams.
    """
    writer.metadata = None
    try:
        writer.xmp_metadata = None
    except (TypeError, AttributeError, ValueError):
        pass


def optimize_writer(writer: PdfWriter) -> None:
    """Dedupe identical objects and drop orphans (lossless structure optimize)."""
    writer.compress_identical_objects(remove_identicals=True, remove_orphans=True)


def strip_metadata_file(path: Path | str) -> Path:
    """
    Rewrite ``path`` with document Info/XMP removed.

    Uses pikepdf when available (more thorough Root.Metadata removal);
    falls back to a pypdf rewrite.
    """
    pdf_path = Path(path)
    if _try_strip_with_pikepdf(pdf_path):
        return pdf_path.resolve()
    return _strip_with_pypdf(pdf_path)


def _strip_with_pypdf(pdf_path: Path) -> Path:
    reader = PdfReader(str(pdf_path))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    # Do not copy reader metadata / attachments intentionally.
    strip_metadata_writer(writer)
    tmp = pdf_path.with_suffix(pdf_path.suffix + ".tmp")
    try:
        with tmp.open("wb") as fh:
            writer.write(fh)
        tmp.replace(pdf_path)
    finally:
        if tmp.exists():
            tmp.unlink(missing_ok=True)
    return pdf_path.resolve()


def _try_strip_with_pikepdf(pdf_path: Path) -> bool:
    try:
        import pikepdf
    except ImportError:
        return False

    with pikepdf.open(pdf_path, allow_overwriting_input=True) as pdf:
        # Clear document info keys.
        with pdf.open_metadata(set_pikepdf_as_editor=False) as meta:
            meta.clear()
        # Remove XMP stream on the catalog if present.
        try:
            del pdf.Root.Metadata
        except KeyError:
            pass
        # Empty classic Info dict.
        try:
            for key in list(pdf.docinfo.keys()):
                del pdf.docinfo[key]
        except Exception:
            pdf.docinfo = pdf.make_indirect(pikepdf.Dictionary())
        pdf.save(pdf_path)
    return True


def ocr_available() -> bool:
    """
    True if OCR can run: ocrmypdf (package or CLI) plus a Tesseract binary.

    Activates tools bundled next to a frozen/installed binary first.
    """
    configure_bundled_tools()
    try:
        import ocrmypdf  # noqa: F401

        has_ocrmypdf = True
    except ImportError:
        has_ocrmypdf = shutil.which("ocrmypdf") is not None
    if not has_ocrmypdf:
        return False
    # Package alone is not enough; ocrmypdf shells out to tesseract.
    return tesseract_on_path()


def linearize_available() -> bool:
    """True if pikepdf is importable or the qpdf binary is on PATH."""
    configure_bundled_tools()
    try:
        import pikepdf  # noqa: F401

        return True
    except ImportError:
        return shutil.which("qpdf") is not None


def run_ocr(
    path: Path | str,
    *,
    language: str = "eng",
    skip_text: bool = True,
) -> Path:
    """
    OCR ``path`` in place via ocrmypdf.

    Requires the ``ocrmypdf`` package (and system Tesseract). Pages that
    already contain text are skipped when ``skip_text`` is True.
    """
    pdf_path = Path(path)
    configure_bundled_tools()
    if not ocr_available():
        raise DependencyError(
            "OCR requires ocrmypdf and Tesseract. "
            "Install with: pip install 'coversheets[ocr]' and system Tesseract, "
            "or use the Windows full installer (bundles both)."
        )

    language = (language or "eng").strip() or "eng"
    tmp_path = _closed_temp_pdf("coversheets_ocr_")
    try:
        try:
            import ocrmypdf

            ocrmypdf.ocr(
                str(pdf_path),
                str(tmp_path),
                language=language,
                skip_text=skip_text,
                force_ocr=not skip_text,
                progress_bar=False,
                # Keep layout; do not deskew aggressively by default.
                optimize=0,
            )
        except ImportError:
            cmd = [
                "ocrmypdf",
                "--language",
                language,
                "--optimize",
                "0",
                "--quiet",
            ]
            if skip_text:
                cmd.append("--skip-text")
            else:
                cmd.append("--force-ocr")
            cmd.extend([str(pdf_path), str(tmp_path)])
            proc = subprocess.run(cmd, capture_output=True, text=True, check=False)
            if proc.returncode != 0:
                detail = (proc.stderr or proc.stdout or "").strip()
                raise RuntimeError(
                    f"ocrmypdf failed (exit {proc.returncode}): {detail or 'unknown error'}"
                ) from None

        tmp_path.replace(pdf_path)
    finally:
        if tmp_path.exists():
            tmp_path.unlink(missing_ok=True)
    return pdf_path.resolve()


def linearize_pdf(path: Path | str) -> Path:
    """
    Linearize (web-optimize) ``path`` in place.

    Prefers pikepdf; falls back to the ``qpdf`` CLI.
    """
    pdf_path = Path(path)
    if not linearize_available():
        raise DependencyError(
            "Linearize requires pikepdf or the qpdf CLI. Install with: "
            "pip install 'coversheets[optimize]' (or install qpdf)."
        )

    if _try_linearize_with_pikepdf(pdf_path):
        return pdf_path.resolve()
    return _linearize_with_qpdf(pdf_path)


def _try_linearize_with_pikepdf(pdf_path: Path) -> bool:
    try:
        import pikepdf
    except ImportError:
        return False

    with pikepdf.open(pdf_path, allow_overwriting_input=True) as pdf:
        pdf.save(pdf_path, linearize=True)
    return True


def _linearize_with_qpdf(pdf_path: Path) -> Path:
    qpdf = shutil.which("qpdf")
    if not qpdf:
        raise DependencyError("qpdf not found on PATH")
    # qpdf cannot always --replace-input on all platforms; use temp then replace.
    tmp_path = _closed_temp_pdf("coversheets_lin_")
    try:
        proc = subprocess.run(
            [qpdf, "--linearize", str(pdf_path), str(tmp_path)],
            capture_output=True,
            text=True,
            check=False,
        )
        if proc.returncode != 0:
            detail = (proc.stderr or proc.stdout or "").strip()
            raise RuntimeError(
                f"qpdf linearize failed (exit {proc.returncode}): {detail or 'unknown error'}"
            )
        tmp_path.replace(pdf_path)
    finally:
        if tmp_path.exists():
            tmp_path.unlink(missing_ok=True)
    return pdf_path.resolve()
