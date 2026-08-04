"""Processing options shared by CLI, GUI, and batch runner."""

from __future__ import annotations

from dataclasses import dataclass

from coversheets.cover import DEFAULT_VERTICAL_POSITION, normalize_vertical_position


@dataclass(frozen=True)
class ProcessOptions:
    """Flags that control how each PDF is written after the cover is added."""

    compress: bool = True
    force: bool = False
    dry_run: bool = False
    rename_to_label: bool = False
    # Document Info + XMP removal (not just "don't copy source fields").
    strip_metadata: bool = True
    # Optional post-process: OCR via ocrmypdf (requires extra install + tesseract).
    ocr: bool = False
    ocr_language: str = "eng"
    ocr_skip_text: bool = True
    # pypdf: compress_identical_objects (dedupe / orphan cleanup).
    optimize: bool = True
    # Web-optimized linearization via pikepdf or qpdf CLI.
    linearize: bool = False
    # Cover title vertical placement: "center" | "top_third".
    vertical_position: str = DEFAULT_VERTICAL_POSITION

    def describe(self) -> list[str]:
        """Human-readable enabled options for logs."""
        lines: list[str] = []
        position = normalize_vertical_position(self.vertical_position)
        if position == "top_third":
            lines.append("title position: top third")
        else:
            lines.append("title position: center")
        if self.compress:
            lines.append("compress page streams")
        if self.strip_metadata:
            lines.append("strip metadata (Info + XMP)")
        if self.optimize:
            lines.append("optimize (dedupe objects)")
        if self.linearize:
            lines.append("linearize")
        if self.ocr:
            mode = "skip pages with text" if self.ocr_skip_text else "force OCR"
            lines.append(f"OCR ({self.ocr_language}; {mode})")
        if self.rename_to_label:
            lines.append("name output after label")
        return lines
