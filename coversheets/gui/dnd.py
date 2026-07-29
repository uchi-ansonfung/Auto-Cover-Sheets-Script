"""Optional drag-and-drop helpers (tkinterdnd2 when available)."""

from __future__ import annotations

import sys
from collections.abc import Callable, Sequence
from pathlib import Path
from typing import Any

_DND_READY = False
_IMPORT_ERROR: str | None = None

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore

    _DND_READY = True
except Exception as exc:  # pragma: no cover - environment dependent
    DND_FILES = "DND_FILES"  # type: ignore[misc, assignment]
    TkinterDnD = None  # type: ignore[misc, assignment]
    _IMPORT_ERROR = str(exc)


def dnd_available() -> bool:
    """True if tkinterdnd2 imported successfully."""
    return _DND_READY


def dnd_import_error() -> str | None:
    return _IMPORT_ERROR


def make_dnd_root(ctk_app_class: type) -> type:
    """
    Return a class that mixes TkinterDnD.DnDWrapper into a CTk root when possible.

    CustomTkinter's CTk already subclasses Tk; we add DnD methods at runtime
    for the instance when available.
    """
    if not _DND_READY or TkinterDnD is None:
        return ctk_app_class

    # Prefer subclassing both when possible.
    try:

        class DnDApp(ctk_app_class, TkinterDnD.DnDWrapper):  # type: ignore[misc, valid-type]
            def __init__(self, *args: Any, **kwargs: Any) -> None:
                super().__init__(*args, **kwargs)
                # Register this widget as a DnD-aware Tk window.
                self.TkdndVersion = TkinterDnD._require(self)  # type: ignore[attr-defined]

        return DnDApp
    except Exception:
        return ctk_app_class


def parse_drop_paths(data: str) -> list[Path]:
    """
    Parse a Tk DND_FILES payload into paths.

    Handles brace-wrapped paths with spaces: ``{C:/My Files/a.pdf} b.pdf``.
    """
    if not data:
        return []
    raw = data.strip()
    paths: list[str] = []
    token = ""
    in_brace = False
    for ch in raw:
        if ch == "{":
            in_brace = True
            token = ""
            continue
        if ch == "}":
            in_brace = False
            if token:
                paths.append(token)
            token = ""
            continue
        if ch in (" ", "\n", "\r", "\t") and not in_brace:
            if token:
                paths.append(token)
                token = ""
            continue
        token += ch
    if token:
        paths.append(token)

    result: list[Path] = []
    for item in paths:
        text = item.strip().strip('"')
        if not text:
            continue
        # Some platforms prefix file://
        if text.startswith("file://"):
            text = text[7:]
            if sys.platform == "win32" and text.startswith("/"):
                # file:///C:/...
                text = text.lstrip("/")
        result.append(Path(text))
    return result


def register_drop_target(
    widget: Any,
    on_paths: Callable[[Sequence[Path]], None],
) -> bool:
    """
    Register *widget* as a file drop target.

    Returns True if registration succeeded.
    """
    if not _DND_READY:
        return False
    try:
        widget.drop_target_register(DND_FILES)

        def _handler(event: Any) -> None:
            paths = parse_drop_paths(getattr(event, "data", "") or "")
            if paths:
                on_paths(paths)

        widget.dnd_bind("<<Drop>>", _handler)
        return True
    except Exception:
        return False
