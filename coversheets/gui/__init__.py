"""GUI package: CustomTkinter list UI for exhibit cover sheets.

Heavy Toolkit imports (CustomTkinter / Tk) load only when you import
``coversheets.gui.app`` or call :func:`run_app`, so pure helpers under
``coversheets.gui.copy`` and ``coversheets.gui.dnd`` stay testable without a
display / ``_tkinter``.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from coversheets.gui.app import CoverSheetsApp

__all__ = ["CoverSheetsApp", "run_app"]


def run_app(initial_folder=None):
    """Launch the GUI and block until the window closes."""
    from coversheets.gui.app import run_app as _run_app

    return _run_app(initial_folder=initial_folder)


def __getattr__(name: str):
    if name == "CoverSheetsApp":
        from coversheets.gui.app import CoverSheetsApp

        return CoverSheetsApp
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
