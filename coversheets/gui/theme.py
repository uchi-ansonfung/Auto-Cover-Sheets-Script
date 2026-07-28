"""CustomTkinter appearance helpers."""

from __future__ import annotations

import customtkinter as ctk

APPEARANCE_MODES = ("System", "Light", "Dark")
DEFAULT_COLOR_THEME = "blue"


def normalize_appearance_mode(mode: str | None) -> str:
    """Return a valid CustomTkinter appearance mode name."""
    text = (mode or "System").strip().title()
    if text not in APPEARANCE_MODES:
        return "System"
    return text


def apply_appearance(mode: str | None = "System") -> str:
    """
    Apply global CustomTkinter appearance and color theme.

    Returns the normalized mode that was applied.
    """
    normalized = normalize_appearance_mode(mode)
    ctk.set_appearance_mode(normalized)
    ctk.set_default_color_theme(DEFAULT_COLOR_THEME)
    return normalized
