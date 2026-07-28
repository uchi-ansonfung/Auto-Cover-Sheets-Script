"""Persistent GUI preferences (JSON on disk)."""

from __future__ import annotations

import json
import os
import sys
from dataclasses import asdict, dataclass, fields
from pathlib import Path
from typing import Any


PREFS_VERSION = 1


@dataclass
class AppPreferences:
    """User settings restored across GUI sessions."""

    version: int = PREFS_VERSION
    # Last input locations
    last_folder: str = ""
    last_file_dialog_dir: str = ""
    # Output
    output_dir: str = ""
    # Window geometry e.g. "900x560+120+80"
    window_geometry: str = "900x560"
    # Toggles (defaults match ProcessOptions / GUI defaults)
    compress: bool = True
    force: bool = False
    rename_to_label: bool = False
    strip_metadata: bool = True
    optimize: bool = True
    linearize: bool = False
    ocr: bool = False
    ocr_language: str = "eng"
    open_when_done: bool = True

    def resolved_last_folder(self) -> Path | None:
        return _existing_dir(self.last_folder)

    def resolved_file_dialog_dir(self) -> Path | None:
        return _existing_dir(self.last_file_dialog_dir) or self.resolved_last_folder()

    def resolved_output_dir(self) -> Path | None:
        return _existing_dir(self.output_dir)


def _existing_dir(raw: str) -> Path | None:
    text = (raw or "").strip()
    if not text:
        return None
    path = Path(text).expanduser()
    try:
        path = path.resolve()
    except OSError:
        return None
    return path if path.is_dir() else None


def default_prefs_path() -> Path:
    """
    Platform config path for preferences.

    - macOS: ~/Library/Application Support/coversheets/prefs.json
    - Windows: %APPDATA%/coversheets/prefs.json
    - Linux/other: $XDG_CONFIG_HOME/coversheets/prefs.json or ~/.config/...
    """
    if sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support" / "coversheets"
    elif sys.platform == "win32":
        appdata = os.environ.get("APPDATA")
        base = Path(appdata) / "coversheets" if appdata else Path.home() / "coversheets"
    else:
        xdg = os.environ.get("XDG_CONFIG_HOME")
        base = Path(xdg) / "coversheets" if xdg else Path.home() / ".config" / "coversheets"
    return base / "prefs.json"


def load_preferences(path: Path | None = None) -> AppPreferences:
    """Load preferences from disk; return defaults if missing or invalid."""
    prefs_path = path if path is not None else default_prefs_path()
    if not prefs_path.is_file():
        return AppPreferences()
    try:
        raw = json.loads(prefs_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        return AppPreferences()
    if not isinstance(raw, dict):
        return AppPreferences()
    return preferences_from_dict(raw)


def preferences_from_dict(raw: dict[str, Any]) -> AppPreferences:
    """Build preferences from a JSON object, ignoring unknown keys."""
    known = {f.name for f in fields(AppPreferences)}
    data: dict[str, Any] = {}
    for key, value in raw.items():
        if key not in known:
            continue
        data[key] = value
    prefs = AppPreferences()
    for f in fields(AppPreferences):
        if f.name not in data:
            continue
        value = data[f.name]
        expected = type(getattr(prefs, f.name))
        try:
            if expected is bool:
                if isinstance(value, bool):
                    setattr(prefs, f.name, value)
                elif isinstance(value, (int, float)) and value in (0, 1):
                    setattr(prefs, f.name, bool(value))
                elif isinstance(value, str) and value.lower() in {
                    "true",
                    "false",
                    "1",
                    "0",
                    "yes",
                    "no",
                }:
                    setattr(prefs, f.name, value.lower() in {"true", "1", "yes"})
            elif expected is int:
                setattr(prefs, f.name, int(value))
            elif expected is str:
                setattr(prefs, f.name, str(value) if value is not None else "")
            else:
                setattr(prefs, f.name, value)
        except (TypeError, ValueError):
            continue
    prefs.version = PREFS_VERSION
    return prefs


def save_preferences(prefs: AppPreferences, path: Path | None = None) -> Path:
    """Write preferences atomically; return the path written."""
    prefs_path = path if path is not None else default_prefs_path()
    prefs_path.parent.mkdir(parents=True, exist_ok=True)
    payload = asdict(prefs)
    payload["version"] = PREFS_VERSION
    text = json.dumps(payload, indent=2, sort_keys=True) + "\n"
    tmp = prefs_path.with_suffix(prefs_path.suffix + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    tmp.replace(prefs_path)
    return prefs_path.resolve()
