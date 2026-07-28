"""Tests for preference load/save (no GUI)."""

from __future__ import annotations

from pathlib import Path

from coversheets.prefs import (
    AppPreferences,
    load_preferences,
    preferences_from_dict,
    save_preferences,
)


def test_preferences_from_dict_ignores_unknown_and_bad_types() -> None:
    prefs = preferences_from_dict(
        {
            "compress": False,
            "ocr_language": "deu",
            "unknown_key": 123,
            "optimize": "not-a-bool",  # invalid → keep default True
            "version": 99,
        }
    )
    assert prefs.compress is False
    assert prefs.ocr_language == "deu"
    assert prefs.version == 1  # rewritten to current
    assert prefs.optimize is True
    # Unknown key dropped; defaults remain for unset fields
    assert prefs.force is False


def test_save_and_load_roundtrip(tmp_path: Path) -> None:
    path = tmp_path / "prefs.json"
    original = AppPreferences(
        last_folder=str(tmp_path / "in"),
        last_file_dialog_dir=str(tmp_path / "in"),
        output_dir=str(tmp_path / "out"),
        window_geometry="1000x700+10+20",
        appearance_mode="Dark",
        compress=False,
        force=True,
        rename_to_label=True,
        strip_metadata=False,
        optimize=False,
        linearize=True,
        ocr=True,
        ocr_language="eng+spa",
        open_when_done=False,
    )
    (tmp_path / "in").mkdir()
    (tmp_path / "out").mkdir()

    save_preferences(original, path)
    loaded = load_preferences(path)

    assert loaded.last_folder == original.last_folder
    assert loaded.output_dir == original.output_dir
    assert loaded.window_geometry == original.window_geometry
    assert loaded.appearance_mode == "Dark"
    assert loaded.compress is False
    assert loaded.force is True
    assert loaded.rename_to_label is True
    assert loaded.strip_metadata is False
    assert loaded.optimize is False
    assert loaded.linearize is True
    assert loaded.ocr is True
    assert loaded.ocr_language == "eng+spa"
    assert loaded.open_when_done is False
    assert loaded.resolved_last_folder() == (tmp_path / "in").resolve()
    assert loaded.resolved_output_dir() == (tmp_path / "out").resolve()


def test_load_missing_file_returns_defaults(tmp_path: Path) -> None:
    prefs = load_preferences(tmp_path / "missing.json")
    assert prefs.compress is True
    assert prefs.optimize is True
    assert prefs.last_folder == ""
    assert prefs.appearance_mode == "System"


def test_appearance_mode_normalized() -> None:
    prefs = preferences_from_dict({"appearance_mode": "dark"})
    assert prefs.appearance_mode == "Dark"
    prefs = preferences_from_dict({"appearance_mode": "nope"})
    assert prefs.appearance_mode == "System"


def test_load_corrupt_json_returns_defaults(tmp_path: Path) -> None:
    path = tmp_path / "bad.json"
    path.write_text("{not json", encoding="utf-8")
    prefs = load_preferences(path)
    assert prefs.ocr_language == "eng"


def test_resolved_paths_ignore_missing(tmp_path: Path) -> None:
    prefs = AppPreferences(last_folder=str(tmp_path / "gone"))
    assert prefs.resolved_last_folder() is None
