"""Legal-workflow options, advanced accordion, presets, and output mode."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import customtkinter as ctk
from tkinter import filedialog

from coversheets.cover import DEFAULT_VERTICAL_POSITION, normalize_vertical_position
from coversheets.gui.copy import (
    PRESET_LABELS,
    VERTICAL_POSITION_LABELS,
    linearize_unavailable_message,
    ocr_unavailable_message,
    output_example,
)
from coversheets.pdf_ops import linearize_available, ocr_available
from coversheets.prefs import AppPreferences


class OptionsPanel(ctk.CTkFrame):
    """
    Front-and-center legal options + collapsed Advanced.

    Exposes BooleanVars / StringVars the app can read when generating.
    """

    def __init__(
        self,
        master: ctk.CTkBaseClass,
        prefs: AppPreferences,
        *,
        on_change: Callable[[], None] | None = None,
        dialog_initial_dir: Callable[[], str | None] | None = None,
    ) -> None:
        super().__init__(master)
        self._on_change = on_change
        self._dialog_initial_dir = dialog_initial_dir
        self._suppress_preset = False

        self.output_mode_var = ctk.StringVar(value=prefs.output_mode or "beside")
        self.output_dir_var = ctk.StringVar(value=prefs.output_dir or "")
        self.compress_var = ctk.BooleanVar(value=prefs.compress)
        self.force_var = ctk.BooleanVar(value=prefs.force)
        self.rename_to_label_var = ctk.BooleanVar(value=prefs.rename_to_label)
        self.strip_metadata_var = ctk.BooleanVar(value=prefs.strip_metadata)
        self.optimize_var = ctk.BooleanVar(value=prefs.optimize)
        self.linearize_var = ctk.BooleanVar(value=prefs.linearize)
        self.ocr_var = ctk.BooleanVar(value=prefs.ocr)
        self.ocr_language_var = ctk.StringVar(value=prefs.ocr_language or "eng")
        self.open_when_done_var = ctk.BooleanVar(value=prefs.open_when_done)
        position_id = normalize_vertical_position(prefs.vertical_position)
        self.vertical_position_var = ctk.StringVar(
            value=VERTICAL_POSITION_LABELS.get(
                position_id, VERTICAL_POSITION_LABELS[DEFAULT_VERTICAL_POSITION]
            )
        )
        self.preset_var = ctk.StringVar(
            value=PRESET_LABELS.get(prefs.preset, PRESET_LABELS["recommended"])
        )
        self._advanced_open = bool(prefs.advanced_expanded)
        self.example_var = ctk.StringVar(value="")

        self._build()
        self._sync_output_widgets()
        self._update_example()
        self._apply_capability_states()

    # --- public API ------------------------------------------------------

    def collect_into_prefs(self, base: AppPreferences) -> AppPreferences:
        """Return a copy of *base* with option fields from this panel."""
        preset_id = self._preset_id_from_label(self.preset_var.get())
        return AppPreferences(
            version=base.version,
            last_folder=base.last_folder,
            last_file_dialog_dir=base.last_file_dialog_dir,
            output_dir=self.output_dir_var.get().strip(),
            output_mode=self.output_mode_var.get().strip() or "beside",
            window_geometry=base.window_geometry,
            appearance_mode=base.appearance_mode,
            compress=self.compress_var.get(),
            force=self.force_var.get(),
            rename_to_label=self.rename_to_label_var.get(),
            strip_metadata=self.strip_metadata_var.get(),
            optimize=self.optimize_var.get(),
            linearize=self.linearize_var.get(),
            ocr=self.ocr_var.get(),
            ocr_language=self.ocr_language_var.get().strip() or "eng",
            open_when_done=self.open_when_done_var.get(),
            vertical_position=self.vertical_position_id(),
            show_welcome=base.show_welcome,
            advanced_expanded=self._advanced_open,
            preset=preset_id,
        )

    def vertical_position_id(self) -> str:
        """Return canonical vertical_position id from the UI label."""
        label = self.vertical_position_var.get()
        for key, value in VERTICAL_POSITION_LABELS.items():
            if value == label:
                return key
        return DEFAULT_VERTICAL_POSITION

    def resolved_output_dir(self) -> Path | None:
        if self.output_mode_var.get() != "folder":
            return None
        raw = self.output_dir_var.get().strip()
        if not raw:
            return None
        path = Path(raw).expanduser()
        try:
            return path.resolve()
        except OSError:
            return path

    def mark_custom_if_needed(self) -> None:
        if self._suppress_preset:
            return
        label = PRESET_LABELS["custom"]
        if self.preset_var.get() != label:
            self.preset_var.set(label)

    # --- build -----------------------------------------------------------

    def _build(self) -> None:
        pad_x, pad_y = 10, 6

        # Preset + output
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=pad_x, pady=(pad_y, 2))

        ctk.CTkLabel(top, text="Preset", font=ctk.CTkFont(weight="bold")).pack(
            side="left"
        )
        self.preset_menu = ctk.CTkOptionMenu(
            top,
            values=list(PRESET_LABELS.values()),
            variable=self.preset_var,
            width=220,
            command=self._on_preset_chosen,
        )
        self.preset_menu.pack(side="left", padx=(8, 16))

        ctk.CTkLabel(top, text="Save results", font=ctk.CTkFont(weight="bold")).pack(
            side="left"
        )
        self.beside_radio = ctk.CTkRadioButton(
            top,
            text="Next to each original",
            variable=self.output_mode_var,
            value="beside",
            command=self._on_output_mode,
        )
        self.beside_radio.pack(side="left", padx=(8, 8))
        self.folder_radio = ctk.CTkRadioButton(
            top,
            text="One folder…",
            variable=self.output_mode_var,
            value="folder",
            command=self._on_output_mode,
        )
        self.folder_radio.pack(side="left", padx=(0, 8))

        out_row = ctk.CTkFrame(self, fg_color="transparent")
        out_row.pack(fill="x", padx=pad_x, pady=(0, 2))
        self.output_entry = ctk.CTkEntry(out_row, textvariable=self.output_dir_var)
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.browse_btn = ctk.CTkButton(
            out_row, text="Choose folder…", width=120, command=self._on_browse_output
        )
        self.browse_btn.pack(side="left")
        self.output_dir_var.trace_add("write", lambda *_: self._on_field_change())

        ctk.CTkLabel(
            self,
            textvariable=self.example_var,
            text_color=("gray40", "gray60"),
            anchor="w",
            justify="left",
        ).pack(fill="x", padx=pad_x, pady=(0, pad_y))

        # Title vertical position (layout)
        pos_row = ctk.CTkFrame(self, fg_color="transparent")
        pos_row.pack(fill="x", padx=pad_x, pady=(0, 4))
        ctk.CTkLabel(
            pos_row, text="Title position", font=ctk.CTkFont(weight="bold")
        ).pack(side="left")
        self.center_pos_radio = ctk.CTkRadioButton(
            pos_row,
            text=VERTICAL_POSITION_LABELS["center"],
            variable=self.vertical_position_var,
            value=VERTICAL_POSITION_LABELS["center"],
            command=self._on_vertical_position,
        )
        self.center_pos_radio.pack(side="left", padx=(8, 8))
        self.top_third_pos_radio = ctk.CTkRadioButton(
            pos_row,
            text=VERTICAL_POSITION_LABELS["top_third"],
            variable=self.vertical_position_var,
            value=VERTICAL_POSITION_LABELS["top_third"],
            command=self._on_vertical_position,
        )
        self.top_third_pos_radio.pack(side="left", padx=(0, 8))

        # Legal options
        legal = ctk.CTkFrame(self, fg_color="transparent")
        legal.pack(fill="x", padx=pad_x, pady=(0, 4))
        self._legal_cbs: list[ctk.CTkCheckBox] = []
        for text, var, key in (
            ("Remove hidden document info", self.strip_metadata_var, "strip"),
            ("Make text searchable (OCR)", self.ocr_var, "ocr"),
            ("Replace existing +files", self.force_var, "force"),
            ("Open folder when finished", self.open_when_done_var, "open"),
        ):
            cb = ctk.CTkCheckBox(
                legal,
                text=text,
                variable=var,
                command=self._on_toggle,
            )
            cb.pack(side="left", padx=(0, 14))
            self._legal_cbs.append(cb)
            if key == "ocr":
                self._ocr_cb = cb

        self._ocr_hint = ctk.CTkLabel(
            self,
            text="",
            text_color=("gray40", "gray60"),
            anchor="w",
            font=ctk.CTkFont(size=12),
        )
        self._ocr_hint.pack(fill="x", padx=pad_x, pady=(0, 4))

        # Advanced accordion
        adv_header = ctk.CTkFrame(self, fg_color="transparent")
        adv_header.pack(fill="x", padx=pad_x, pady=(2, 0))
        self._adv_toggle = ctk.CTkButton(
            adv_header,
            text=self._advanced_button_text(),
            width=160,
            fg_color="transparent",
            border_width=1,
            text_color=("gray20", "gray80"),
            command=self._toggle_advanced,
        )
        self._adv_toggle.pack(side="left")

        self._adv_body = ctk.CTkFrame(self)
        row1 = ctk.CTkFrame(self._adv_body, fg_color="transparent")
        row1.pack(fill="x", padx=10, pady=(8, 4))
        for text, var in (
            ("Compress PDF data", self.compress_var),
            ("Shrink duplicate PDF data", self.optimize_var),
            ("Faster web viewing", self.linearize_var),
            ("Name output file after cover title", self.rename_to_label_var),
        ):
            cb = ctk.CTkCheckBox(
                row1, text=text, variable=var, command=self._on_toggle
            )
            cb.pack(side="left", padx=(0, 12))
            if "web viewing" in text:
                self._linearize_cb = cb

        row2 = ctk.CTkFrame(self._adv_body, fg_color="transparent")
        row2.pack(fill="x", padx=10, pady=(0, 8))
        ctk.CTkLabel(row2, text="OCR language:").pack(side="left")
        self._ocr_lang_entry = ctk.CTkEntry(
            row2, textvariable=self.ocr_language_var, width=80
        )
        self._ocr_lang_entry.pack(side="left", padx=(6, 8))
        ctk.CTkLabel(
            row2,
            text="Usually eng for English",
            text_color=("gray40", "gray60"),
        ).pack(side="left")
        self.ocr_language_var.trace_add("write", lambda *_: self._on_field_change())
        self.rename_to_label_var.trace_add(
            "write", lambda *_: self._update_example()
        )

        if self._advanced_open:
            self._adv_body.pack(fill="x", padx=pad_x, pady=(4, pad_y))

    def _advanced_button_text(self) -> str:
        return "▾ More options" if self._advanced_open else "▸ More options"

    def _toggle_advanced(self) -> None:
        self._advanced_open = not self._advanced_open
        self._adv_toggle.configure(text=self._advanced_button_text())
        if self._advanced_open:
            self._adv_body.pack(fill="x", padx=10, pady=(4, 6))
        else:
            self._adv_body.pack_forget()
        self._notify()

    def _apply_capability_states(self) -> None:
        if ocr_available():
            self._ocr_cb.configure(state="normal")
            self._ocr_hint.configure(text="")
        else:
            self.ocr_var.set(False)
            self._ocr_cb.configure(state="disabled")
            self._ocr_hint.configure(
                text="Searchable text isn’t available in this install "
                "(use the Windows full setup package)."
            )
        if linearize_available():
            self._linearize_cb.configure(state="normal")
        else:
            self.linearize_var.set(False)
            self._linearize_cb.configure(state="disabled")

    def _sync_output_widgets(self) -> None:
        folder_mode = self.output_mode_var.get() == "folder"
        state = "normal" if folder_mode else "disabled"
        self.output_entry.configure(state=state)
        self.browse_btn.configure(state=state)

    def _update_example(self) -> None:
        self.example_var.set(
            output_example(
                mode=self.output_mode_var.get(),
                folder=self.output_dir_var.get().strip() or None,
                rename_to_label=self.rename_to_label_var.get(),
            )
        )

    def _on_output_mode(self) -> None:
        self._sync_output_widgets()
        self._update_example()
        self.mark_custom_if_needed()
        self._notify()

    def _on_vertical_position(self) -> None:
        self.mark_custom_if_needed()
        self._notify()

    def _on_browse_output(self) -> None:
        kwargs: dict[str, Any] = {"title": "Choose a folder for finished PDFs"}
        out = self.output_dir_var.get().strip()
        if out and Path(out).expanduser().is_dir():
            kwargs["initialdir"] = str(Path(out).expanduser())
        elif self._dialog_initial_dir is not None:
            initial = self._dialog_initial_dir()
            if initial:
                kwargs["initialdir"] = initial
        path = filedialog.askdirectory(**kwargs)
        if path:
            self.output_mode_var.set("folder")
            self.output_dir_var.set(path)
            self._sync_output_widgets()
            self._update_example()
            self.mark_custom_if_needed()
            self._notify()

    def _on_toggle(self) -> None:
        self.mark_custom_if_needed()
        self._update_example()
        self._notify()

    def _on_field_change(self) -> None:
        self.mark_custom_if_needed()
        self._update_example()
        self._notify()

    def _preset_id_from_label(self, label: str) -> str:
        for key, value in PRESET_LABELS.items():
            if value == label:
                return key
        return "custom"

    def _on_preset_chosen(self, label: str) -> None:
        preset_id = self._preset_id_from_label(label)
        if preset_id == "custom":
            self._notify()
            return
        self._suppress_preset = True
        try:
            # Shared baseline
            self.compress_var.set(True)
            self.optimize_var.set(True)
            self.linearize_var.set(False)
            self.rename_to_label_var.set(False)
            self.strip_metadata_var.set(True)
            self.force_var.set(False)
            self.open_when_done_var.set(True)
            self.ocr_language_var.set("eng")
            if preset_id == "recommended":
                self.ocr_var.set(False)
            elif preset_id == "searchable":
                if ocr_available():
                    self.ocr_var.set(True)
                else:
                    self.ocr_var.set(False)
                    self.preset_var.set(PRESET_LABELS["recommended"])
            self._apply_capability_states()
            self._update_example()
        finally:
            self._suppress_preset = False
        self._notify()

    def _notify(self) -> None:
        if self._on_change is not None:
            self._on_change()

    def ensure_capabilities_or_warn(self) -> str | None:
        """
        Return an error message if the user asked for unavailable features.
        """
        if self.ocr_var.get() and not ocr_available():
            return ocr_unavailable_message()
        if self.linearize_var.get() and not linearize_available():
            return linearize_unavailable_message()
        return None
