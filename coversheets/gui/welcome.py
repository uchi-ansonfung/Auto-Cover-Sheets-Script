"""First-run welcome dialog for non-technical users."""

from __future__ import annotations

import tkinter as tk
from collections.abc import Callable

import customtkinter as ctk

from coversheets import OUTPUT_PREFIX
from coversheets.gui.dialogs import center_on_master


class WelcomeDialog(ctk.CTkToplevel):
    """Short guided intro shown once (unless the user opts out)."""

    def __init__(
        self,
        master: tk.Misc,
        *,
        on_finish: Callable[[bool], None] | None = None,
    ) -> None:
        super().__init__(master)
        self.title("Welcome")
        self.resizable(False, False)
        self.transient(master)
        self._on_finish = on_finish
        self._dont_show = ctk.BooleanVar(value=False)

        pad = {"padx": 20, "pady": 6}

        ctk.CTkLabel(
            self,
            text="Add cover sheets to exhibit PDFs",
            font=ctk.CTkFont(size=16, weight="bold"),
            anchor="w",
        ).pack(anchor="w", padx=20, pady=(20, 8))

        steps = [
            "1. Open a folder of PDFs (or add files).",
            "2. Check each cover title — edit if you like.",
            "3. Click Generate cover sheets.",
            f"4. Find new files named {OUTPUT_PREFIX}…pdf "
            "(your originals stay unchanged).",
        ]
        for line in steps:
            ctk.CTkLabel(self, text=line, anchor="w", justify="left").pack(
                anchor="w", **pad
            )

        ctk.CTkLabel(
            self,
            text="Tip: turn on “Make text searchable” if the scans should be searchable.",
            text_color=("gray40", "gray60"),
            anchor="w",
            justify="left",
            wraplength=400,
        ).pack(anchor="w", padx=20, pady=(8, 4))

        ctk.CTkCheckBox(
            self,
            text="Don’t show this again",
            variable=self._dont_show,
        ).pack(anchor="w", padx=20, pady=(12, 4))

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=(8, 20))
        ctk.CTkButton(btn_row, text="Get started", width=120, command=self._close).pack(
            side="right"
        )

        self.protocol("WM_DELETE_WINDOW", self._close)
        self.grab_set()
        self.focus_set()
        center_on_master(self, master)

    def _close(self) -> None:
        dont = bool(self._dont_show.get())
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()
        if self._on_finish is not None:
            self._on_finish(dont)
