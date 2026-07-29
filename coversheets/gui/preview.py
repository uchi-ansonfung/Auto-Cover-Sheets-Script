"""Live cover-sheet mockup (no PDF rasterization)."""

from __future__ import annotations

import tkinter as tk

import customtkinter as ctk


class CoverPreview(ctk.CTkFrame):
    """
    Portrait card that approximates the letter cover sheet.

    Shows the cover title centered in bold, matching the real PDF layout
    closely enough for non-technical users to trust the label.
    """

    def __init__(self, master: ctk.CTkBaseClass, **kwargs: object) -> None:
        super().__init__(master, **kwargs)  # type: ignore[arg-type]
        self._title = ""
        self._filename = ""

        ctk.CTkLabel(
            self,
            text="Cover preview",
            font=ctk.CTkFont(weight="bold"),
            anchor="w",
        ).pack(anchor="w", padx=8, pady=(8, 4))

        # Letter aspect ~ 8.5 x 11 → width/height ≈ 0.773
        self._page = ctk.CTkFrame(
            self,
            fg_color=("white", "#f0f0f0"),
            border_width=1,
            border_color=("gray70", "gray40"),
            width=200,
            height=260,
        )
        self._page.pack(padx=16, pady=8)
        self._page.pack_propagate(False)

        self._title_label = ctk.CTkLabel(
            self._page,
            text="Select a file\nto preview",
            font=ctk.CTkFont(family="Times", size=16, weight="bold"),
            text_color=("gray30", "gray20"),
            justify="center",
            wraplength=170,
        )
        self._title_label.place(relx=0.5, rely=0.5, anchor="center")

        self._caption = ctk.CTkLabel(
            self,
            text="",
            text_color=("gray40", "gray60"),
            anchor="w",
            justify="left",
            wraplength=220,
        )
        self._caption.pack(fill="x", padx=8, pady=(0, 8))

        self.bind("<Configure>", self._on_resize)

    def clear(self, message: str = "Select a file to preview its cover.") -> None:
        self._title = ""
        self._filename = ""
        self._title_label.configure(text=message.replace(". ", ".\n"))
        self._caption.configure(text="")

    def show(self, *, title: str, filename: str = "") -> None:
        text = (title or "").strip() or "(empty title)"
        self._title = text
        self._filename = filename
        self._title_label.configure(text=text)
        if filename:
            self._caption.configure(text=f"Selected: {filename}")
        else:
            self._caption.configure(text="")

    def _on_resize(self, _event: tk.Event | None = None) -> None:  # type: ignore[type-arg]
        # Keep page roughly letter-shaped within available width.
        try:
            width = max(self.winfo_width() - 40, 140)
        except tk.TclError:
            return
        page_w = min(220, max(140, width))
        page_h = int(page_w * 11 / 8.5)
        page_h = min(page_h, 300)
        self._page.configure(width=page_w, height=page_h)
        self._title_label.configure(wraplength=max(page_w - 24, 80))
