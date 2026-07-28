"""Scrollable job list with per-row include checkbox and editable label."""

from __future__ import annotations

import sys
import tkinter as tk
from collections.abc import Callable, Sequence

import customtkinter as ctk

from coversheets.cover import cover_label_from_filename
from coversheets.process import JobItem

# Column weights for grid layout (include is fixed-ish via minsize).
_COL_INCLUDE = 0
_COL_FILE = 1
_COL_LABEL = 2
_COL_FOLDER = 3


class JobRow(ctk.CTkFrame):
    """One PDF job row: include, filename, cover label, folder."""

    def __init__(
        self,
        master: ctk.CTkBaseClass,
        job: JobItem,
        index: int,
        *,
        on_select: Callable[[int, bool, bool], None],
        on_include_changed: Callable[[], None],
        on_label_changed: Callable[[], None],
    ) -> None:
        super().__init__(master, corner_radius=4, border_width=1)
        self.job = job
        self.index = index
        self._on_select = on_select
        self._on_include_changed = on_include_changed
        self._on_label_changed = on_label_changed
        self._selected = False
        self._include_var = ctk.BooleanVar(value=job.include)

        self.grid_columnconfigure(_COL_FILE, weight=2, minsize=100)
        self.grid_columnconfigure(_COL_LABEL, weight=3, minsize=120)
        self.grid_columnconfigure(_COL_FOLDER, weight=2, minsize=100)

        self.include_cb = ctk.CTkCheckBox(
            self,
            text="",
            variable=self._include_var,
            width=28,
            command=self._on_include,
        )
        self.include_cb.grid(row=0, column=_COL_INCLUDE, padx=(8, 4), pady=4)

        self.file_label = ctk.CTkLabel(
            self, text=job.source.name, anchor="w", cursor="hand2"
        )
        self.file_label.grid(row=0, column=_COL_FILE, sticky="ew", padx=4, pady=4)

        self.label_entry = ctk.CTkEntry(self)
        self.label_entry.insert(0, job.label)
        self.label_entry.grid(row=0, column=_COL_LABEL, sticky="ew", padx=4, pady=4)
        self.label_entry.bind("<Return>", self._commit_label)
        self.label_entry.bind("<FocusOut>", self._commit_label)

        self.folder_label = ctk.CTkLabel(
            self,
            text=str(job.source.parent),
            anchor="w",
            text_color=("gray40", "gray60"),
            cursor="hand2",
        )
        self.folder_label.grid(row=0, column=_COL_FOLDER, sticky="ew", padx=(4, 8), pady=4)

        for widget in (self, self.file_label, self.folder_label):
            widget.bind("<Button-1>", self._on_click)
        # Checkbox click should not also toggle selection in a surprising way,
        # but row selection on the frame still works.

        self._apply_selected_style()

    def _on_include(self) -> None:
        self.job.include = bool(self._include_var.get())
        self._on_include_changed()

    def _commit_label(self, _event: object | None = None) -> None:
        raw = self.label_entry.get().strip()
        if not raw:
            raw = cover_label_from_filename(self.job.source)
            self.label_entry.delete(0, "end")
            self.label_entry.insert(0, raw)
        if raw != self.job.label:
            self.job.label = raw
            self._on_label_changed()

    def _on_click(self, event: tk.Event) -> None:  # type: ignore[type-arg]
        ctrl = bool(event.state & 0x0004)  # Control
        if sys.platform == "darwin":
            # Command key on macOS is Mod1 (0x0008) in Tk.
            ctrl = ctrl or bool(event.state & 0x0008)
        shift = bool(event.state & 0x0001)
        self._on_select(self.index, ctrl, shift)

    def set_selected(self, selected: bool) -> None:
        if self._selected == selected:
            return
        self._selected = selected
        self._apply_selected_style()

    def _apply_selected_style(self) -> None:
        if self._selected:
            self.configure(border_color=("#3B8ED0", "#1F6AA5"), border_width=2)
        else:
            self.configure(border_color=("gray70", "gray35"), border_width=1)

    def sync_from_job(self) -> None:
        """Refresh widgets from the bound JobItem."""
        self._include_var.set(self.job.include)
        current = self.label_entry.get()
        if current != self.job.label and self.label_entry.focus_get() is not self.label_entry:
            self.label_entry.delete(0, "end")
            self.label_entry.insert(0, self.job.label)
        self.file_label.configure(text=self.job.source.name)
        self.folder_label.configure(text=str(self.job.source.parent))


class JobListFrame(ctk.CTkFrame):
    """Header + scrollable list of JobRow widgets with multi-select."""

    def __init__(
        self,
        master: ctk.CTkBaseClass,
        *,
        on_jobs_changed: Callable[[], None] | None = None,
    ) -> None:
        super().__init__(master, fg_color="transparent")
        self._on_jobs_changed = on_jobs_changed
        self.jobs: list[JobItem] = []
        self._rows: list[JobRow] = []
        self._selected: set[int] = set()
        self._anchor: int | None = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        header.grid_columnconfigure(_COL_FILE, weight=2, minsize=100)
        header.grid_columnconfigure(_COL_LABEL, weight=3, minsize=120)
        header.grid_columnconfigure(_COL_FOLDER, weight=2, minsize=100)

        bold = ctk.CTkFont(weight="bold")
        ctk.CTkLabel(header, text="Include", font=bold, width=56).grid(
            row=0, column=_COL_INCLUDE, padx=(8, 4)
        )
        ctk.CTkLabel(header, text="File", font=bold, anchor="w").grid(
            row=0, column=_COL_FILE, sticky="ew", padx=4
        )
        ctk.CTkLabel(header, text="Cover Label", font=bold, anchor="w").grid(
            row=0, column=_COL_LABEL, sticky="ew", padx=4
        )
        ctk.CTkLabel(header, text="Folder", font=bold, anchor="w").grid(
            row=0, column=_COL_FOLDER, sticky="ew", padx=(4, 8)
        )

        self.scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll.grid(row=1, column=0, sticky="nsew")
        self.scroll.grid_columnconfigure(0, weight=1)

        # Select-all shortcuts (window-level binding is attached by the app).
        self.bind("<Control-a>", self._on_select_all)
        self.bind("<Control-A>", self._on_select_all)
        self.scroll.bind("<Control-a>", self._on_select_all)
        self.scroll.bind("<Control-A>", self._on_select_all)
        if sys.platform == "darwin":
            self.bind("<Command-a>", self._on_select_all)
            self.bind("<Command-A>", self._on_select_all)
            self.scroll.bind("<Command-a>", self._on_select_all)
            self.scroll.bind("<Command-A>", self._on_select_all)

    def set_jobs(self, jobs: Sequence[JobItem]) -> None:
        self.jobs = list(jobs)
        self._selected.clear()
        self._anchor = None
        self._rebuild()
        self._notify()

    def get_jobs(self) -> list[JobItem]:
        self._commit_all_labels()
        return self.jobs

    def selected_indices(self) -> list[int]:
        return sorted(i for i in self._selected if 0 <= i < len(self.jobs))

    def select_all(self) -> None:
        if not self.jobs:
            self._selected.clear()
            self._anchor = None
        else:
            self._selected = set(range(len(self.jobs)))
            self._anchor = 0
        self._refresh_selection_styles()

    def clear_selection(self) -> None:
        self._selected.clear()
        self._anchor = None
        self._refresh_selection_styles()

    def remove_selected(self) -> None:
        indices = self.selected_indices()
        if not indices:
            return
        self._commit_all_labels()
        for index in reversed(indices):
            del self.jobs[index]
        self._selected.clear()
        self._anchor = None
        self._rebuild()
        self._notify()

    def remove_all(self) -> None:
        if not self.jobs:
            return
        self.jobs.clear()
        self._selected.clear()
        self._anchor = None
        self._rebuild()
        self._notify()

    def reset_labels(self) -> None:
        self._commit_all_labels()
        indices = self.selected_indices()
        targets = indices if indices else list(range(len(self.jobs)))
        for index in targets:
            job = self.jobs[index]
            job.label = cover_label_from_filename(job.source)
        for row in self._rows:
            row.sync_from_job()
        self._notify()

    def _on_select_all(self, _event: object | None = None) -> str:
        self.select_all()
        return "break"

    def _on_row_select(self, index: int, ctrl: bool, shift: bool) -> None:
        if not (0 <= index < len(self.jobs)):
            return
        if shift and self._anchor is not None:
            lo, hi = sorted((self._anchor, index))
            if ctrl:
                self._selected.update(range(lo, hi + 1))
            else:
                self._selected = set(range(lo, hi + 1))
        elif ctrl:
            if index in self._selected:
                self._selected.discard(index)
            else:
                self._selected.add(index)
            self._anchor = index
        else:
            self._selected = {index}
            self._anchor = index
        self._refresh_selection_styles()

    def _refresh_selection_styles(self) -> None:
        for row in self._rows:
            row.set_selected(row.index in self._selected)

    def _rebuild(self) -> None:
        for row in self._rows:
            row.destroy()
        self._rows.clear()
        for index, job in enumerate(self.jobs):
            row = JobRow(
                self.scroll,
                job,
                index,
                on_select=self._on_row_select,
                on_include_changed=self._notify,
                on_label_changed=self._notify,
            )
            row.grid(row=index, column=0, sticky="ew", pady=2)
            row.set_selected(index in self._selected)
            self._rows.append(row)

    def _commit_all_labels(self) -> None:
        for row in self._rows:
            row._commit_label()

    def _notify(self) -> None:
        if self._on_jobs_changed is not None:
            self._on_jobs_changed()
