"""Progress and completion dialogs for the GUI."""

from __future__ import annotations

import sys
import tkinter as tk
from collections.abc import Callable, Sequence
from pathlib import Path

import customtkinter as ctk
from tkinter import messagebox

from coversheets.process import BatchResult
from coversheets.util import format_result_summary, open_in_file_manager


def center_on_master(window: ctk.CTkToplevel | tk.Toplevel, master: tk.Misc) -> None:
    """Position *window* roughly centered over *master*."""
    window.update_idletasks()
    try:
        mx = master.winfo_rootx()
        my = master.winfo_rooty()
        mw = master.winfo_width()
        mh = master.winfo_height()
        w = window.winfo_width() or window.winfo_reqwidth()
        h = window.winfo_height() or window.winfo_reqheight()
        window.geometry(f"+{mx + (mw - w) // 2}+{my + (mh - h) // 2}")
    except tk.TclError:
        pass


class ProgressWindow(ctk.CTkToplevel):
    """Modal-style progress dialog shown while generating cover sheets."""

    def __init__(
        self,
        master: tk.Misc,
        *,
        total: int,
        on_cancel: Callable[[], None] | None = None,
    ) -> None:
        super().__init__(master)
        self.title("Generating cover sheets…")
        self.resizable(True, True)
        self.minsize(420, 280)
        self.geometry("520x340")
        self.transient(master)
        self.protocol("WM_DELETE_WINDOW", self._on_close_attempt)

        self._total = max(total, 1)
        self._on_cancel = on_cancel
        self._cancel_requested = False
        self.status_var = ctk.StringVar(value="Starting…")
        self.count_var = ctk.StringVar(value=f"0 / {total}")

        pad = {"padx": 12, "pady": 6}

        ctk.CTkLabel(
            self,
            text="Please wait while cover sheets are generated.",
            font=ctk.CTkFont(size=13),
            anchor="w",
        ).pack(anchor="w", **pad)

        ctk.CTkLabel(self, textvariable=self.status_var, anchor="w").pack(
            anchor="w", padx=12
        )

        bar_row = ctk.CTkFrame(self, fg_color="transparent")
        bar_row.pack(fill="x", padx=12, pady=8)
        self.bar = ctk.CTkProgressBar(bar_row, mode="determinate")
        self.bar.set(0.0)
        self.bar.pack(fill="x", side="left", expand=True)
        ctk.CTkLabel(bar_row, textvariable=self.count_var, width=60).pack(
            side="right", padx=(8, 0)
        )

        log_frame = ctk.CTkFrame(self)
        log_frame.pack(fill="both", expand=True, padx=12, pady=6)
        ctk.CTkLabel(log_frame, text="Log", anchor="w").pack(
            anchor="w", padx=8, pady=(6, 0)
        )
        mono = (
            ctk.CTkFont(family="Menlo", size=11)
            if sys.platform == "darwin"
            else ctk.CTkFont(family="Consolas", size=11)
        )
        self.log = ctk.CTkTextbox(log_frame, height=140, font=mono, wrap="word")
        self.log.pack(fill="both", expand=True, padx=8, pady=6)
        self.log.configure(state="disabled")

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=12, pady=(0, 12))
        self.cancel_btn = ctk.CTkButton(
            btn_row, text="Cancel", width=100, command=self._request_cancel
        )
        self.cancel_btn.pack(side="right")
        ctk.CTkLabel(
            btn_row,
            text="Cancel stops after the current file.",
            text_color=("gray40", "gray60"),
            anchor="w",
        ).pack(side="left")

        center_on_master(self, master)
        self.grab_set()
        self.focus_set()

    def _request_cancel(self) -> None:
        if self._cancel_requested:
            return
        self._cancel_requested = True
        self.cancel_btn.configure(state="disabled", text="Cancelling…")
        self.status_var.set("Cancelling after current file…")
        self.append_log("Cancel requested — will stop after the current file.")
        if self._on_cancel is not None:
            self._on_cancel()

    def _on_close_attempt(self) -> None:
        if self._cancel_requested:
            messagebox.showinfo(
                "Cancelling",
                "Cancel already requested.\n"
                "This window will close when the current file finishes.",
                parent=self,
            )
            return
        if messagebox.askyesno(
            "Cancel generation?",
            "Stop after the current file finishes?\n\n"
            "Files already written will be kept.",
            parent=self,
        ):
            self._request_cancel()

    def set_progress(self, current: int, total: int, name: str) -> None:
        total = max(total, 1)
        self._total = total
        # Show progress for the file about to start / in progress.
        pct = ((current - 1) / total) if current > 0 else 0.0
        self.bar.set(min(max(pct, 0.0), 1.0))
        self.count_var.set(f"{current} / {total}")
        if not self._cancel_requested:
            self.status_var.set(f"Processing: {name}")

    def append_log(self, line: str) -> None:
        if not line:
            return
        self.log.configure(state="normal")
        self.log.insert("end", line.rstrip() + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def mark_complete(self, *, cancelled: bool = False) -> None:
        if not cancelled:
            self.bar.set(1.0)
        self.count_var.set(f"{self._total} / {self._total}")
        self.status_var.set("Cancelled." if cancelled else "Finished.")
        try:
            self.cancel_btn.configure(state="disabled")
        except tk.TclError:
            pass


class DoneDialog(ctk.CTkToplevel):
    """Success / summary dialog with optional open-folder actions."""

    def __init__(
        self,
        master: tk.Misc,
        *,
        result: BatchResult,
        folders: Sequence[Path],
        open_when_done: bool,
    ) -> None:
        super().__init__(master)
        self.transient(master)
        self.resizable(False, False)
        self.folders = [Path(p) for p in folders]

        ok = result.ok and result.failed == 0 and not result.was_cancelled
        if result.was_cancelled:
            self.title("Cancelled")
            headline = "Generation cancelled."
        elif ok and result.succeeded == 0 and result.skipped:
            self.title("Nothing new written")
            headline = "All selected files were skipped."
        elif ok:
            self.title("Success")
            headline = "Cover sheets generated successfully."
        else:
            self.title("Finished with errors")
            headline = "Finished, but some files failed."

        pad = {"padx": 16, "pady": 6}

        ctk.CTkLabel(
            self,
            text=headline,
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w",
        ).pack(anchor="w", padx=16, pady=(16, 4))
        ctk.CTkLabel(
            self, text=format_result_summary(result), anchor="w", justify="left"
        ).pack(anchor="w", **pad)

        if result.errors:
            err_frame = ctk.CTkFrame(self)
            err_frame.pack(fill="both", expand=True, padx=16, pady=6)
            ctk.CTkLabel(err_frame, text="Errors", anchor="w").pack(
                anchor="w", padx=8, pady=(6, 0)
            )
            err_text = ctk.CTkTextbox(
                err_frame,
                height=min(120, max(48, 20 * len(result.errors))),
                wrap="word",
            )
            err_text.pack(fill="both", expand=True, padx=8, pady=6)
            for name, msg in result.errors:
                err_text.insert("end", f"• {name}: {msg}\n")
            err_text.configure(state="disabled")

        if self.folders:
            folder_label = (
                str(self.folders[0])
                if len(self.folders) == 1
                else f"{len(self.folders)} folders (outputs next to sources)"
            )
            ctk.CTkLabel(
                self,
                text=f"Output: {folder_label}",
                text_color=("gray40", "gray60"),
                anchor="w",
                justify="left",
            ).pack(anchor="w", padx=16, pady=(4, 8))

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=16, pady=(4, 16))

        if self.folders:
            open_label = (
                "Open Output Folder"
                if len(self.folders) == 1
                else "Open Output Folders"
            )
            ctk.CTkButton(
                btn_row, text=open_label, command=self._open_folders
            ).pack(side="left")

        ctk.CTkButton(btn_row, text="Close", width=90, command=self._close).pack(
            side="right"
        )

        self.protocol("WM_DELETE_WINDOW", self._close)
        self.grab_set()
        self.focus_set()
        center_on_master(self, master)

        if open_when_done and self.folders and ok:
            # Defer so the dialog is visible first.
            self.after(150, self._open_folders)

    def _open_folders(self) -> None:
        errors: list[str] = []
        for folder in self.folders:
            try:
                open_in_file_manager(folder)
            except OSError as exc:
                errors.append(f"{folder}: {exc}")
        if errors:
            messagebox.showerror(
                "Could not open folder",
                "\n".join(errors),
                parent=self,
            )

    def _close(self) -> None:
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()
