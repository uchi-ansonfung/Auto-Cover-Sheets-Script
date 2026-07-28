"""Main CustomTkinter window: job list, options, and generate flow."""

from __future__ import annotations

import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from typing import Any

import customtkinter as ctk
from tkinter import filedialog, messagebox

from coversheets import OUTPUT_PREFIX, __author__, __version__
from coversheets.bundled_tools import configure_bundled_tools
from coversheets.gui.dialogs import DoneDialog, ProgressWindow
from coversheets.gui.job_list import JobListFrame
from coversheets.gui.theme import APPEARANCE_MODES, apply_appearance, normalize_appearance_mode
from coversheets.options import ProcessOptions
from coversheets.pdf_ops import linearize_available, ocr_available
from coversheets.prefs import AppPreferences, load_preferences, save_preferences
from coversheets.process import (
    BatchResult,
    JobItem,
    is_output_filename,
    jobs_from_folder,
    jobs_from_paths,
    process_jobs,
)
from coversheets.util import format_result_summary, resolve_result_folders


class CoverSheetsApp(ctk.CTk):
    """Main window: editable list of PDFs and cover labels."""

    def __init__(self, initial_folder: Path | None = None) -> None:
        apply_appearance("System")  # refined after prefs load
        super().__init__()
        configure_bundled_tools()
        self.title(f"Automatic Exhibit Cover Sheets v{__version__}")
        self.minsize(820, 520)

        self._prefs = load_preferences()
        self._prefs_path: Path | None = None
        self._initial_folder_arg = initial_folder

        apply_appearance(self._prefs.appearance_mode)

        self._worker: threading.Thread | None = None
        self._msg_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self._busy = False
        self._cancel_event = threading.Event()
        self._progress_win: ProgressWindow | None = None
        self._run_output_dir: Path | None = None
        self._run_jobs: list[JobItem] = []
        self._toolbar_buttons: list[ctk.CTkButton] = []

        p = self._prefs
        self.output_dir_var = ctk.StringVar(value=p.output_dir or "")
        self.compress_var = ctk.BooleanVar(value=p.compress)
        self.force_var = ctk.BooleanVar(value=p.force)
        self.rename_to_label_var = ctk.BooleanVar(value=p.rename_to_label)
        self.strip_metadata_var = ctk.BooleanVar(value=p.strip_metadata)
        self.optimize_var = ctk.BooleanVar(value=p.optimize)
        self.linearize_var = ctk.BooleanVar(value=p.linearize)
        self.ocr_var = ctk.BooleanVar(value=p.ocr)
        self.ocr_language_var = ctk.StringVar(value=p.ocr_language or "eng")
        self.open_when_done_var = ctk.BooleanVar(value=p.open_when_done)
        self.status_var = ctk.StringVar(value="Add PDFs or open a folder to begin.")
        self.hint_var = ctk.StringVar()
        self.appearance_var = ctk.StringVar(
            value=normalize_appearance_mode(p.appearance_mode)
        )

        geo = (p.window_geometry or "900x560").strip() or "900x560"
        try:
            self.geometry(geo)
        except tk.TclError:
            self.geometry("900x560")

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Control-a>", self._on_select_all)
        self.bind("<Control-A>", self._on_select_all)
        if sys.platform == "darwin":
            self.bind("<Command-a>", self._on_select_all)
            self.bind("<Command-A>", self._on_select_all)
        self.after(100, self._poll_queue)
        self.after(80, self._restore_session)

    # --- UI construction -------------------------------------------------

    def _build_ui(self) -> None:
        pad_x, pad_y = 10, 6

        # Toolbar
        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill="x", padx=pad_x, pady=(pad_y, 2))

        left = ctk.CTkFrame(toolbar, fg_color="transparent")
        left.pack(side="left", fill="x", expand=True)

        for text, cmd in (
            ("Open Folder…", self._on_open_folder),
            ("Add PDFs…", self._on_add_pdfs),
            ("Remove", self._on_remove),
            ("Remove All", self._on_remove_all),
            ("Reset Labels", self._on_reset_labels),
        ):
            btn = ctk.CTkButton(left, text=text, width=0, command=cmd)
            btn.pack(side="left", padx=(0, 6))
            self._toolbar_buttons.append(btn)

        theme_frame = ctk.CTkFrame(toolbar, fg_color="transparent")
        theme_frame.pack(side="right")
        ctk.CTkLabel(theme_frame, text="Theme").pack(side="left", padx=(0, 6))
        self.appearance_menu = ctk.CTkOptionMenu(
            theme_frame,
            values=list(APPEARANCE_MODES),
            variable=self.appearance_var,
            width=100,
            command=self._on_appearance_change,
        )
        self.appearance_menu.pack(side="left")

        # Output folder
        out_row = ctk.CTkFrame(self, fg_color="transparent")
        out_row.pack(fill="x", padx=pad_x, pady=pad_y)
        ctk.CTkLabel(out_row, text="Output folder:").pack(side="left")
        self.output_entry = ctk.CTkEntry(out_row, textvariable=self.output_dir_var)
        self.output_entry.pack(side="left", fill="x", expand=True, padx=8)
        browse_btn = ctk.CTkButton(
            out_row, text="Browse…", width=90, command=self._on_browse_output
        )
        browse_btn.pack(side="left")
        self._toolbar_buttons.append(browse_btn)
        ctk.CTkLabel(
            out_row,
            text="(empty = next to each source)",
            text_color=("gray40", "gray60"),
        ).pack(side="left", padx=(8, 0))

        # Process options
        process_box = ctk.CTkFrame(self)
        process_box.pack(fill="x", padx=pad_x, pady=(2, pad_y))
        ctk.CTkLabel(
            process_box,
            text="Process options",
            font=ctk.CTkFont(weight="bold"),
            anchor="w",
        ).pack(anchor="w", padx=10, pady=(8, 4))
        process_row = ctk.CTkFrame(process_box, fg_color="transparent")
        process_row.pack(fill="x", padx=10, pady=(0, 8))
        for text, var, extra in (
            ("Compress page streams", self.compress_var, {}),
            ("Overwrite existing +outputs", self.force_var, {}),
            (
                "Name output file after label",
                self.rename_to_label_var,
                {"command": self._update_hint},
            ),
            ("Open output folder when done", self.open_when_done_var, {}),
        ):
            ctk.CTkCheckBox(process_row, text=text, variable=var, **extra).pack(
                side="left", padx=(0, 14)
            )

        # PDF options
        pdf_box = ctk.CTkFrame(self)
        pdf_box.pack(fill="x", padx=pad_x, pady=(0, pad_y))
        ctk.CTkLabel(
            pdf_box,
            text="PDF options",
            font=ctk.CTkFont(weight="bold"),
            anchor="w",
        ).pack(anchor="w", padx=10, pady=(8, 4))

        pdf_row1 = ctk.CTkFrame(pdf_box, fg_color="transparent")
        pdf_row1.pack(fill="x", padx=10, pady=(0, 4))
        ctk.CTkCheckBox(
            pdf_row1,
            text="Strip metadata (Info + XMP)",
            variable=self.strip_metadata_var,
        ).pack(side="left", padx=(0, 14))
        ctk.CTkCheckBox(
            pdf_row1,
            text="Optimize (dedupe objects)",
            variable=self.optimize_var,
        ).pack(side="left", padx=(0, 14))
        lin_state = "normal" if linearize_available() else "disabled"
        self._linearize_cb = ctk.CTkCheckBox(
            pdf_row1,
            text="Linearize (web optimize)",
            variable=self.linearize_var,
            state=lin_state,
        )
        self._linearize_cb.pack(side="left", padx=(0, 14))

        pdf_row2 = ctk.CTkFrame(pdf_box, fg_color="transparent")
        pdf_row2.pack(fill="x", padx=10, pady=(0, 8))
        ocr_state = "normal" if ocr_available() else "disabled"
        self._ocr_cb = ctk.CTkCheckBox(
            pdf_row2,
            text="OCR (ocrmypdf)",
            variable=self.ocr_var,
            state=ocr_state,
        )
        self._ocr_cb.pack(side="left")
        ctk.CTkLabel(pdf_row2, text="Language:").pack(side="left", padx=(12, 4))
        self._ocr_lang_entry = ctk.CTkEntry(
            pdf_row2, textvariable=self.ocr_language_var, width=70
        )
        self._ocr_lang_entry.pack(side="left")
        if not ocr_available():
            ctk.CTkLabel(
                pdf_row2,
                text="(use Windows full installer, or coversheets[ocr] + Tesseract)",
                text_color=("gray50", "gray55"),
            ).pack(side="left", padx=(8, 0))
        if not linearize_available():
            ctk.CTkLabel(
                pdf_row2,
                text="Linearize needs coversheets[optimize] or qpdf",
                text_color=("gray50", "gray55"),
            ).pack(side="left", padx=(12, 0))

        # Hint
        self._update_hint()
        ctk.CTkLabel(
            self,
            textvariable=self.hint_var,
            text_color=("gray40", "gray60"),
            anchor="w",
        ).pack(fill="x", padx=pad_x)

        # Job list
        self.job_list = JobListFrame(self, on_jobs_changed=self._on_jobs_changed)
        self.job_list.pack(fill="both", expand=True, padx=pad_x, pady=pad_y)

        # Bottom status + generate
        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(fill="x", padx=pad_x, pady=(0, pad_y))
        ctk.CTkLabel(bottom, textvariable=self.status_var, anchor="w").pack(
            side="left", fill="x", expand=True
        )
        self.generate_btn = ctk.CTkButton(
            bottom,
            text="Generate Cover Sheets",
            width=180,
            height=32,
            font=ctk.CTkFont(weight="bold"),
            command=self._on_generate,
        )
        self.generate_btn.pack(side="right")

        ctk.CTkLabel(
            self,
            text=f"{__author__} · v{__version__}",
            text_color=("gray50", "gray55"),
            anchor="e",
            font=ctk.CTkFont(size=11),
        ).pack(anchor="e", padx=pad_x, pady=(0, 8))

    # --- Preferences -----------------------------------------------------

    def _dialog_initial_dir(self) -> str | None:
        path = self._prefs.resolved_file_dialog_dir()
        return str(path) if path is not None else None

    def _collect_preferences(self) -> AppPreferences:
        try:
            geometry = self.geometry()
        except tk.TclError:
            geometry = self._prefs.window_geometry or "900x560"
        return AppPreferences(
            last_folder=self._prefs.last_folder,
            last_file_dialog_dir=self._prefs.last_file_dialog_dir,
            output_dir=self.output_dir_var.get().strip(),
            window_geometry=geometry,
            appearance_mode=normalize_appearance_mode(self.appearance_var.get()),
            compress=self.compress_var.get(),
            force=self.force_var.get(),
            rename_to_label=self.rename_to_label_var.get(),
            strip_metadata=self.strip_metadata_var.get(),
            optimize=self.optimize_var.get(),
            linearize=self.linearize_var.get(),
            ocr=self.ocr_var.get(),
            ocr_language=self.ocr_language_var.get().strip() or "eng",
            open_when_done=self.open_when_done_var.get(),
        )

    def _save_preferences(self) -> None:
        self._prefs = self._collect_preferences()
        try:
            save_preferences(self._prefs, self._prefs_path)
        except OSError:
            pass

    def _remember_folder(self, folder: Path) -> None:
        folder = folder.expanduser().resolve()
        self._prefs.last_folder = str(folder)
        self._prefs.last_file_dialog_dir = str(folder)

    def _restore_session(self) -> None:
        if self._initial_folder_arg is not None:
            self._load_folder(self._initial_folder_arg)
            return
        last = self._prefs.resolved_last_folder()
        if last is not None:
            self._load_folder(last, remember=True, from_prefs=True)

    def _on_appearance_change(self, mode: str) -> None:
        apply_appearance(mode)
        self._save_preferences()

    # --- Job list helpers ------------------------------------------------

    def _update_hint(self) -> None:
        if self.rename_to_label_var.get():
            naming = f"Outputs are named {OUTPUT_PREFIX}CoverLabel.pdf (sanitized)"
        else:
            naming = f"Outputs are named {OUTPUT_PREFIX}OriginalName.pdf"
        self.hint_var.set(
            "Edit Cover Label in the list. "
            "Use Include checkboxes to choose files. "
            f"{naming}."
        )

    def _on_jobs_changed(self) -> None:
        jobs = self.job_list.jobs
        included = sum(1 for j in jobs if j.include)
        if jobs:
            self.status_var.set(f"{len(jobs)} file(s) loaded · {included} included")
        else:
            self.status_var.set("Add PDFs or open a folder to begin.")

    def _on_select_all(self, _event: object | None = None) -> str:
        # Don't steal Ctrl+A from text entries.
        focus = self.focus_get()
        if focus is not None:
            cls = focus.winfo_class()
            if cls in {"Entry", "Text", "TEntry", "CTkEntry"}:
                return ""
            # CustomTkinter entry internal widget
            name = str(focus).lower()
            if "entry" in name or "text" in name:
                return ""
        self.job_list.select_all()
        return "break"

    def _merge_jobs(self, new_jobs: list[JobItem], *, replace: bool) -> None:
        if replace:
            jobs = list(new_jobs)
        else:
            jobs = list(self.job_list.get_jobs())
            existing = {job.source for job in jobs}
            for job in new_jobs:
                if job.source not in existing:
                    jobs.append(job)
                    existing.add(job.source)
            jobs.sort(key=lambda j: j.source.name.casefold())
        self.job_list.set_jobs(jobs)

    def _load_folder(
        self,
        folder: Path,
        *,
        remember: bool = True,
        from_prefs: bool = False,
    ) -> None:
        folder = folder.expanduser().resolve()
        if not folder.is_dir():
            if not from_prefs:
                messagebox.showerror("Invalid folder", f"Not a directory:\n{folder}")
            return
        jobs = jobs_from_folder(folder)
        if not jobs and not from_prefs:
            messagebox.showinfo(
                "No PDFs",
                f"No input PDFs found in:\n{folder}\n\n"
                f"(Files starting with '{OUTPUT_PREFIX}' are skipped.)",
            )
        self._merge_jobs(jobs, replace=True)
        if not self.output_dir_var.get().strip():
            self.output_dir_var.set(str(folder))
        if remember:
            self._remember_folder(folder)
            self._save_preferences()
        if from_prefs and jobs:
            included = sum(1 for j in self.job_list.jobs if j.include)
            self.status_var.set(
                f"Restored last folder · {len(self.job_list.jobs)} file(s) · "
                f"{included} included"
            )

    # --- Toolbar actions -------------------------------------------------

    def _on_open_folder(self) -> None:
        if self._busy:
            return
        kwargs: dict[str, Any] = {"title": "Select a folder containing PDFs"}
        initial = self._dialog_initial_dir()
        if initial:
            kwargs["initialdir"] = initial
        path = filedialog.askdirectory(**kwargs)
        if path:
            self._load_folder(Path(path))

    def _on_add_pdfs(self) -> None:
        if self._busy:
            return
        kwargs: dict[str, Any] = {
            "title": "Select PDF files",
            "filetypes": [("PDF files", "*.pdf"), ("All files", "*.*")],
        }
        initial = self._dialog_initial_dir()
        if initial:
            kwargs["initialdir"] = initial
        paths = filedialog.askopenfilenames(**kwargs)
        if not paths:
            return
        filtered = [
            Path(p) for p in paths if not is_output_filename(Path(p).name)
        ]
        skipped = len(paths) - len(filtered)
        jobs = jobs_from_paths(filtered)
        self._merge_jobs(jobs, replace=False)
        if jobs:
            parent = jobs[0].source.parent
            self._prefs.last_file_dialog_dir = str(parent)
            if not self._prefs.last_folder:
                self._prefs.last_folder = str(parent)
            self._save_preferences()
        if skipped:
            messagebox.showinfo(
                "Skipped outputs",
                f"Skipped {skipped} file(s) that look like prior outputs "
                f"(name starts with '{OUTPUT_PREFIX}').",
            )

    def _on_remove(self) -> None:
        if self._busy:
            return
        self.job_list.remove_selected()

    def _on_remove_all(self) -> None:
        if self._busy:
            return
        if not self.job_list.jobs:
            return
        if not messagebox.askyesno(
            "Remove all?",
            f"Remove all {len(self.job_list.jobs)} PDF(s) from the list?\n\n"
            "This does not delete any files on disk.",
            parent=self,
        ):
            return
        self.job_list.remove_all()

    def _on_reset_labels(self) -> None:
        if self._busy:
            return
        self.job_list.reset_labels()

    def _on_browse_output(self) -> None:
        kwargs: dict[str, Any] = {"title": "Select output folder"}
        out = self.output_dir_var.get().strip()
        if out and Path(out).expanduser().is_dir():
            kwargs["initialdir"] = str(Path(out).expanduser())
        else:
            initial = self._dialog_initial_dir()
            if initial:
                kwargs["initialdir"] = initial
        path = filedialog.askdirectory(**kwargs)
        if path:
            self.output_dir_var.set(path)
            self._save_preferences()

    # --- Generate --------------------------------------------------------

    def _on_generate(self) -> None:
        if self._busy:
            return
        jobs = self.job_list.get_jobs()
        included = [j for j in jobs if j.include]
        if not included:
            messagebox.showwarning(
                "Nothing to process",
                "Include at least one PDF in the list.",
            )
            return

        out_raw = self.output_dir_var.get().strip()
        output_dir: Path | None = Path(out_raw).expanduser() if out_raw else None
        if output_dir is not None:
            try:
                output_dir = output_dir.resolve()
                output_dir.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                messagebox.showerror("Output folder", str(exc))
                return

        jobs_snapshot = [
            JobItem(source=j.source, label=j.label, include=j.include) for j in jobs
        ]
        options = ProcessOptions(
            compress=self.compress_var.get(),
            force=self.force_var.get(),
            dry_run=False,
            rename_to_label=self.rename_to_label_var.get(),
            strip_metadata=self.strip_metadata_var.get(),
            ocr=self.ocr_var.get() and ocr_available(),
            ocr_language=self.ocr_language_var.get().strip() or "eng",
            ocr_skip_text=True,
            optimize=self.optimize_var.get(),
            linearize=self.linearize_var.get() and linearize_available(),
        )
        open_when_done = self.open_when_done_var.get()

        if self.ocr_var.get() and not ocr_available():
            messagebox.showerror(
                "OCR unavailable",
                "OCR requires ocrmypdf and Tesseract.\n\n"
                "• Windows: use the full installer release (recommended), or\n"
                "• pip install 'coversheets[ocr]' and install Tesseract.",
                parent=self,
            )
            return
        if self.linearize_var.get() and not linearize_available():
            messagebox.showerror(
                "Linearize unavailable",
                "Linearize requires pikepdf or the qpdf CLI.\n"
                "Install with: pip install 'coversheets[optimize]'",
                parent=self,
            )
            return

        self._save_preferences()

        self._run_output_dir = output_dir
        self._run_jobs = jobs_snapshot
        self._cancel_event.clear()
        self._set_busy(True)
        self.status_var.set("Generating…")

        self._progress_win = ProgressWindow(
            self,
            total=len(included),
            on_cancel=self._cancel_event.set,
        )
        self._progress_win.append_log(f"Processing {len(included)} PDF(s)…")
        for line in options.describe():
            self._progress_win.append_log(f"Option: {line}")

        def worker() -> None:
            def log(msg: str) -> None:
                self._msg_queue.put(("log", msg))

            def progress(current: int, total: int, name: str) -> None:
                self._msg_queue.put(("progress", (current, total, name)))

            try:
                result = process_jobs(
                    jobs_snapshot,
                    output_dir=output_dir,
                    options=options,
                    progress=progress,
                    log=log,
                    cancel_check=self._cancel_event.is_set,
                )
                self._msg_queue.put(("done", (result, open_when_done)))
            except Exception as exc:  # pragma: no cover - defensive
                self._msg_queue.put(("error", str(exc)))

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        state = "disabled" if busy else "normal"
        self.generate_btn.configure(state=state)
        for btn in self._toolbar_buttons:
            try:
                btn.configure(state=state)
            except tk.TclError:
                pass
        try:
            self.appearance_menu.configure(state=state)
        except tk.TclError:
            pass

    def _close_progress_window(self) -> None:
        win = self._progress_win
        self._progress_win = None
        if win is not None:
            try:
                win.grab_release()
            except tk.TclError:
                pass
            try:
                win.destroy()
            except tk.TclError:
                pass

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self._msg_queue.get_nowait()
                if kind == "progress":
                    current, total, name = payload
                    if self._progress_win is not None:
                        self._progress_win.set_progress(current, total, name)
                    self.status_var.set(f"[{current}/{total}] {name}")
                elif kind == "log":
                    if self._progress_win is not None:
                        self._progress_win.append_log(str(payload))
                elif kind == "done":
                    result, open_when_done = payload
                    self._on_done(result, open_when_done=open_when_done)
                elif kind == "error":
                    self._close_progress_window()
                    self._set_busy(False)
                    messagebox.showerror("Error", str(payload), parent=self)
                    self.status_var.set("Failed.")
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _on_done(self, result: BatchResult, *, open_when_done: bool) -> None:
        if self._progress_win is not None:
            self._progress_win.mark_complete(cancelled=result.was_cancelled)
            self.after(
                250,
                lambda: self._finish_done(result, open_when_done=open_when_done),
            )
        else:
            self._finish_done(result, open_when_done=open_when_done)

    def _finish_done(self, result: BatchResult, *, open_when_done: bool) -> None:
        self._close_progress_window()
        self._set_busy(False)
        summary = format_result_summary(result)
        self.status_var.set(summary)

        folders = resolve_result_folders(self._run_jobs, self._run_output_dir)
        DoneDialog(
            self,
            result=result,
            folders=folders,
            open_when_done=open_when_done,
        )

    def _on_close(self) -> None:
        if self._busy:
            if not messagebox.askyesno(
                "Busy",
                "Generation is still running.\n\n"
                "Quit anyway? The current file may finish writing, "
                "but remaining files will be cancelled.",
                parent=self,
            ):
                return
            self._cancel_event.set()
        self._save_preferences()
        self._close_progress_window()
        self.destroy()


def run_app(initial_folder: Path | None = None) -> int:
    """Launch the GUI and block until the window closes."""
    app = CoverSheetsApp(initial_folder=initial_folder)
    app.mainloop()
    return 0
