"""Main CustomTkinter window: guided list UI for non-technical users."""

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
from coversheets.gui.copy import (
    STEP_LABELS,
    empty_list_blurb,
    plain_option_lines,
    preview_target_index,
    status_for_jobs,
)
from coversheets.gui.dialogs import DoneDialog, ProgressWindow
from coversheets.gui.dnd import dnd_available, make_dnd_root, register_drop_target
from coversheets.gui.job_list import JobListFrame
from coversheets.gui.options_panel import OptionsPanel
from coversheets.gui.preview import CoverPreview
from coversheets.gui.theme import APPEARANCE_MODES, apply_appearance, normalize_appearance_mode
from coversheets.gui.welcome import WelcomeDialog
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
from coversheets.util import format_result_summary, resolve_app_asset, resolve_result_folders


class _CoverSheetsAppBase(ctk.CTk):
    """Base CTk window (may be mixed with Tk DnD)."""


CoverSheetsAppRoot = make_dnd_root(_CoverSheetsAppBase)


def _apply_window_icon(window: ctk.CTk) -> None:
    """Set taskbar/title-bar icon when assets are present; never raise."""
    try:
        ico = resolve_app_asset("app-icon.ico")
        png = resolve_app_asset("app-icon.png")
        # Windows: .ico is most reliable for the taskbar/title bar.
        if sys.platform == "win32" and ico is not None:
            try:
                window.iconbitmap(default=str(ico))
            except tk.TclError:
                pass
        if png is not None:
            # Keep a reference so Tk does not garbage-collect the image.
            photo = tk.PhotoImage(file=str(png))
            window.iconphoto(True, photo)
            window._app_icon_photo = photo  # type: ignore[attr-defined]
        elif ico is not None and sys.platform != "win32":
            try:
                window.iconbitmap(str(ico))
            except tk.TclError:
                pass
    except Exception:
        # Icon is cosmetic — never block startup.
        return


class CoverSheetsApp(CoverSheetsAppRoot):  # type: ignore[misc, valid-type]
    """Main window: editable list of PDFs and cover titles."""

    def __init__(self, initial_folder: Path | None = None) -> None:
        apply_appearance("System")
        super().__init__()
        configure_bundled_tools()
        self.title(f"Automatic Exhibit Cover Sheets v{__version__}")
        self.minsize(920, 560)
        _apply_window_icon(self)

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
        self._dnd_enabled = False

        self.appearance_var = ctk.StringVar(
            value=normalize_appearance_mode(self._prefs.appearance_mode)
        )
        self.status_var = ctk.StringVar(value=status_for_jobs(0, 0))

        geo = (self._prefs.window_geometry or "1000x640").strip() or "1000x640"
        try:
            self.geometry(geo)
        except tk.TclError:
            self.geometry("1000x640")

        self._build_ui()
        self._setup_dnd()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Control-a>", self._on_select_all)
        self.bind("<Control-A>", self._on_select_all)
        if sys.platform == "darwin":
            self.bind("<Command-a>", self._on_select_all)
            self.bind("<Command-A>", self._on_select_all)
        self.after(100, self._poll_queue)
        self.after(80, self._restore_session)
        self.after(200, self._maybe_show_welcome)

    # --- UI construction -------------------------------------------------

    def _build_ui(self) -> None:
        pad_x, pad_y = 12, 6

        # Title row: steps + theme + help
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=pad_x, pady=(pad_y, 2))

        self._step_labels: list[ctk.CTkLabel] = []
        steps_frame = ctk.CTkFrame(top, fg_color="transparent")
        steps_frame.pack(side="left", fill="x", expand=True)
        for i, text in enumerate(STEP_LABELS):
            lbl = ctk.CTkLabel(
                steps_frame,
                text=text,
                font=ctk.CTkFont(size=13, weight="bold"),
                text_color=("gray50", "gray55"),
            )
            lbl.pack(side="left", padx=(0, 12))
            self._step_labels.append(lbl)
            if i < len(STEP_LABELS) - 1:
                ctk.CTkLabel(
                    steps_frame,
                    text="→",
                    text_color=("gray60", "gray50"),
                ).pack(side="left", padx=(0, 12))

        right_top = ctk.CTkFrame(top, fg_color="transparent")
        right_top.pack(side="right")
        ctk.CTkButton(
            right_top,
            text="?",
            width=32,
            command=self._show_welcome,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkLabel(right_top, text="Theme").pack(side="left", padx=(0, 6))
        self.appearance_menu = ctk.CTkOptionMenu(
            right_top,
            values=list(APPEARANCE_MODES),
            variable=self.appearance_var,
            width=100,
            command=self._on_appearance_change,
        )
        self.appearance_menu.pack(side="left")

        # Toolbar
        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill="x", padx=pad_x, pady=(2, 2))
        for text, cmd in (
            ("Open folder…", self._on_open_folder),
            ("Add PDFs…", self._on_add_pdfs),
            ("Remove selected", self._on_remove),
            ("Clear list", self._on_remove_all),
            ("Reset titles", self._on_reset_labels),
            ("Include all", self._on_include_all),
            ("Exclude all", self._on_exclude_all),
        ):
            btn = ctk.CTkButton(toolbar, text=text, width=0, command=cmd)
            btn.pack(side="left", padx=(0, 6))
            self._toolbar_buttons.append(btn)

        # Main split: list + preview
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=pad_x, pady=pad_y)
        body.grid_columnconfigure(0, weight=3)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        list_wrap = ctk.CTkFrame(body, fg_color="transparent")
        list_wrap.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        list_wrap.grid_columnconfigure(0, weight=1)
        list_wrap.grid_rowconfigure(0, weight=1)

        blurb = empty_list_blurb(dnd_available=dnd_available())
        self.job_list = JobListFrame(
            list_wrap,
            on_jobs_changed=self._on_jobs_changed,
            on_selection_changed=self._on_selection_changed,
            on_empty_open_folder=self._on_open_folder,
            on_empty_add_pdfs=self._on_add_pdfs,
            empty_blurb=blurb,
        )
        self.job_list.grid(row=0, column=0, sticky="nsew")

        self.preview = CoverPreview(body, width=240)
        self.preview.grid(row=0, column=1, sticky="nsew")

        # Options
        self.options = OptionsPanel(
            self,
            self._prefs,
            on_change=self._on_options_changed,
            dialog_initial_dir=self._dialog_initial_dir,
        )
        self.options.pack(fill="x", padx=pad_x, pady=(0, pad_y))
        self._sync_preview_position()

        # Bottom status + generate
        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(fill="x", padx=pad_x, pady=(0, pad_y))
        ctk.CTkLabel(bottom, textvariable=self.status_var, anchor="w").pack(
            side="left", fill="x", expand=True
        )
        self.generate_btn = ctk.CTkButton(
            bottom,
            text="Generate cover sheets",
            width=200,
            height=36,
            font=ctk.CTkFont(size=14, weight="bold"),
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

        self._update_steps()

    def _setup_dnd(self) -> None:
        if not dnd_available():
            return
        self._dnd_enabled = register_drop_target(self, self._on_drop_paths)
        if self._dnd_enabled:
            self.job_list.set_empty_blurb(empty_list_blurb(dnd_available=True))

    # --- Preferences -----------------------------------------------------

    def _dialog_initial_dir(self) -> str | None:
        path = self._prefs.resolved_file_dialog_dir()
        return str(path) if path is not None else None

    def _collect_preferences(self) -> AppPreferences:
        try:
            geometry = self.geometry()
        except tk.TclError:
            geometry = self._prefs.window_geometry or "1000x640"
        base = self.options.collect_into_prefs(self._prefs)
        base.window_geometry = geometry
        base.appearance_mode = normalize_appearance_mode(self.appearance_var.get())
        base.last_folder = self._prefs.last_folder
        base.last_file_dialog_dir = self._prefs.last_file_dialog_dir
        base.show_welcome = self._prefs.show_welcome
        return base

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

    def _maybe_show_welcome(self) -> None:
        if self._prefs.show_welcome:
            self._show_welcome()

    def _show_welcome(self) -> None:
        def finish(dont_show: bool) -> None:
            if dont_show:
                self._prefs.show_welcome = False
                self._save_preferences()

        WelcomeDialog(self, on_finish=finish)

    def _on_appearance_change(self, mode: str) -> None:
        apply_appearance(mode)
        self._save_preferences()

    def _on_options_changed(self) -> None:
        self._update_steps()
        self._sync_preview_position()
        # Auto-save lightly so advanced/output mode stick
        try:
            self._prefs = self._collect_preferences()
        except Exception:
            pass

    def _sync_preview_position(self) -> None:
        try:
            self.preview.set_vertical_position(self.options.vertical_position_id())
        except Exception:
            pass

    # --- Steps / preview -------------------------------------------------

    def _update_steps(self) -> None:
        jobs = self.job_list.jobs
        included = sum(1 for j in jobs if j.include)
        if not jobs:
            active = 0
        elif included == 0:
            active = 1
        else:
            active = 2
        for i, lbl in enumerate(self._step_labels):
            if i == active:
                lbl.configure(text_color=("#1F6AA5", "#3B8ED0"))
            elif i < active:
                lbl.configure(text_color=("gray30", "gray70"))
            else:
                lbl.configure(text_color=("gray50", "gray55"))

    def _on_selection_changed(self) -> None:
        jobs = self.job_list.jobs
        idx = preview_target_index(
            self.job_list.selected_set,
            anchor=self.job_list.anchor_index,
            job_count=len(jobs),
        )
        if idx is None:
            if not jobs:
                self.preview.clear()
            else:
                self.preview.clear("Select a file to preview its cover.")
            return
        job = jobs[idx]
        self.preview.show(title=job.label, filename=job.source.name)

    def _on_jobs_changed(self) -> None:
        jobs = self.job_list.jobs
        included = sum(1 for j in jobs if j.include)
        self.status_var.set(status_for_jobs(len(jobs), included))
        self._update_steps()
        self._on_selection_changed()

    def _on_select_all(self, _event: object | None = None) -> str:
        focus = self.focus_get()
        if focus is not None:
            cls = focus.winfo_class()
            if cls in {"Entry", "Text", "TEntry", "CTkEntry"}:
                return ""
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
                "No PDFs found",
                f"No PDFs found in:\n{folder}\n\n"
                f"(Files starting with “{OUTPUT_PREFIX}” are skipped — "
                "those look like previous outputs.)",
            )
        self._merge_jobs(jobs, replace=True)
        # Default single-folder mode to this folder if user chose "one folder"
        if (
            self.options.output_mode_var.get() == "folder"
            and not self.options.output_dir_var.get().strip()
        ):
            self.options.output_dir_var.set(str(folder))
        if remember:
            self._remember_folder(folder)
            self._save_preferences()
        if from_prefs and jobs:
            included = sum(1 for j in self.job_list.jobs if j.include)
            self.status_var.set(
                f"Restored last folder · {status_for_jobs(len(self.job_list.jobs), included)}"
            )

    def _on_drop_paths(self, paths: list[Path] | Any) -> None:
        if self._busy:
            return
        path_list = [Path(p) for p in paths]
        folders = [p for p in path_list if p.is_dir()]
        files = [p for p in path_list if p.is_file()]
        if folders:
            # First folder wins for replace-load (same as Open folder).
            self._load_folder(folders[0])
            return
        if not files:
            return
        filtered = [p for p in files if p.suffix.lower() == ".pdf"]
        filtered = [p for p in filtered if not is_output_filename(p.name)]
        if not filtered:
            messagebox.showinfo(
                "Nothing to add",
                "Drop PDF files (not folders that already start with "
                f"“{OUTPUT_PREFIX}”).",
                parent=self,
            )
            return
        jobs = jobs_from_paths(filtered)
        self._merge_jobs(jobs, replace=False)
        if jobs:
            parent = jobs[0].source.parent
            self._prefs.last_file_dialog_dir = str(parent)
            if not self._prefs.last_folder:
                self._prefs.last_folder = str(parent)
            self._save_preferences()

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
                "Skipped previous outputs",
                f"Skipped {skipped} file(s) that look like prior outputs "
                f"(name starts with “{OUTPUT_PREFIX}”).",
            )

    def _on_remove(self) -> None:
        if self._busy:
            return
        if not self.job_list.selected_indices():
            messagebox.showinfo(
                "Nothing selected",
                "Click a row to select it, then Remove selected.\n\n"
                "Tip: hold Ctrl (⌘ on Mac) to select more than one.",
                parent=self,
            )
            return
        self.job_list.remove_selected()

    def _on_remove_all(self) -> None:
        if self._busy:
            return
        if not self.job_list.jobs:
            return
        if not messagebox.askyesno(
            "Clear the list?",
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

    def _on_include_all(self) -> None:
        if self._busy:
            return
        self.job_list.include_all()

    def _on_exclude_all(self) -> None:
        if self._busy:
            return
        self.job_list.exclude_all()

    # --- Generate --------------------------------------------------------

    def _on_generate(self) -> None:
        if self._busy:
            return
        jobs = self.job_list.get_jobs()
        included = [j for j in jobs if j.include]
        if not included:
            messagebox.showwarning(
                "Nothing to process",
                "Include at least one PDF.\n\n"
                "Use the Include checkbox on each row, or click Include all.",
                parent=self,
            )
            return

        cap_err = self.options.ensure_capabilities_or_warn()
        if cap_err:
            messagebox.showerror("Option not available", cap_err, parent=self)
            return

        output_dir = self.options.resolved_output_dir()
        if self.options.output_mode_var.get() == "folder":
            if output_dir is None:
                messagebox.showerror(
                    "Choose a folder",
                    "You chose “One folder…” — pick where finished PDFs should go.",
                    parent=self,
                )
                return
            try:
                output_dir.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                messagebox.showerror("Output folder", str(exc), parent=self)
                return

        jobs_snapshot = [
            JobItem(source=j.source, label=j.label, include=j.include) for j in jobs
        ]
        options = ProcessOptions(
            compress=self.options.compress_var.get(),
            force=self.options.force_var.get(),
            dry_run=False,
            rename_to_label=self.options.rename_to_label_var.get(),
            strip_metadata=self.options.strip_metadata_var.get(),
            ocr=self.options.ocr_var.get() and ocr_available(),
            ocr_language=self.options.ocr_language_var.get().strip() or "eng",
            ocr_skip_text=True,
            optimize=self.options.optimize_var.get(),
            linearize=self.options.linearize_var.get() and linearize_available(),
            vertical_position=self.options.vertical_position_id(),
        )
        open_when_done = self.options.open_when_done_var.get()

        self._save_preferences()

        self._run_output_dir = output_dir
        self._run_jobs = jobs_snapshot
        self._cancel_event.clear()
        self._set_busy(True)
        self.status_var.set("Creating cover sheets…")

        self._progress_win = ProgressWindow(
            self,
            total=len(included),
            on_cancel=self._cancel_event.set,
        )
        self._progress_win.append_log(
            f"Creating cover sheets for {len(included)} PDF(s)…"
        )
        for line in plain_option_lines(options):
            self._progress_win.append_log(f"• {line}")

        def worker() -> None:
            def log(msg: str) -> None:
                # Soften a few pipeline messages for the progress log.
                text = msg
                if text.startswith("Writing "):
                    text = text.replace("Writing ", "Saving ", 1)
                self._msg_queue.put(("log", text))

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
                    messagebox.showerror("Something went wrong", str(payload), parent=self)
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
                "Still working",
                "Cover sheets are still being created.\n\n"
                "Quit anyway? The current file may finish writing, "
                "but remaining files will stop.",
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
