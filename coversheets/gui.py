"""GUI list: load PDFs, edit cover labels, generate outputs."""

from __future__ import annotations

import queue
import sys
import threading
import tkinter as tk
from collections.abc import Callable
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Sequence

from coversheets import OUTPUT_PREFIX, __author__, __version__
from coversheets.bundled_tools import configure_bundled_tools
from coversheets.cover import cover_label_from_filename
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
from coversheets.util import (
    format_result_summary,
    open_in_file_manager,
    resolve_result_folders,
)


class ProgressWindow(tk.Toplevel):
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
        self.status_var = tk.StringVar(value="Starting…")
        self.count_var = tk.StringVar(value=f"0 / {total}")
        self.progress_var = tk.DoubleVar(value=0.0)

        pad = {"padx": 12, "pady": 6}

        ttk.Label(
            self,
            text="Please wait while cover sheets are generated.",
            font=("", 11),
        ).pack(anchor=tk.W, **pad)

        ttk.Label(self, textvariable=self.status_var).pack(anchor=tk.W, padx=12)

        bar_row = ttk.Frame(self)
        bar_row.pack(fill=tk.X, padx=12, pady=8)
        self.bar = ttk.Progressbar(
            bar_row,
            variable=self.progress_var,
            maximum=100,
            mode="determinate",
        )
        self.bar.pack(fill=tk.X, side=tk.LEFT, expand=True)
        ttk.Label(bar_row, textvariable=self.count_var, width=10).pack(
            side=tk.RIGHT, padx=(8, 0)
        )

        log_frame = ttk.LabelFrame(self, text="Log")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=6)

        self.log = tk.Text(
            log_frame,
            height=10,
            wrap=tk.WORD,
            state=tk.DISABLED,
            font=("Menlo", 10) if sys.platform == "darwin" else ("Consolas", 9),
        )
        yscroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log.yview)
        self.log.configure(yscrollcommand=yscroll.set)
        self.log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(6, 0), pady=6)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y, pady=6, padx=(0, 6))

        btn_row = ttk.Frame(self)
        btn_row.pack(fill=tk.X, padx=12, pady=(0, 12))
        self.cancel_btn = ttk.Button(
            btn_row, text="Cancel", command=self._request_cancel
        )
        self.cancel_btn.pack(side=tk.RIGHT)
        ttk.Label(
            btn_row,
            text="Cancel stops after the current file.",
            foreground="#555",
        ).pack(side=tk.LEFT)

        self._center_on_master(master)
        self.grab_set()
        self.focus_set()

    def _center_on_master(self, master: tk.Misc) -> None:
        self.update_idletasks()
        try:
            mx = master.winfo_rootx()
            my = master.winfo_rooty()
            mw = master.winfo_width()
            mh = master.winfo_height()
            w = self.winfo_width()
            h = self.winfo_height()
            self.geometry(f"+{mx + (mw - w) // 2}+{my + (mh - h) // 2}")
        except tk.TclError:
            pass

    def _request_cancel(self) -> None:
        if self._cancel_requested:
            return
        self._cancel_requested = True
        self.cancel_btn.configure(state=tk.DISABLED, text="Cancelling…")
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
        pct = ((current - 1) / total) * 100 if current > 0 else 0
        self.progress_var.set(min(max(pct, 0), 100))
        self.count_var.set(f"{current} / {total}")
        if not self._cancel_requested:
            self.status_var.set(f"Processing: {name}")

    def append_log(self, line: str) -> None:
        if not line:
            return
        self.log.configure(state=tk.NORMAL)
        self.log.insert(tk.END, line.rstrip() + "\n")
        self.log.see(tk.END)
        self.log.configure(state=tk.DISABLED)

    def mark_complete(self, *, cancelled: bool = False) -> None:
        self.progress_var.set(100 if not cancelled else self.progress_var.get())
        self.count_var.set(f"{self._total} / {self._total}")
        self.status_var.set("Cancelled." if cancelled else "Finished.")
        try:
            self.cancel_btn.configure(state=tk.DISABLED)
        except tk.TclError:
            pass


class DoneDialog(tk.Toplevel):
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
        self._result_open = False

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

        ttk.Label(self, text=headline, font=("", 12, "bold")).pack(
            anchor=tk.W, padx=16, pady=(16, 4)
        )
        ttk.Label(self, text=format_result_summary(result)).pack(anchor=tk.W, **pad)

        if result.errors:
            err_frame = ttk.LabelFrame(self, text="Errors")
            err_frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=6)
            err_text = tk.Text(err_frame, height=min(6, max(2, len(result.errors))), wrap=tk.WORD)
            err_text.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
            for name, msg in result.errors:
                err_text.insert(tk.END, f"• {name}: {msg}\n")
            err_text.configure(state=tk.DISABLED)

        if self.folders:
            folder_label = (
                str(self.folders[0])
                if len(self.folders) == 1
                else f"{len(self.folders)} folders (outputs next to sources)"
            )
            ttk.Label(self, text=f"Output: {folder_label}", foreground="#444").pack(
                anchor=tk.W, padx=16, pady=(4, 8)
            )

        btn_row = ttk.Frame(self)
        btn_row.pack(fill=tk.X, padx=16, pady=(4, 16))

        if self.folders:
            open_label = (
                "Open Output Folder"
                if len(self.folders) == 1
                else "Open Output Folders"
            )
            ttk.Button(btn_row, text=open_label, command=self._open_folders).pack(
                side=tk.LEFT
            )

        ttk.Button(btn_row, text="Close", command=self._close).pack(side=tk.RIGHT)

        self.protocol("WM_DELETE_WINDOW", self._close)
        self.grab_set()
        self.focus_set()
        self._center_on_master(master)

        if open_when_done and self.folders and ok:
            # Defer so the dialog is visible first.
            self.after(150, self._open_folders)

    def _center_on_master(self, master: tk.Misc) -> None:
        self.update_idletasks()
        try:
            mx = master.winfo_rootx()
            my = master.winfo_rooty()
            mw = master.winfo_width()
            mh = master.winfo_height()
            w = self.winfo_reqwidth()
            h = self.winfo_reqheight()
            self.geometry(f"+{mx + (mw - w) // 2}+{my + (mh - h) // 2}")
        except tk.TclError:
            pass

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
        self.grab_release()
        self.destroy()


class CoverSheetsApp(tk.Tk):
    """Main window: editable table of PDFs and cover labels."""

    def __init__(self, initial_folder: Path | None = None) -> None:
        super().__init__()
        configure_bundled_tools()
        self.title(f"Automatic Exhibit Cover Sheets v{__version__}")
        self.minsize(780, 480)

        self._prefs = load_preferences()
        self._prefs_path: Path | None = None  # use default path
        self._initial_folder_arg = initial_folder

        self.jobs: list[JobItem] = []
        self._iid_to_index: dict[str, int] = {}
        self._edit_entry: ttk.Entry | None = None
        self._edit_iid: str | None = None
        self._edit_column: str | None = None
        self._worker: threading.Thread | None = None
        self._msg_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self._busy = False
        self._cancel_event = threading.Event()
        self._progress_win: ProgressWindow | None = None
        self._run_output_dir: Path | None = None
        self._run_jobs: list[JobItem] = []

        p = self._prefs
        self.output_dir_var = tk.StringVar(value=p.output_dir or "")
        self.compress_var = tk.BooleanVar(value=p.compress)
        self.force_var = tk.BooleanVar(value=p.force)
        self.rename_to_label_var = tk.BooleanVar(value=p.rename_to_label)
        self.strip_metadata_var = tk.BooleanVar(value=p.strip_metadata)
        self.optimize_var = tk.BooleanVar(value=p.optimize)
        self.linearize_var = tk.BooleanVar(value=p.linearize)
        self.ocr_var = tk.BooleanVar(value=p.ocr)
        self.ocr_language_var = tk.StringVar(value=p.ocr_language or "eng")
        self.open_when_done_var = tk.BooleanVar(value=p.open_when_done)
        self.status_var = tk.StringVar(value="Add PDFs or open a folder to begin.")

        geo = (p.window_geometry or "900x560").strip() or "900x560"
        try:
            self.geometry(geo)
        except tk.TclError:
            self.geometry("900x560")

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._poll_queue)
        self.after(80, self._restore_session)

    # --- UI construction -------------------------------------------------

    def _build_ui(self) -> None:
        pad = {"padx": 8, "pady": 6}

        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X, **pad)

        ttk.Button(toolbar, text="Open Folder…", command=self._on_open_folder).pack(
            side=tk.LEFT, padx=(0, 4)
        )
        ttk.Button(toolbar, text="Add PDFs…", command=self._on_add_pdfs).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(toolbar, text="Remove Selected", command=self._on_remove).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(toolbar, text="Reset Labels", command=self._on_reset_labels).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(toolbar, text="Toggle Include", command=self._on_toggle_include).pack(
            side=tk.LEFT, padx=4
        )

        out_row = ttk.Frame(self)
        out_row.pack(fill=tk.X, **pad)
        ttk.Label(out_row, text="Output folder:").pack(side=tk.LEFT)
        ttk.Entry(out_row, textvariable=self.output_dir_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=6
        )
        ttk.Button(out_row, text="Browse…", command=self._on_browse_output).pack(
            side=tk.LEFT
        )
        ttk.Label(
            out_row,
            text="(empty = next to each source)",
            foreground="#555",
        ).pack(side=tk.LEFT, padx=(8, 0))

        opts = ttk.Frame(self)
        opts.pack(fill=tk.X, **pad)
        ttk.Checkbutton(
            opts, text="Compress page streams", variable=self.compress_var
        ).pack(side=tk.LEFT)
        ttk.Checkbutton(
            opts, text="Overwrite existing +outputs", variable=self.force_var
        ).pack(side=tk.LEFT, padx=(16, 0))
        ttk.Checkbutton(
            opts,
            text="Name output file after label",
            variable=self.rename_to_label_var,
            command=self._update_hint,
        ).pack(side=tk.LEFT, padx=(16, 0))

        pdf_opts = ttk.LabelFrame(self, text="PDF options")
        pdf_opts.pack(fill=tk.X, padx=8, pady=(0, 6))

        pdf_row1 = ttk.Frame(pdf_opts)
        pdf_row1.pack(fill=tk.X, padx=8, pady=4)
        ttk.Checkbutton(
            pdf_row1,
            text="Strip metadata (Info + XMP)",
            variable=self.strip_metadata_var,
        ).pack(side=tk.LEFT)
        ttk.Checkbutton(
            pdf_row1,
            text="Optimize (dedupe objects)",
            variable=self.optimize_var,
        ).pack(side=tk.LEFT, padx=(16, 0))
        lin_state = tk.NORMAL if linearize_available() else tk.DISABLED
        self._linearize_cb = ttk.Checkbutton(
            pdf_row1,
            text="Linearize (web optimize)",
            variable=self.linearize_var,
            state=lin_state,
        )
        self._linearize_cb.pack(side=tk.LEFT, padx=(16, 0))

        pdf_row2 = ttk.Frame(pdf_opts)
        pdf_row2.pack(fill=tk.X, padx=8, pady=(0, 6))
        ocr_state = tk.NORMAL if ocr_available() else tk.DISABLED
        self._ocr_cb = ttk.Checkbutton(
            pdf_row2,
            text="OCR (ocrmypdf)",
            variable=self.ocr_var,
            state=ocr_state,
        )
        self._ocr_cb.pack(side=tk.LEFT)
        ttk.Label(pdf_row2, text="Language:").pack(side=tk.LEFT, padx=(12, 4))
        self._ocr_lang_entry = ttk.Entry(
            pdf_row2, textvariable=self.ocr_language_var, width=10
        )
        self._ocr_lang_entry.pack(side=tk.LEFT)
        if not ocr_available():
            ttk.Label(
                pdf_row2,
                text="(use Windows full installer, or coversheets[ocr] + Tesseract)",
                foreground="#888",
            ).pack(side=tk.LEFT, padx=(8, 0))
        elif not linearize_available():
            pass
        if not linearize_available():
            ttk.Label(
                pdf_row2,
                text="Linearize needs coversheets[optimize] or qpdf",
                foreground="#888",
            ).pack(side=tk.LEFT, padx=(12, 0))

        opts2 = ttk.Frame(self)
        opts2.pack(fill=tk.X, padx=8, pady=(0, 6))
        ttk.Checkbutton(
            opts2,
            text="Open output folder when done",
            variable=self.open_when_done_var,
        ).pack(side=tk.LEFT)

        self.hint_var = tk.StringVar()
        hint = ttk.Label(self, textvariable=self.hint_var, foreground="#444")
        hint.pack(fill=tk.X, padx=8)
        self._update_hint()

        table_frame = ttk.Frame(self)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

        columns = ("include", "file", "label", "folder")
        self.tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            selectmode="extended",
        )
        self.tree.heading("include", text="Include")
        self.tree.heading("file", text="File")
        self.tree.heading("label", text="Cover Label")
        self.tree.heading("folder", text="Folder")
        self.tree.column("include", width=70, anchor=tk.CENTER, stretch=False)
        self.tree.column("file", width=200, anchor=tk.W)
        self.tree.column("label", width=260, anchor=tk.W)
        self.tree.column("folder", width=280, anchor=tk.W)

        yscroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<Button-1>", self._on_tree_click, add="+")

        bottom = ttk.Frame(self)
        bottom.pack(fill=tk.X, **pad)

        status_row = ttk.Frame(bottom)
        status_row.pack(fill=tk.X)
        ttk.Label(status_row, textvariable=self.status_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True
        )
        self.generate_btn = ttk.Button(
            status_row, text="Generate Cover Sheets", command=self._on_generate
        )
        self.generate_btn.pack(side=tk.RIGHT)

        footer = ttk.Label(
            self,
            text=f"{__author__} · v{__version__}",
            foreground="#777",
            font=("", 9),
        )
        footer.pack(anchor=tk.E, padx=8, pady=(0, 6))

    # --- Preferences -----------------------------------------------------

    def _dialog_initial_dir(self) -> str | None:
        """Directory to open file/folder pickers in."""
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
            # Non-fatal: disk full / permissions.
            pass

    def _remember_folder(self, folder: Path) -> None:
        folder = folder.expanduser().resolve()
        self._prefs.last_folder = str(folder)
        self._prefs.last_file_dialog_dir = str(folder)

    def _restore_session(self) -> None:
        """After UI is up: reload last folder unless CLI provided one."""
        if self._initial_folder_arg is not None:
            self._load_folder(self._initial_folder_arg)
            return
        last = self._prefs.resolved_last_folder()
        if last is not None:
            self._load_folder(last, remember=True, from_prefs=True)

    # --- Job list helpers ------------------------------------------------

    def _update_hint(self) -> None:
        if self.rename_to_label_var.get():
            naming = f"Outputs are named {OUTPUT_PREFIX}CoverLabel.pdf (sanitized)"
        else:
            naming = f"Outputs are named {OUTPUT_PREFIX}OriginalName.pdf"
        self.hint_var.set(
            "Double-click Cover Label to edit. "
            "Click Include to toggle. "
            f"{naming}."
        )

    def _refresh_tree(self) -> None:
        self._close_editor()
        self.tree.delete(*self.tree.get_children())
        self._iid_to_index.clear()
        for index, job in enumerate(self.jobs):
            iid = str(index)
            self._iid_to_index[iid] = index
            self.tree.insert(
                "",
                tk.END,
                iid=iid,
                values=(
                    "Yes" if job.include else "No",
                    job.source.name,
                    job.label,
                    str(job.source.parent),
                ),
            )
        included = sum(1 for j in self.jobs if j.include)
        self.status_var.set(
            f"{len(self.jobs)} file(s) loaded · {included} included"
            if self.jobs
            else "Add PDFs or open a folder to begin."
        )

    def _selected_indices(self) -> list[int]:
        indices: list[int] = []
        for iid in self.tree.selection():
            if iid in self._iid_to_index:
                indices.append(self._iid_to_index[iid])
        return sorted(indices)

    def _merge_jobs(self, new_jobs: list[JobItem], *, replace: bool) -> None:
        if replace:
            self.jobs = list(new_jobs)
        else:
            existing = {job.source for job in self.jobs}
            for job in new_jobs:
                if job.source not in existing:
                    self.jobs.append(job)
                    existing.add(job.source)
            self.jobs.sort(key=lambda j: j.source.name.casefold())
        self._refresh_tree()

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
            self.status_var.set(
                f"Restored last folder · {len(self.jobs)} file(s) · "
                f"{sum(1 for j in self.jobs if j.include)} included"
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
            # If no last_folder yet, treat this as the session folder.
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
        indices = self._selected_indices()
        if not indices:
            return
        for index in reversed(indices):
            del self.jobs[index]
        self._refresh_tree()

    def _on_reset_labels(self) -> None:
        if self._busy:
            return
        indices = self._selected_indices()
        targets = indices if indices else list(range(len(self.jobs)))
        for index in targets:
            job = self.jobs[index]
            job.label = cover_label_from_filename(job.source)
        self._refresh_tree()

    def _on_toggle_include(self) -> None:
        if self._busy:
            return
        indices = self._selected_indices()
        if not indices:
            return
        for index in indices:
            self.jobs[index].include = not self.jobs[index].include
        self._refresh_tree()
        for index in indices:
            self.tree.selection_add(str(index))

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

    # --- Inline editing --------------------------------------------------

    def _on_tree_click(self, event: tk.Event) -> None:  # type: ignore[type-arg]
        if self._busy:
            return
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.tree.identify_column(event.x)
        iid = self.tree.identify_row(event.y)
        if not iid or iid not in self._iid_to_index:
            return
        if column == "#1":
            index = self._iid_to_index[iid]
            self.jobs[index].include = not self.jobs[index].include
            self.tree.set(iid, "include", "Yes" if self.jobs[index].include else "No")
            included = sum(1 for j in self.jobs if j.include)
            self.status_var.set(
                f"{len(self.jobs)} file(s) loaded · {included} included"
            )

    def _on_tree_double_click(self, event: tk.Event) -> None:  # type: ignore[type-arg]
        if self._busy:
            return
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.tree.identify_column(event.x)
        iid = self.tree.identify_row(event.y)
        if not iid or column != "#3":
            return
        self._start_edit(iid, "label")

    def _start_edit(self, iid: str, column: str) -> None:
        self._close_editor(commit=True)
        bbox = self.tree.bbox(iid, column)
        if not bbox:
            return
        x, y, width, height = bbox
        value = self.tree.set(iid, column)
        entry = ttk.Entry(self.tree)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, value)
        entry.select_range(0, tk.END)
        entry.focus_set()
        entry.bind("<Return>", lambda _e: self._close_editor(commit=True))
        entry.bind("<Escape>", lambda _e: self._close_editor(commit=False))
        entry.bind("<FocusOut>", lambda _e: self._close_editor(commit=True))
        self._edit_entry = entry
        self._edit_iid = iid
        self._edit_column = column

    def _close_editor(self, commit: bool = True) -> None:
        entry = self._edit_entry
        iid = self._edit_iid
        column = self._edit_column
        self._edit_entry = None
        self._edit_iid = None
        self._edit_column = None
        if entry is None or iid is None or column is None:
            return
        new_value = entry.get()
        entry.destroy()
        if not commit or iid not in self._iid_to_index:
            return
        index = self._iid_to_index[iid]
        if column == "label":
            self.jobs[index].label = new_value.strip() or cover_label_from_filename(
                self.jobs[index].source
            )
            self.tree.set(iid, "label", self.jobs[index].label)

    # --- Generate --------------------------------------------------------

    def _on_generate(self) -> None:
        if self._busy:
            return
        self._close_editor(commit=True)
        included = [j for j in self.jobs if j.include]
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
            JobItem(source=j.source, label=j.label, include=j.include)
            for j in self.jobs
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

        # Persist toggles / output before a long run.
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
        self.generate_btn.configure(state=tk.DISABLED if busy else tk.NORMAL)

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
            # Brief pause so the bar hits 100% before closing.
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
