"""Desktop launcher for the Google Drive file uploader.

This GUI collects the same configuration values that the CLI script expects,
launches the uploader in a background subprocess, and streams its logs into
the window so non-technical users can run it with minimal setup.
"""

from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


APP_TITLE = "Drive Upload Launcher"
if getattr(sys, "frozen", False):
    SCRIPT_DIR = Path(sys.executable).resolve().parent
else:
    SCRIPT_DIR = Path(__file__).resolve().parent
SETTINGS_FILE = SCRIPT_DIR / ".drive_file_processor_gui.json"
DEFAULT_CREDENTIALS_FILE = SCRIPT_DIR / "credentials.json"
DEFAULT_TOKEN_FILE = SCRIPT_DIR / "token.pickle"
DEFAULT_LOG_FILE = SCRIPT_DIR / "upload.log"
DEFAULT_ERROR_LOG_FILE = SCRIPT_DIR / "upload_errors.log"
DEFAULT_STATE_FILE = SCRIPT_DIR / "upload_state.json"
DEFAULT_CHUNK_SIZE = 10 * 1024 * 1024
SUPPORTED_EXTENSIONS = {".pdf", ".jpg", ".jpeg", ".mp4"}
UI_EVENT_PREFIX = "UI_EVENT "


@dataclass
class FieldSpec:
    key: str
    label: str
    default: str
    browse_kind: str | None = None
    required: bool = False


FIELD_SPECS: list[FieldSpec] = [
    FieldSpec("root_folder", "Root Folder", "", browse_kind="folder", required=True),
    FieldSpec("drive_folder_id", "Drive Folder ID", "", required=True),
    FieldSpec("output_file", "Output Excel File", "uploaded_files.xlsx", browse_kind="savefile", required=True),
]


class DriveUploaderApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.minsize(940, 680)

        self.values: dict[str, tk.StringVar] = {}
        self._process: subprocess.Popen[str] | None = None
        self._reader_thread: threading.Thread | None = None
        self._queue: queue.Queue[tuple[str, str]] = queue.Queue()
        self._stop_requested = False
        self._total_files = 0
        self._completed_files = 0
        self._uploaded_files = 0
        self._duplicate_files = 0
        self._failed_files = 0
        self._last_summary: dict[str, int] = {}

        self._build_ui()
        self._load_settings()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._poll_queue)

    def _build_ui(self) -> None:
        container = ttk.Frame(self, padding=16)
        container.pack(fill="both", expand=True)

        header = ttk.Frame(container)
        header.pack(fill="x")
        ttk.Label(header, text="Drive Upload Launcher", font=("TkDefaultFont", 18, "bold")).pack(anchor="w")
        ttk.Label(
            header,
            text="Choose the folder, Drive target, and output file, then start the upload in one click.",
        ).pack(anchor="w", pady=(4, 12))

        form = ttk.LabelFrame(container, text="Configuration", padding=12)
        form.pack(fill="x")

        for row, spec in enumerate(FIELD_SPECS):
            ttk.Label(form, text=spec.label).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=6)
            var = tk.StringVar(value=spec.default)
            self.values[spec.key] = var

            entry = ttk.Entry(form, textvariable=var, width=72)
            entry.grid(row=row, column=1, sticky="ew", pady=6)

            if spec.browse_kind:
                ttk.Button(
                    form,
                    text="Browse",
                    command=lambda field=spec: self._browse(field),
                ).grid(row=row, column=2, sticky="w", padx=(10, 0), pady=6)

        form.columnconfigure(1, weight=1)

        note = ttk.Label(
            container,
            text=(
                "Tip: the Drive Folder ID is the long string in the Drive folder URL. "
                "The app will create the matching folder tree automatically."
            ),
            wraplength=920,
        )
        note.pack(fill="x", pady=(10, 10))

        progress_frame = ttk.LabelFrame(container, text="Progress", padding=12)
        progress_frame.pack(fill="x")
        self.progress_var = tk.StringVar(value="0 / 0 files processed")
        self.summary_var = tk.StringVar(value="Waiting to start")
        self.status_var = tk.StringVar(value="Ready")
        self.progress_bar = ttk.Progressbar(progress_frame, mode="determinate", maximum=1, value=0)
        self.progress_bar.pack(fill="x")
        ttk.Label(progress_frame, textvariable=self.progress_var).pack(anchor="w", pady=(6, 0))
        ttk.Label(progress_frame, textvariable=self.summary_var).pack(anchor="w", pady=(2, 0))

        controls = ttk.Frame(container)
        controls.pack(fill="x", pady=(10, 0))
        self.start_button = ttk.Button(controls, text="Run Upload", command=self._start_upload)
        self.start_button.pack(side="left")
        self.stop_button = ttk.Button(controls, text="Stop", command=self._stop_upload, state="disabled")
        self.stop_button.pack(side="left", padx=(8, 0))
        ttk.Label(controls, textvariable=self.status_var).pack(side="right")

        log_frame = ttk.LabelFrame(container, text="Live Log", padding=8)
        log_frame.pack(fill="both", expand=True, pady=(12, 0))

        self.log_text = tk.Text(log_frame, wrap="word", height=24, state="disabled")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def _browse(self, spec: FieldSpec) -> None:
        if spec.browse_kind == "folder":
            selected = filedialog.askdirectory(title=f"Select {spec.label}")
        elif spec.browse_kind == "file":
            selected = filedialog.askopenfilename(title=f"Select {spec.label}")
        elif spec.browse_kind == "savefile":
            selected = filedialog.asksaveasfilename(title=f"Select {spec.label}")
        else:
            selected = ""

        if selected:
            self.values[spec.key].set(selected)

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message)
        if not message.endswith("\n"):
            self.log_text.insert("end", "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _count_supported_files(self, root_folder: Path) -> int:
        total = 0
        for current_root, _, files in os.walk(root_folder):
            current_root_path = Path(current_root)
            for file_name in files:
                if (current_root_path / file_name).suffix.lower() in SUPPORTED_EXTENSIONS:
                    total += 1
        return total

    def _refresh_progress(self) -> None:
        maximum = max(1, self._total_files)
        self.progress_bar.configure(maximum=maximum)
        self.progress_bar["value"] = min(self._completed_files, maximum)
        self.progress_var.set(f"{self._completed_files} / {self._total_files} files processed")
        self.summary_var.set(
            f"Uploaded {self._uploaded_files} | Duplicates {self._duplicate_files} | Failed {self._failed_files}"
        )

    def _build_summary_text(self) -> str:
        scanned = self._last_summary.get("scanned", self._total_files)
        uploaded = self._last_summary.get("uploaded", self._uploaded_files)
        duplicates = self._last_summary.get("duplicates", self._duplicate_files)
        failed = self._last_summary.get("failed", self._failed_files)
        return (
            f"Processed {self._completed_files}/{scanned} files. "
            f"Uploaded {uploaded}, duplicates {duplicates}, failed {failed}."
        )

    def _handle_ui_event(self, payload: dict[str, object]) -> None:
        event_type = str(payload.get("type", ""))
        if event_type == "run_started":
            scanned = payload.get("scanned")
            if isinstance(scanned, int):
                self._total_files = scanned
            self.status_var.set("Upload started")
            self._refresh_progress()
            return

        if event_type == "file_complete":
            self._completed_files += 1
            status = str(payload.get("status", ""))
            if status == "uploaded":
                self._uploaded_files += 1
            elif status == "duplicate":
                self._duplicate_files += 1
            elif status in {"failed", "permission_failed"}:
                self._failed_files += 1
            self._refresh_progress()
            return

        if event_type == "run_completed":
            summary = {}
            for key in ("scanned", "uploaded", "duplicates", "failed"):
                value = payload.get(key)
                if isinstance(value, int):
                    summary[key] = value
            self._last_summary = summary
            self.status_var.set("Upload completed")
            self._refresh_progress()
            return

        if event_type == "run_interrupted":
            self.status_var.set("Upload interrupted")
            return

        if event_type == "run_failed":
            error = str(payload.get("error", "Unknown error"))
            self.status_var.set("Upload failed")
            self.summary_var.set(error)
            return

    def _load_settings(self) -> None:
        if not SETTINGS_FILE.exists():
            return

        try:
            payload = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        except Exception:
            return

        for spec in FIELD_SPECS:
            value = payload.get(spec.key)
            if isinstance(value, str):
                self.values[spec.key].set(value)

    def _save_settings(self) -> None:
        payload = {spec.key: self.values[spec.key].get().strip() for spec in FIELD_SPECS}
        SETTINGS_FILE.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def _validate(self) -> dict[str, str] | None:
        values = {spec.key: self.values[spec.key].get().strip() for spec in FIELD_SPECS}
        missing = [spec.label for spec in FIELD_SPECS if spec.required and not values[spec.key]]
        if missing:
            messagebox.showerror("Missing Information", "Please fill in: " + ", ".join(missing))
            return None

        root_folder = Path(values["root_folder"]).expanduser()
        if not root_folder.exists() or not root_folder.is_dir():
            messagebox.showerror("Invalid Root Folder", "Select a valid existing folder for Root Folder.")
            return None

        output_file = Path(values["output_file"]).expanduser()
        if not output_file.suffix:
            values["output_file"] = str(output_file.with_suffix(".xlsx"))
        else:
            values["output_file"] = str(output_file)

        credentials_file = Path(os.environ.get("GOOGLE_CREDENTIALS_FILE", str(DEFAULT_CREDENTIALS_FILE)))
        if not credentials_file.exists():
            messagebox.showerror(
                "Missing Credentials",
                f"Could not find {credentials_file}. Place credentials.json next to the app or exe.",
            )
            return None

        return values

    def _start_upload(self) -> None:
        if self._process is not None:
            messagebox.showinfo("Upload Running", "An upload is already in progress.")
            return

        values = self._validate()
        if values is None:
            return

        self._save_settings()
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        root_folder = Path(values["root_folder"]).expanduser()
        self._total_files = self._count_supported_files(root_folder)
        self._completed_files = 0
        self._uploaded_files = 0
        self._duplicate_files = 0
        self._failed_files = 0
        self._last_summary = {}
        self._refresh_progress()
        self.summary_var.set("Preparing upload...")
        self.status_var.set("Starting upload...")
        self._stop_requested = False
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")

        env = os.environ.copy()
        env.update(
            {
                "ROOT_FOLDER": values["root_folder"],
                "DRIVE_FOLDER_ID": values["drive_folder_id"],
                "OUTPUT_FILE": values["output_file"],
                "GOOGLE_CREDENTIALS_FILE": str(Path(env.get("GOOGLE_CREDENTIALS_FILE") or DEFAULT_CREDENTIALS_FILE)),
                "GOOGLE_TOKEN_FILE": str(Path(env.get("GOOGLE_TOKEN_FILE") or DEFAULT_TOKEN_FILE)),
                "UPLOAD_LOG_FILE": str(Path(env.get("UPLOAD_LOG_FILE") or DEFAULT_LOG_FILE)),
                "UPLOAD_ERROR_LOG_FILE": str(Path(env.get("UPLOAD_ERROR_LOG_FILE") or DEFAULT_ERROR_LOG_FILE)),
                "UPLOAD_STATE_FILE": str(Path(env.get("UPLOAD_STATE_FILE") or DEFAULT_STATE_FILE)),
                "DRIVE_CHUNK_SIZE": str(env.get("DRIVE_CHUNK_SIZE") or DEFAULT_CHUNK_SIZE),
                "UI_PROGRESS": "1",
            }
        )

        backend_exe = SCRIPT_DIR / "drive_file_processor.exe"
        backend_script = SCRIPT_DIR / "drive_file_processor.py"
        if backend_exe.exists():
            command = [str(backend_exe)]
        elif backend_script.exists():
            command = [sys.executable, str(backend_script)]
        else:
            messagebox.showerror(
                "Backend Missing",
                "Could not find drive_file_processor.py or drive_file_processor.exe next to the launcher.",
            )
            self.start_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            self.status_var.set("Backend missing")
            return

        creationflags = 0
        if sys.platform.startswith("win") and hasattr(subprocess, "CREATE_NO_WINDOW"):
            creationflags = subprocess.CREATE_NO_WINDOW  # type: ignore[attr-defined]

        try:
            self._process = subprocess.Popen(
                command,
                cwd=str(SCRIPT_DIR),
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                creationflags=creationflags,
            )
        except Exception as exc:
            self._process = None
            self.start_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            self.status_var.set("Failed to start upload")
            messagebox.showerror("Start Failed", str(exc))
            return

        self._reader_thread = threading.Thread(target=self._read_process_output, daemon=True)
        self._reader_thread.start()

    def _read_process_output(self) -> None:
        process = self._process
        if process is None or process.stdout is None:
            self._queue.put(("status", "No process output available."))
            return

        for line in process.stdout:
            text = line.rstrip("\n")
            if text.startswith(UI_EVENT_PREFIX):
                payload_text = text[len(UI_EVENT_PREFIX) :]
                try:
                    payload = json.loads(payload_text)
                except json.JSONDecodeError:
                    self._queue.put(("log", text))
                else:
                    self._queue.put(("event", json.dumps(payload)))
            else:
                self._queue.put(("log", text))

        return_code = process.wait()
        self._queue.put(("done", str(return_code)))

    def _stop_upload(self) -> None:
        process = self._process
        if process is None:
            return

        if messagebox.askyesno("Stop Upload", "Stop the running upload now?"):
            self._stop_requested = True
            process.terminate()
            self.status_var.set("Stopping upload...")

    def _finish_process(self, return_code: int) -> None:
        self._process = None
        self.start_button.configure(state="normal")
        self.stop_button.configure(state="disabled")

        summary_text = self._build_summary_text()
        if return_code == 0:
            self.status_var.set("Upload completed successfully.")
            self.summary_var.set(summary_text)
            messagebox.showinfo("Upload Complete", summary_text)
        elif self._stop_requested or return_code == 130:
            self.status_var.set("Upload stopped.")
            self.summary_var.set(summary_text)
            messagebox.showwarning("Upload Stopped", summary_text)
        else:
            self.status_var.set(f"Upload finished with errors (code {return_code}).")
            self.summary_var.set(summary_text)
            messagebox.showerror("Upload Failed", f"{summary_text}\nExit code: {return_code}")

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                if kind == "log":
                    self._append_log(payload)
                elif kind == "status":
                    self.status_var.set(payload)
                elif kind == "event":
                    try:
                        event = json.loads(payload)
                    except json.JSONDecodeError:
                        self._append_log(payload)
                    else:
                        self._handle_ui_event(event)
                elif kind == "done":
                    self._append_log(f"\nProcess exited with code {payload}")
                    self._finish_process(int(payload))
        except queue.Empty:
            pass
        finally:
            self.after(100, self._poll_queue)

    def _on_close(self) -> None:
        if self._process is not None:
            if not messagebox.askyesno("Quit", "An upload is running. Stop it and exit?"):
                return
            self._stop_requested = True
            self._process.terminate()
        self.destroy()


def main() -> int:
    app = DriveUploaderApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
