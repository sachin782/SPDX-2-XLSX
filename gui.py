"""Main GUI for the SPDX JSON to Excel Converter."""

import os
import pathlib
import platform
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Optional

from excel_writer import build_workbook, save_workbook
from parser import SPDXParseError, load_spdx, parse_spdx
from utils import get_downloads_dir, logger

# ── Colour palettes ────────────────────────────────────────────────────────────
LIGHT = {
    "bg": "#F4F6F8",
    "surface": "#FFFFFF",
    "primary": "#2E4057",
    "primary_hover": "#3D5470",
    "accent": "#048A81",
    "accent_hover": "#036B64",
    "text": "#1A1A2E",
    "subtext": "#6B7280",
    "border": "#D1D5DB",
    "success": "#059669",
    "error": "#DC2626",
    "btn_text": "#FFFFFF",
    "progress_trough": "#E5E7EB",
}
DARK = {
    "bg": "#1A1A2E",
    "surface": "#16213E",
    "primary": "#4A90D9",
    "primary_hover": "#6AAEE8",
    "accent": "#0F9B8E",
    "accent_hover": "#14C0B1",
    "text": "#E2E8F0",
    "subtext": "#94A3B8",
    "border": "#334155",
    "success": "#34D399",
    "error": "#F87171",
    "btn_text": "#FFFFFF",
    "progress_trough": "#334155",
}

LAST_DIR_FILE = pathlib.Path.home() / ".spdx_converter_last_dir"


def _load_last_dir() -> str:
    """Return the last-used directory or home."""
    try:
        return LAST_DIR_FILE.read_text().strip()
    except Exception:
        return str(pathlib.Path.home())


def _save_last_dir(path: str) -> None:
    try:
        LAST_DIR_FILE.write_text(path)
    except Exception:
        pass


def _open_file(path: pathlib.Path) -> None:
    """Open a file with the system default application."""
    system = platform.system()
    try:
        if system == "Windows":
            os.startfile(str(path))
        elif system == "Darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception as exc:
        logger.warning("Could not open file automatically: %s", exc)


class ConverterApp(tk.Tk):
    """Main application window."""

    def __init__(self) -> None:
        super().__init__()
        self.title("SPDX JSON to Excel Converter")
        self.resizable(False, False)
        self.minsize(540, 420)

        self._dark = tk.BooleanVar(value=False)
        self._selected_file: Optional[pathlib.Path] = None
        self._c = LIGHT  # active colour palette

        self._build_ui()
        self._apply_theme()
        self._center()

        # Attempt drag-and-drop (requires tkinterdnd2)
        self._try_dnd()

    # ── UI Construction ────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)

        # ── Title bar row ──────────────────────────────────────────────────────
        top = tk.Frame(self, height=56)
        top.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        top.columnconfigure(0, weight=1)

        self._title_lbl = tk.Label(
            top,
            text="SPDX JSON → Excel Converter",
            font=("Arial", 16, "bold"),
        )
        self._title_lbl.grid(row=0, column=0, padx=24, pady=12, sticky="w")

        self._theme_btn = tk.Button(
            top,
            text="🌙  Dark",
            command=self._toggle_theme,
            relief="flat",
            cursor="hand2",
            font=("Arial", 10),
            bd=0,
            padx=10,
            pady=6,
        )
        self._theme_btn.grid(row=0, column=1, padx=16, pady=12, sticky="e")

        # ── Drop zone / upload card ────────────────────────────────────────────
        self._drop_frame = tk.Frame(self, relief="flat", bd=0, height=160)
        self._drop_frame.grid(row=1, column=0, padx=24, pady=(4, 0), sticky="ew")
        self._drop_frame.columnconfigure(0, weight=1)

        self._drop_icon = tk.Label(self._drop_frame, text="📂", font=("Arial", 32))
        self._drop_icon.grid(row=0, column=0, pady=(20, 4))

        self._drop_hint = tk.Label(
            self._drop_frame,
            text="Drag & drop an SPDX JSON file here, or click below to browse",
            font=("Arial", 10),
            wraplength=460,
        )
        self._drop_hint.grid(row=1, column=0, padx=16, pady=(0, 12))

        self._upload_btn = tk.Button(
            self._drop_frame,
            text="  Upload SPDX JSON  ",
            command=self._browse_file,
            relief="flat",
            cursor="hand2",
            font=("Arial", 11, "bold"),
            padx=16,
            pady=8,
        )
        self._upload_btn.grid(row=2, column=0, pady=(0, 20))

        # ── Status card ────────────────────────────────────────────────────────
        self._status_frame = tk.Frame(self, relief="flat", bd=0)
        self._status_frame.grid(row=2, column=0, padx=24, pady=12, sticky="ew")
        self._status_frame.columnconfigure(1, weight=1)

        self._file_icon = tk.Label(self._status_frame, text="📄", font=("Arial", 18))
        self._file_icon.grid(row=0, column=0, padx=(12, 8), pady=12)

        inner = tk.Frame(self._status_frame)
        inner.grid(row=0, column=1, sticky="ew", pady=12)
        inner.columnconfigure(0, weight=1)

        self._filename_lbl = tk.Label(
            inner, text="No file selected", font=("Arial", 11, "bold"), anchor="w"
        )
        self._filename_lbl.grid(row=0, column=0, sticky="w")

        self._status_lbl = tk.Label(
            inner, text="Select a file to begin", font=("Arial", 9), anchor="w"
        )
        self._status_lbl.grid(row=1, column=0, sticky="w")

        # ── Progress bar ───────────────────────────────────────────────────────
        self._progress = ttk.Progressbar(
            self, orient="horizontal", mode="determinate", length=400
        )
        self._progress.grid(row=3, column=0, padx=24, pady=(0, 4), sticky="ew")
        self._progress["value"] = 0
        self._progress.grid_remove()

        # ── Convert button ─────────────────────────────────────────────────────
        self._convert_btn = tk.Button(
            self,
            text="  Convert to Excel  ",
            command=self._start_conversion,
            relief="flat",
            cursor="hand2",
            font=("Arial", 12, "bold"),
            padx=20,
            pady=10,
            state="disabled",
        )
        self._convert_btn.grid(row=4, column=0, pady=(4, 24))

    # ── Theme ──────────────────────────────────────────────────────────────────

    def _toggle_theme(self) -> None:
        self._dark.set(not self._dark.get())
        self._c = DARK if self._dark.get() else LIGHT
        self._theme_btn.config(text="☀️  Light" if self._dark.get() else "🌙  Dark")
        self._apply_theme()

    def _apply_theme(self) -> None:
        c = self._c
        self.configure(bg=c["bg"])

        # Title bar
        for w in (self.nametowidget(self.children.get(k, "")) for k in list(self.children)):
            pass  # skip — handled per widget below

        self._title_lbl.master.configure(bg=c["primary"])
        self._title_lbl.configure(bg=c["primary"], fg=c["btn_text"])
        self._theme_btn.configure(
            bg=c["primary"],
            fg=c["btn_text"],
            activebackground=c["primary_hover"],
            activeforeground=c["btn_text"],
        )

        # Drop zone
        self._drop_frame.configure(bg=c["surface"])
        self._drop_icon.configure(bg=c["surface"], fg=c["primary"])
        self._drop_hint.configure(bg=c["surface"], fg=c["subtext"])
        self._upload_btn.configure(
            bg=c["primary"],
            fg=c["btn_text"],
            activebackground=c["primary_hover"],
            activeforeground=c["btn_text"],
        )

        # Status card
        self._status_frame.configure(bg=c["surface"])
        self._file_icon.configure(bg=c["surface"], fg=c["accent"])
        self._filename_lbl.configure(bg=c["surface"], fg=c["text"])
        self._status_lbl.configure(bg=c["surface"], fg=c["subtext"])
        self._status_frame.nametowidget(self._status_frame.children.get("!frame", "!frame2", ))
        # inner frame
        for child in self._status_frame.winfo_children():
            if isinstance(child, tk.Frame):
                child.configure(bg=c["surface"])

        # Convert button
        state = str(self._convert_btn["state"])
        if state == "disabled":
            self._convert_btn.configure(
                bg=c["border"], fg=c["subtext"],
                activebackground=c["border"], activeforeground=c["subtext"],
            )
        else:
            self._convert_btn.configure(
                bg=c["accent"],
                fg=c["btn_text"],
                activebackground=c["accent_hover"],
                activeforeground=c["btn_text"],
            )

        # Progress bar style
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "TProgressbar",
            troughcolor=c["progress_trough"],
            background=c["accent"],
            thickness=8,
        )

    def _center(self) -> None:
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    # ── File handling ──────────────────────────────────────────────────────────

    def _browse_file(self) -> None:
        initial = _load_last_dir()
        path_str = filedialog.askopenfilename(
            title="Select SPDX JSON File",
            initialdir=initial,
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if path_str:
            self._load_file(pathlib.Path(path_str))

    def _load_file(self, path: pathlib.Path) -> None:
        _save_last_dir(str(path.parent))
        self._selected_file = path
        self._filename_lbl.config(text=path.name)
        self._set_status("File selected — ready to convert.", "normal")
        self._convert_btn.config(state="normal")
        self._apply_theme()

    def _try_dnd(self) -> None:
        """Enable drag-and-drop if tkinterdnd2 is installed and supported."""
        try:
            from tkinterdnd2 import DND_FILES  # type: ignore

            self._drop_frame.drop_target_register(DND_FILES)
            self._drop_frame.dnd_bind(
                "<<Drop>>",
                lambda e: self._load_file(pathlib.Path(e.data.strip("{}"))),
            )
            self._drop_hint.config(text="Drag & drop an SPDX JSON file here, or click below to browse")
        except (ImportError, tk.TclError) as exc:
            logger.warning("Drag-and-drop unavailable: %s", exc)
            self._drop_hint.config(text="Click below to browse for an SPDX JSON file")

    # ── Conversion ─────────────────────────────────────────────────────────────

    def _set_status(self, msg: str, level: str = "normal") -> None:
        colour_map = {
            "normal": self._c["subtext"],
            "success": self._c["success"],
            "error": self._c["error"],
        }
        self._status_lbl.config(text=msg, fg=colour_map.get(level, self._c["subtext"]))

    def _set_progress(self, value: int) -> None:
        self._progress["value"] = value
        self.update_idletasks()

    def _start_conversion(self) -> None:
        if not self._selected_file:
            return
        self._convert_btn.config(state="disabled")
        self._upload_btn.config(state="disabled")
        self._progress.grid()
        self._set_progress(0)
        threading.Thread(target=self._run_conversion, daemon=True).start()

    def _run_conversion(self) -> None:
        try:
            self._set_status("Loading SPDX file…")
            self._set_progress(15)

            data = load_spdx(self._selected_file)
            self._set_progress(35)

            self._set_status("Parsing SPDX sections…")
            sections = parse_spdx(data)
            if not sections:
                raise SPDXParseError("No recognisable SPDX sections found in the document.")
            self._set_progress(60)

            self._set_status("Building Excel workbook…")
            wb = build_workbook(sections)
            self._set_progress(80)

            downloads = get_downloads_dir()
            out_name = self._selected_file.with_suffix(".xlsx").name
            out_path = downloads / out_name

            self._set_status("Saving Excel file…")
            save_workbook(wb, out_path)
            self._set_progress(100)

            self.after(0, self._on_success, out_path)

        except SPDXParseError as exc:
            self.after(0, self._on_error, str(exc))
        except PermissionError as exc:
            self.after(0, self._on_error, f"Permission denied:\n{exc}")
        except Exception as exc:
            logger.exception("Unexpected error during conversion")
            self.after(0, self._on_error, f"Unexpected error:\n{exc}")

    def _on_success(self, out_path: pathlib.Path) -> None:
        self._set_status(f"Saved to {out_path.name}", "success")
        self._convert_btn.config(state="normal")
        self._upload_btn.config(state="normal")
        self._apply_theme()

        answer = messagebox.askquestion(
            "Conversion Successful",
            f"Conversion completed successfully.\n\nExcel file saved to:\n{out_path}\n\nOpen the file now?",
            icon="info",
        )
        if answer == "yes":
            _open_file(out_path)

        self._progress.grid_remove()

    def _on_error(self, message: str) -> None:
        self._set_status("Conversion failed.", "error")
        self._convert_btn.config(state="normal")
        self._upload_btn.config(state="normal")
        self._apply_theme()
        self._progress.grid_remove()
        messagebox.showerror("Conversion Error", message)
