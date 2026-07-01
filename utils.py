"""Utility functions for SPDX Excel Converter."""

import logging
import pathlib
import platform
from typing import Any


def setup_logging() -> logging.Logger:
    """Configure and return application logger."""
    log_path = pathlib.Path.home() / "converter.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )
    return logging.getLogger("spdx_converter")


logger = setup_logging()


def get_downloads_dir() -> pathlib.Path:
    """Return the user's Downloads directory in a cross-platform way."""
    system = platform.system()
    if system == "Windows":
        import winreg
        try:
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders",
            ) as key:
                downloads = winreg.QueryValueEx(key, "{374DE290-123F-4565-9164-39C4925E467B}")[0]
                return pathlib.Path(downloads)
        except Exception:
            pass
    downloads = pathlib.Path.home() / "Downloads"
    downloads.mkdir(parents=True, exist_ok=True)
    return downloads


def safe_get(obj: Any, *keys: str, default: Any = "") -> Any:
    """Safely traverse nested dicts/lists using dot-notation keys."""
    for key in keys:
        if isinstance(obj, dict):
            obj = obj.get(key, default)
        else:
            return default
    return obj if obj is not None else default


def flatten_list(value: Any) -> str:
    """Convert a list or other value to a readable string."""
    if isinstance(value, list):
        return "\n".join(str(v) for v in value)
    return str(value) if value is not None else ""


def flatten_checksums(checksums: Any) -> str:
    """Format checksums list into readable multiline string."""
    if not isinstance(checksums, list):
        return ""
    return "\n".join(
        f"{c.get('algorithm', '')}:{c.get('checksumValue', '')}"
        for c in checksums
        if isinstance(c, dict)
    )
