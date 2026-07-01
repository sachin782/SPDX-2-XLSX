"""Entry point for SPDX JSON to Excel Converter."""

import sys

# Ensure local modules are importable when running from the project directory
import pathlib
sys.path.insert(0, str(pathlib.Path(__file__).parent))

from gui import ConverterApp


def main() -> None:
    """Launch the converter application."""
    app = ConverterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
