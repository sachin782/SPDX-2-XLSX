# SPDX JSON to Excel Converter

A desktop application that converts SPDX JSON Software Bill of Materials (SBOM) files into structured Excel spreadsheets.

## Features

- **Simple GUI** — browse or drag-and-drop an SPDX JSON file
- **Multi-sheet output** — separate worksheets for Document Info, Packages, Files, Relationships, and Extracted Licensing Info
- **Professional formatting** — bold headers, frozen rows, auto-sized columns, text wrapping
- **Cross-platform** — Windows, macOS, and Linux
- **Dark mode** — toggle with the button in the top-right corner
- **Progress bar** — visual feedback during conversion
- **Auto-open** — optionally opens the Excel file after saving
- **Last-directory memory** — remembers where you last opened a file
- **Logging** — writes `~/converter.log` for troubleshooting

## Requirements

- Python 3.11 or newer
- Tkinter (included with most Python distributions)

## Installation

```bash
pip install -r requirements.txt
```

> `tkinterdnd2` is optional; it enables drag-and-drop. Remove it from `requirements.txt` if installation fails.

## Running

```bash
cd spdx_excel_converter
python main.py
```

## Building a Standalone Executable

### Windows installer (recommended)

Creates a setup program that installs prerequisites (Microsoft Visual C++ Redistributable, if needed) and then installs the application. Publisher: **Sachin Rawat**.

**Requirements:** Python 3.11+, [Inno Setup 6](https://jrsoftware.org/isdl.php) (for the final `.exe` installer)

```powershell
powershell -ExecutionPolicy Bypass -File installer\build_installer.ps1
```

**Output:**
| Artifact | Location |
|----------|----------|
| Standalone app | `dist\SPDX-Excel-Converter.exe` |
| Windows installer | `dist\installer\SPDX-Excel-Converter-Setup.exe` |

End users run `SPDX-Excel-Converter-Setup.exe` — no Python installation required.

### Manual build (all platforms)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "SPDX-Excel-Converter" main.py
```

The executable will appear in the `dist/` directory.

### macOS (app bundle)

```bash
pyinstaller --onefile --windowed --name "SPDX-Excel-Converter" \
    --osx-bundle-identifier com.sachinrawat.spdxconverter main.py
```

## Usage

1. Launch the application (`python main.py` or the built executable).
2. Click **Upload SPDX JSON** (or drag and drop a `.json` file onto the window).
3. Click **Convert to Excel**.
4. The Excel file is saved automatically to your **Downloads** folder with the same base name as the input file.
5. Click **Yes** in the success dialog to open the file immediately.

## Output Location

| Platform | Path |
|----------|------|
| Windows  | `C:\Users\<username>\Downloads\<filename>.xlsx` |
| macOS    | `~/Downloads/<filename>.xlsx` |
| Linux    | `~/Downloads/<filename>.xlsx` |

## Project Structure

```
spdx_excel_converter/
├── main.py          # Entry point
├── gui.py           # Tkinter UI (dark mode, progress bar, drag-and-drop)
├── parser.py        # SPDX JSON parser
├── excel_writer.py  # openpyxl workbook builder
├── utils.py         # Logging, Downloads path, helpers
├── installer/       # Windows build & installer (publisher: Sachin Rawat)
│   ├── build_installer.ps1
│   ├── spdx_converter.spec
│   ├── setup.iss
│   └── version_info.txt
├── requirements.txt
└── README.md
```

## SPDX Sections Extracted

| Worksheet | Source field |
|-----------|-------------|
| Document Info | Top-level document metadata |
| Packages | `packages[]` |
| Files | `files[]` |
| Relationships | `relationships[]` |
| Extracted Licensing Info | `hasExtractedLicensingInfos[]` |

Sections missing from the input file are silently skipped.

## Troubleshooting

- **"Not an SPDX document"** — ensure the JSON file contains a `spdxVersion` field.
- **Drag-and-drop not working** — install `tkinterdnd2` or use the browse button instead.
- **File won't open automatically** — the file is still saved to Downloads; open it manually.
- Check `~/converter.log` for detailed error information.

## Setup
Setup for this application is in Setup folder.
