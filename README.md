# ReadScreen

**ReadScreen** is a Windows utility that reads real-time directional drilling survey data — **Depth**, **Inclination (INC)**, and **Azimuth (AZI)** — directly from the *Rig Floor Console* application screen using Tesseract OCR, then outputs the parsed values to `output.csv` for live survey calculations.

It comes in two flavours: a **CLI version** (`main-auto.py`) with a rich terminal UI and a **GUI version** (`main-gui.py`) with a compact dark-themed tabbed interface.

---

## Features

- **Zero configuration** auto-detection of the Rig Floor Console window (no manual coordinate entry needed)
- **Click-to-capture** coordinate picker in GUI — click two corners on screen to set manual crop region
- Supports **MWD** and **RSS** survey rows simultaneously; **Motor** mode reads only the MWD row
- Three OCR preprocessing methods: `replace` (green-on-dark display), `threshold` (Otsu), `original`
- **Motor** and **RSS** tool types with automatic crop ratios
- Configurable scale factor (1–10×) and capture interval (1–60 s)
- Outputs to `output.csv` at the set interval — plug directly into the companion Excel sheet
- Bundled local Tesseract — **no system-wide install required**
- Ships as a folder-based `.exe` (PyInstaller) — compatible with **Windows 10 and 11**

---

## Quick Start (Pre-compiled EXE)

> No Python install required on the target machine.

1. Download and extract the latest release from [Releases](https://github.com/mrkaqz/readscreen/releases)
2. Ensure the folder structure is intact:
   ```
   readscreen-gui-v1.4/
   ├── main-gui.exe        # GUI version
   ├── _internal/          # Python runtime (do not move)
   ├── tess_config.json    # OCR configuration
   └── tesseract/          # Local Tesseract engine
   ```
3. Run `main-gui.exe` for the GUI, or `main-auto.exe` for the terminal version
4. If the app fails to start, install [Visual C++ Redistributable x64](https://aka.ms/vs/17/release/vc_redist.x64.exe)

---

## GUI Overview

The GUI is split into two tabs:

**Setup tab** — configure before starting:

| Control | Description |
|---|---|
| Method | OCR preprocessing: `replace` / `threshold` / `original` |
| Tool | `RSS` (reads MWD + RSS rows) or `Motor` (reads MWD only) |
| Scale | Resize factor before OCR (1–10) |
| Interval (s) | Capture interval in seconds (1–60) |
| Locate | `Auto` detects the Rig Floor Console window; `Manual` uses XY coordinates |
| + Pick | Click two corners on screen to auto-fill X1 Y1 X2 Y2 |
| START / STOP | Start or stop the capture loop |
| Save / Load Config | Save or load all settings to `tess_config.json` |

**Data tab** — shown automatically when running:
- Live **MWD** card (green) and **RSS** card (blue) showing Depth / INC / AZI
- Values flash white on each update
- Switches back to Setup tab automatically on STOP

**Status bar + Log** — always visible at the bottom of both tabs.

---

## Running from Source

### Prerequisites

| Requirement | Version |
|---|---|
| Python | 3.10+ (3.13 tested) |
| Windows | 10 / 11 |

### Installation

```bash
# Clone the repo
git clone https://github.com/mrkaqz/readscreen.git
cd readscreen

# Create and activate virtual environment
python -m venv venv
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### Tesseract Setup

Run once to extract a local Tesseract copy into `./tesseract/`:

```bash
# Place tesseract-ocr-w64-setup-v5.0.0-alpha.20201127.exe in the project folder first
python setup_tesseract.py
```

Alternatively, install [Tesseract system-wide](https://github.com/UB-Mannheim/tesseract/wiki) — the scripts will fall back to `C:\Program Files\Tesseract-OCR\tesseract.exe` automatically.

---

## Usage

```bash
# Activate venv first
venv\Scripts\activate

# GUI version (recommended)
python main-gui.py

# CLI version
python main-auto.py

# Debug OCR — tests all 3 preprocessing methods side-by-side
python ocr-debug.py

# Generate tess_config.json by clicking screen coordinates
python create-config.py

# Find the exact window title of the Rig Floor Console
python winname.py
```

---

## Configuration (`tess_config.json`)

```jsonc
{
  "method": "replace",         // "replace" | "threshold" | "original"
  "tesseract_config": "--psm 6 --oem 1",
  "scale_factor": "3",         // OCR resize multiplier (1–10)
  "interval": "2",             // Capture interval in seconds (1–60)
  "tool": "rss",               // "rss" | "motor"
  "locate": "auto",            // "auto" | "manual"
  "loc_x1": "0",               // Manual crop coords (ignored when locate=auto)
  "loc_y1": "0",
  "loc_x2": "0",
  "loc_y2": "0"
}
```

| Field | Description |
|---|---|
| `method` | OCR preprocessing. Use `replace` for green-on-dark displays, `threshold` for high-contrast, `original` for raw |
| `tesseract_config` | Tesseract CLI flags passed to `pytesseract` |
| `scale_factor` | Integer resize multiplier applied before OCR (higher = better accuracy, slower) |
| `interval` | Seconds between each capture cycle |
| `tool` | `rss` reads both MWD and RSS rows; `motor` reads MWD only |
| `locate` | `auto` finds the Rig Floor Console window by title; `manual` uses `loc_x1/y1/x2/y2` |
| `loc_x1/y1/x2/y2` | Screen pixel coordinates of the survey data area (GUI: use **+ Pick** button) |

---

## How Auto Window Detection Works

In **auto mode**, the app:
1. Waits for the foreground window title to contain `"Rig Floor Console -"` (click on the window before pressing Start)
2. Each loop calls `win32gui.FindWindow()` + `GetWindowRect()` to get the current position and size — so it tracks the window even if you move it

No manual coordinate entry is needed.

---

## Companion Excel Sheet

The companion Excel workbook reads `output.csv` every few seconds and calculates:

- Build Rate & Turn Rate
- Dogleg Severity (DLS)
- Survey propagation

| Version | File |
|---|---|
| Latest (v0.3.8) | `Survey Realtime Track Sheet-Auto V.0.3.8.xlsm` |

> **Macro setup**: Enable macros and click *Start Auto-Update* to begin live refreshing.

---

## Building from Source (PyInstaller)

```bash
# Activate venv
venv\Scripts\activate

# Install PyInstaller
pip install pyinstaller

# Build GUI version (no console window)
pyinstaller --onedir --noconsole --icon=icon.ico --name main-gui --distpath dist-gui main-gui.py

# Build CLI version (with console)
pyinstaller --onedir --icon=icon.ico --name main-auto --distpath dist-exe main-auto.py
```

Copy `tess_config.json` and the `tesseract/` folder into the output folder (`dist-gui/main-gui/` or `dist-exe/main-auto/`) before distributing.

> **Note:** PyInstaller produces a folder-based distribution (`--onedir`), which is compatible with both Windows 10 and 11. The output folder contains the `.exe` and an `_internal/` directory — keep them together.

---

## Project Structure

```
readscreen/
├── main-auto.py              # CLI version (v1.4) — production
├── main-gui.py               # GUI version — compact tabbed tkinter app
├── main.py                   # Legacy CLI (v0.3.1, uses mss + config.json)
├── main-replace.py           # Replace-method only variant
├── main-threshold.py         # Threshold-method only variant
├── ocr-debug.py              # Side-by-side OCR debug (all 3 methods)
├── create-config.py          # Interactive coord capture → config.json
├── winname.py                # Print foreground window title in a loop
├── setup_tesseract.py        # Extract local Tesseract from installer
├── replace-traindata.py      # Copy eng.traineddata to tessdata dir
├── tess_config.json          # Active OCR config
├── config.json               # Legacy config (used by main.py)
├── icon.ico                  # App icon
├── requirements.txt          # Python dependencies
└── Survey Realtime Track Sheet-Auto V.0.3.8.xlsm  # Companion Excel
```

---

## License

Apache License 2.0 — see [LICENSE](LICENSE) for details.

Copyright (c) 2021 Ronnarong Wongmalasit
