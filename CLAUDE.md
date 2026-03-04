# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**ReadScreen** is a Windows-only Python utility that reads directional drilling survey data (Depth, Inclination, Azimuth) from a "Rig Floor Console" application screen using Tesseract OCR, then continuously outputs the parsed data to `output.csv` for real-time survey calculations.

## Running the Scripts

```bash
# Activate virtual environment first
source venv/Scripts/activate  # or venv\Scripts\activate on Windows cmd

# Run the main (latest) production script
python main-auto.py

# Run earlier version (uses config.json instead of tess_config.json)
python main.py

# Debug OCR output - tests all three preprocessing methods side-by-side
python ocr-debug.py

# Generate config file by clicking mouse coordinates
python create-config.py

# Identify foreground window name (useful for finding the right win_name)
python winname.py
```

## Building Executables (Nuitka)

```bash
python -m nuitka --onefile --windows-icon-from-ico=icon.ico --zig --follow-imports main-auto.py
```

Output is a single `main-auto.exe` in the current directory. Place `tesseract/` folder and `tess_config.json` alongside the `.exe` before distributing.

## Prerequisites

- Run `setup_tesseract.py` once to extract a local Tesseract copy into `./tesseract/` (no system-wide install needed). Requires `tesseract-ocr-w64-setup-v5.0.0-alpha.20201127.exe` in the project folder and 7-Zip installed.
- Falls back to `C:\Program Files\Tesseract-OCR\tesseract.exe` if no local copy exists.

## Architecture

### Main Data Flow (`main-auto.py`)
1. **Locate window**: Either auto-find "Rig Floor Console -" window title via `win32gui`, or use manual XY coords from `tess_config.json`
2. **Capture screenshot**: `ImageGrab.grab()` with `all_screens=True`
3. **Crop**: Proportional crop based on tool type (RSS vs Motor) using fixed percentage ratios of window dimensions
4. **Resize**: `cv2.resize()` using `scale_factor` (user-provided, 1-10)
5. **Preprocess** (3 methods controlled by `tess_config.json["method"]`):
   - `replace` — pixel-level: replaces green channel values 100-150 → white, all else → black (designed for green-on-dark displays)
   - `threshold` — Otsu binary inverse thresholding via OpenCV
   - `original` — no preprocessing, raw resize only
6. **OCR**: `pytesseract.image_to_string()` with config `--psm 6 --oem 1`
7. **Parse**: Strip non-`[\d.MR\n]` chars; split on `M` delimiter for MWD row, `R` for RSS row → `[depth, inc, azi]`
8. **Validate** (`data_check()`): Insert decimal if missing; flag out-of-range if INC ≥ 100 or AZI ≥ 360
9. **Output**: Display rich table to console + write `output.csv` (2 rows: MWD, RSS)
10. **Loop**: 2-second progress bar delay between captures

### Configuration Files
- **`tess_config.json`** — Used by `main-auto.py`: `method`, `tesseract_config`, `loc_x1/y1/x2/y2`
- **`config.json`** — Used by older `main.py`: same coordinates plus `monitor` (monitor index for `mss`) and `scale_factor`

### Script Variants
| Script | Purpose |
|--------|---------|
| `main-auto.py` | Current production version (v1.3.2); rich UI, auto window detection |
| `main.py` | Older version (v0.3.1); uses `mss` for capture, simpler text output |
| `main-gui.py` | GUI version |
| `main-replace.py` | Replace-method only variant |
| `main-threshold.py` | Threshold-method only variant |
| `ocr-debug.py` | Runs all three preprocessing methods and prints results for comparison |
| `create-config.py` | Interactive mouse-click coord capture → writes `config.json` |
| `winname.py` | Prints foreground window name in a loop (for finding correct window title) |
| `replace-traindata.py` | Copies `eng.traineddata` to Tesseract tessdata dir (requires admin via `elevate`) |

## Key Constraints

- **Windows-only**: Uses `win32gui`, `ctypes.windll`, `ImageGrab(all_screens=True)`, `msvcrt`
- **Tesseract auto-detection**: `main-auto.py` prefers a local `./tesseract/tesseract.exe` (set up via `setup_tesseract.py`) and falls back to `C:\Program Files\Tesseract-OCR\tesseract.exe`. Path resolved via `sys.argv[0]` so it works in both script and Nuitka-compiled exe.
- The parsed text relies on `M` and `R` suffix characters in the OCR output as delimiters; the display format of the source application is assumed fixed
- `output.csv` is overwritten each loop iteration (not appended)
