"""
setup_tesseract.py
------------------
Extracts the bundled Tesseract-OCR installer into a local ./tesseract/ folder
so that main-auto.py can run without a system-wide Tesseract installation.

Run this script ONCE before using main-auto.py for the first time.
After setup, main-auto.py will automatically detect and use the local copy.

Strategy: use 7-Zip to extract the NSIS installer directly — no UAC needed,
no registry writes, and works even when the script path contains spaces.
Falls back to the NSIS silent-install method if 7-Zip is not found.
"""

import os
import subprocess
import shutil
import sys
import tempfile
import pytesseract as tess
from rich.console import Console
from rich import print

console = Console()

SCRIPT_DIR      = os.path.dirname(os.path.abspath(__file__))
INSTALLER       = os.path.join(SCRIPT_DIR, 'tesseract-ocr-w64-setup-v5.0.0-alpha.20201127.exe')
TESS_DIR        = os.path.join(SCRIPT_DIR, 'tesseract')
TESS_EXE        = os.path.join(TESS_DIR, 'tesseract.exe')
TESSDATA_DIR    = os.path.join(TESS_DIR, 'tessdata')
LOCAL_TRAINDATA = os.path.join(SCRIPT_DIR, 'eng.traineddata')

# Known 7-Zip locations on Windows
SEVEN_ZIP_CANDIDATES = [
    r'C:\Program Files\7-Zip\7z.exe',
    r'C:\Program Files (x86)\7-Zip\7z.exe',
]

print('[bold purple4]ReadScreen — Local Tesseract Setup[/bold purple4]\n')

# ── 1. Check if already set up ───────────────────────────────────────────────
if os.path.exists(TESS_EXE):
    print(f'[green]Tesseract already present at:[/green] {TESS_DIR}')
else:
    # ── 2. Verify the installer file exists ──────────────────────────────────
    if not os.path.exists(INSTALLER):
        print(f'[red]Installer not found:[/red] {INSTALLER}')
        print('[yellow]Ensure tesseract-ocr-w64-setup-v5.0.0-alpha.20201127.exe is in the script folder.[/yellow]')
        sys.exit(1)

    print(f'Extracting Tesseract to local folder:')
    print(f'  [cyan]{TESS_DIR}[/cyan]\n')

    os.makedirs(TESS_DIR, exist_ok=True)

    # ── 3a. Try 7-Zip extraction (preferred) ─────────────────────────────────
    # 7-Zip can open NSIS installers as archives. This avoids running the
    # installer executable, so no UAC is required and path spaces don't matter.
    seven_zip = next((p for p in SEVEN_ZIP_CANDIDATES if os.path.exists(p)), None)

    if seven_zip:
        print(f'Using 7-Zip: [cyan]{seven_zip}[/cyan]')

        # Extract into a temp directory first (avoids any path-with-spaces
        # quirks in 7-Zip's -o flag parsing), then move to TESS_DIR.
        tmp_dir = tempfile.mkdtemp(prefix='tess_setup_')
        try:
            result = subprocess.run(
                [seven_zip, 'x', INSTALLER, f'-o{tmp_dir}', '-y'],
                capture_output=True, text=True
            )
            if result.returncode != 0:
                print(f'[red]7-Zip failed (code {result.returncode}):[/red]')
                console.print(result.stderr or result.stdout)
                sys.exit(1)

            # Move extracted files into TESS_DIR, skipping the NSIS plugin folder
            for item in os.listdir(tmp_dir):
                if item == '$PLUGINSDIR':          # NSIS internals, not needed
                    continue
                src  = os.path.join(tmp_dir, item)
                dest = os.path.join(TESS_DIR, item)
                if os.path.isdir(src):
                    if os.path.exists(dest):
                        shutil.rmtree(dest)
                    shutil.copytree(src, dest)
                else:
                    shutil.copy2(src, dest)
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    else:
        # ── 3b. Fall back: NSIS silent install via a short temp path ─────────
        # The NSIS /D= flag breaks when the destination contains spaces.
        # Workaround: install to a short temp path, then move the files.
        print('[yellow]7-Zip not found — using NSIS silent install (UAC prompt may appear).[/yellow]')
        tmp_tess = tempfile.mkdtemp(prefix='tess_', dir=os.environ.get('TEMP', 'C:\\Temp'))
        shutil.rmtree(tmp_tess)          # mkdtemp creates it; NSIS needs to create it itself

        result = subprocess.run(
            [INSTALLER, '/NCRC', '/S', f'/D={tmp_tess}'],
            capture_output=True, text=True
        )
        if result.returncode != 0 or not os.path.exists(os.path.join(tmp_tess, 'tesseract.exe')):
            print('[red]NSIS installer failed. Please install Tesseract manually.[/red]')
            print('[yellow]https://github.com/UB-Mannheim/tesseract/wiki[/yellow]')
            sys.exit(1)

        # Move from temp path (no spaces) to the final TESS_DIR
        if os.path.exists(TESS_DIR):
            shutil.rmtree(TESS_DIR)
        shutil.move(tmp_tess, TESS_DIR)

    if not os.path.exists(TESS_EXE):
        print('[red]Extraction finished but tesseract.exe was not found.[/red]')
        print(f'[yellow]Inspect: {TESS_DIR}[/yellow]')
        sys.exit(1)

    print('[green]Tesseract extracted successfully.[/green]\n')

# ── 4. Replace eng.traineddata with the improved custom model ─────────────────
os.makedirs(TESSDATA_DIR, exist_ok=True)
dest_traindata = os.path.join(TESSDATA_DIR, 'eng.traineddata')

if os.path.exists(LOCAL_TRAINDATA):
    src_size  = os.path.getsize(LOCAL_TRAINDATA)
    dest_size = os.path.getsize(dest_traindata) if os.path.exists(dest_traindata) else 0

    if src_size > dest_size:
        size_mb = src_size // 1024 // 1024
        print(f'Copying improved eng.traineddata ({size_mb} MB) to tessdata/ ...')
        shutil.copy2(LOCAL_TRAINDATA, dest_traindata)
        print('[green]Training data updated.[/green]\n')
    else:
        print('[green]Training data already up to date.[/green]\n')
else:
    print('[yellow]Custom eng.traineddata not found — using the default from the installer.[/yellow]\n')

# ── 5. Verify the local Tesseract binary works ───────────────────────────────
tess.pytesseract.tesseract_cmd = TESS_EXE
# Tesseract 5 (UB-Mannheim build) expects TESSDATA_PREFIX to point to the
# tessdata folder itself, not its parent directory.
os.environ['TESSDATA_PREFIX'] = TESSDATA_DIR
try:
    ver = tess.get_tesseract_version()
    print(f'Tesseract version: [green]{ver}[/green]')
except Exception as e:
    print(f'[red]Could not verify Tesseract binary: {e}[/red]')
    sys.exit(1)

print('\n[bold green]Setup complete![/bold green]')
print('You can now run [cyan]main-auto.py[/cyan] without a system-wide Tesseract installation.')
input('\nPress Enter to exit...')
