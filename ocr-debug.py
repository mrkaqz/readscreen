"""
ocr-debug.py — side-by-side OCR debug for all 3 preprocessing methods.

Usage:
    python ocr-debug.py [screenshot.png]

If no image path is given the script captures from the Rig Floor Console
window automatically (same logic as main-auto.py).  Each iteration shows
the raw OCR result and parsed survey values for all three methods, then
waits for Enter before the next capture.
"""

import os
import sys
import re
import cv2
import numpy as np
import pytesseract as tess
from PIL import Image, ImageGrab
import win32gui

# ── Tesseract auto-detection (mirrors main-auto.py) ───────────────────────
_script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
_local_tess  = os.path.join(_script_dir, 'tesseract', 'tesseract.exe')
_system_tess = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
if os.path.exists(_local_tess):
    tess.pytesseract.tesseract_cmd = _local_tess
    os.environ['TESSDATA_PREFIX'] = os.path.join(_script_dir, 'tesseract', 'tessdata')
    print(f'[Tesseract] local: {_local_tess}')
else:
    tess.pytesseract.tesseract_cmd = _system_tess
    print(f'[Tesseract] system: {_system_tess}')

TESS_OPT = '--psm 6 --oem 1 -c tessedit_char_whitelist=0123456789.MR'

_sharpen = np.array([[-1, -1, -1],
                     [-1,  9, -1],
                     [-1, -1, -1]])

SEP  = '-' * 52
SEP2 = '=' * 52


# ── Preprocessing ─────────────────────────────────────────────────────────
def preprocess(cv_img, method):
    """Apply one preprocessing method; return PIL image ready for OCR."""
    img = cv2.filter2D(cv_img, -1, _sharpen)

    if method == 'replace':
        hsv  = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
        mask = cv2.inRange(hsv,
                           np.array([30,  80,  80]),
                           np.array([100, 255, 255]))
        out  = np.full_like(img, 255)
        out[mask > 0] = 0
        out  = cv2.copyMakeBorder(out, 20, 20, 20, 20,
                                  cv2.BORDER_CONSTANT, value=[255, 255, 255])
        return Image.fromarray(cv2.cvtColor(out, cv2.COLOR_BGR2RGB))

    if method == 'threshold':
        img = cv2.bilateralFilter(img, 5, 75, 75)
        gry = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        thr = cv2.threshold(gry, 0, 255,
                            cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
        thr = cv2.morphologyEx(
            thr, cv2.MORPH_OPEN,
            cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2)))
        thr = cv2.copyMakeBorder(thr, 20, 20, 20, 20,
                                 cv2.BORDER_CONSTANT, value=255)
        return Image.fromarray(thr)

    # original
    img = cv2.copyMakeBorder(img, 20, 20, 20, 20,
                             cv2.BORDER_CONSTANT, value=[0, 0, 0])
    return Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))


# ── Parse ─────────────────────────────────────────────────────────────────
def parse_line(raw, delim):
    parts = raw.split(delim)
    if len(parts) == 3:
        return parts
    nums = re.findall(r'\d+\.\d+', raw)
    if len(nums) >= 3:
        return nums[:3]
    return None


def parse_ocr(raw_text):
    """Return (mwd_list_or_None, rss_list_or_None)."""
    cleaned = re.sub(r'[^\d\.MR\n]', '', raw_text)
    cleaned = cleaned.replace('.M', 'M').replace('.R', 'R').replace('R.', 'M')
    lines   = [l for l in cleaned.split('\n') if l.strip()]
    mwd = parse_line(lines[0], 'M') if len(lines) > 0 else None
    rss = parse_line(lines[1], 'R') if len(lines) > 1 else None
    return mwd, rss


# ── Capture ───────────────────────────────────────────────────────────────
def capture_from_window(tool='rss'):
    """Find Rig Floor Console window and return cropped cv2 image."""
    print('Waiting for Rig Floor Console window (click on it)...')
    win_name = None
    while True:
        name = win32gui.GetWindowText(win32gui.GetForegroundWindow())
        if 'Rig Floor Console -' in name:
            win_name = name
            break
        import time; time.sleep(0.4)

    hwnd = win32gui.FindWindow(None, win_name)
    b    = win32gui.GetWindowRect(hwnd)
    img  = ImageGrab.grab((b[0]+7, b[1], b[2]-7, b[3]-7), all_screens=True)
    w, h = img.size
    if tool == 'rss':
        crop = (w - round(w*0.545), round(h*0.925),
                w - round(w*0.225), round(h*0.995))
    else:
        crop = (w - round(w*0.60),  round(h*0.956),
                w - round(w*0.24),  round(h*0.995))
    img  = img.crop(crop)
    img.save('screenshot.png')
    print(f'Captured from: {win_name}')
    return cv2.imread('screenshot.png')


def load_from_file(path):
    cv_img = cv2.imread(path)
    if cv_img is None:
        print(f'ERROR: cannot read {path}')
        sys.exit(1)
    print(f'Loaded: {path}  ({cv_img.shape[1]}x{cv_img.shape[0]} px)')
    return cv_img


# ── Main ──────────────────────────────────────────────────────────────────
def run_once(cv_img, scale):
    if scale != 1:
        cv_img = cv2.resize(cv_img, None, fx=scale, fy=scale,
                            interpolation=cv2.INTER_LANCZOS4)

    print(f'\n{SEP2}')
    print(f'Image size after resize: {cv_img.shape[1]}x{cv_img.shape[0]}  scale={scale}')
    print(SEP2)

    for method in ('replace', 'threshold', 'original'):
        pil    = preprocess(cv_img, method)
        raw    = tess.image_to_string(pil, config=TESS_OPT)
        mwd, rss = parse_ocr(raw)

        print(f'\n[{method.upper()}]')
        print(SEP)
        # show raw OCR (compact — replace newlines)
        raw_display = raw.replace('\n', ' ↵ ').strip()
        print(f'  Raw   : {raw_display!r}')
        print(f'  MWD   : {mwd}')
        print(f'  RSS   : {rss}')
        print(SEP)

    # Save debug images
    for method in ('replace', 'threshold', 'original'):
        pil = preprocess(cv_img, method)
        pil.save(f'debug_{method}.png')
    print('\nDebug images saved: debug_replace.png  debug_threshold.png  debug_original.png')


def main():
    # ── Args ──
    image_path = sys.argv[1] if len(sys.argv) > 1 else None

    print(SEP2)
    print('  ReadScreen  OCR Debug')
    print(SEP2)

    # Scale
    try:
        scale = int(input('Scale factor (1-10) [default 3]: ').strip() or '3')
        scale = max(1, min(10, scale))
    except ValueError:
        scale = 3
    print(f'Using scale={scale}')

    # Tool (only matters for window capture)
    tool = 'rss'
    if image_path is None:
        t = input('Tool type — rss / motor [default rss]: ').strip().lower()
        tool = 'motor' if t == 'motor' else 'rss'

    while True:
        print()
        if image_path:
            cv_img = load_from_file(image_path)
        else:
            cv_img = capture_from_window(tool)

        run_once(cv_img, scale)

        try:
            again = input('\nPress Enter to capture again, or q+Enter to quit: ').strip().lower()
        except (EOFError, KeyboardInterrupt):
            break
        if again == 'q':
            break

        # Allow switching image path on the fly
        if image_path:
            new_path = input(f'Image path [Enter to reuse {image_path}]: ').strip()
            if new_path:
                image_path = new_path

    print('\nDone.')


if __name__ == '__main__':
    main()
