"""
ocr-debug.py — run OCR on screenshot.png using all 3 preprocessing methods.

Usage:
    python ocr-debug.py [image_path]   (default: screenshot.png)
"""

import os
import sys
import re
import cv2
import numpy as np
import pytesseract as tess
from PIL import Image

# ── Tesseract auto-detection ───────────────────────────────────────────────
_script_dir  = os.path.dirname(os.path.abspath(sys.argv[0]))
_local_tess  = os.path.join(_script_dir, 'tesseract', 'tesseract.exe')
_system_tess = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
if os.path.exists(_local_tess):
    tess.pytesseract.tesseract_cmd = _local_tess
    os.environ['TESSDATA_PREFIX'] = os.path.join(_script_dir, 'tesseract', 'tessdata')
else:
    tess.pytesseract.tesseract_cmd = _system_tess

TESS_OPT = '--psm 6 --oem 1 -c tessedit_char_whitelist=0123456789.MR'

_sharpen = np.array([[-1, -1, -1],
                     [-1,  9, -1],
                     [-1, -1, -1]])

SEP = '-' * 48

# ── Preprocessing ─────────────────────────────────────────────────────────
def preprocess(cv_img, method):
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
def parse_ocr(raw_text):
    cleaned = re.sub(r'[^\d\.MR\n]', '', raw_text)
    cleaned = cleaned.replace('.M', 'M').replace('.R', 'R').replace('R.', 'M')
    lines   = [l for l in cleaned.split('\n') if l.strip()]
    def parse_line(line, delim):
        parts = line.split(delim)
        if len(parts) == 3:
            return parts
        nums = re.findall(r'\d+\.\d+', line)
        return nums[:3] if len(nums) >= 3 else None
    mwd = parse_line(lines[0], 'M') if len(lines) > 0 else None
    rss = parse_line(lines[1], 'R') if len(lines) > 1 else None
    return mwd, rss


# ── Main ──────────────────────────────────────────────────────────────────
image_path = sys.argv[1] if len(sys.argv) > 1 else 'screenshot.png'

cv_img = cv2.imread(image_path)
if cv_img is None:
    print(f'ERROR: cannot read {image_path}')
    sys.exit(1)

print(f'Image : {image_path}  ({cv_img.shape[1]}x{cv_img.shape[0]})')
print(SEP)

for method in ('replace', 'threshold', 'original'):
    pil  = preprocess(cv_img, method)
    raw  = tess.image_to_string(pil, config=TESS_OPT)
    mwd, rss = parse_ocr(raw)

    pil.save(f'debug_{method}.png')

    print(f'[{method}]')
    print(f'  raw : {raw.replace(chr(10), " ").strip()!r}')
    print(f'  MWD : {mwd}')
    print(f'  RSS : {rss}')
    print(SEP)

print('Saved: debug_replace.png  debug_threshold.png  debug_original.png')
