import os
import sys
import mss
import mss.tools
import pytesseract as tess

# Auto-detect Tesseract: prefer a local copy in ./tesseract/ so the script
# works without a system-wide installation.  Run setup_tesseract.py once to
# create the local copy; the system install is used as a fallback.
#
# sys.argv[0] is always the path of the running .exe or .py — works for plain
# scripts, Nuitka-compiled binaries, and PyInstaller bundles alike.
_script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
_local_tess  = os.path.join(_script_dir, 'tesseract', 'tesseract.exe')
_system_tess = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
if os.path.exists(_local_tess):
    tess.pytesseract.tesseract_cmd = _local_tess
    # Tesseract 5 (UB-Mannheim build) expects TESSDATA_PREFIX to point to the
    # tessdata folder itself, not its parent directory.
    os.environ['TESSDATA_PREFIX'] = os.path.join(_script_dir, 'tesseract', 'tessdata')
else:
    tess.pytesseract.tesseract_cmd = _system_tess

from PIL import Image, ImageFilter, ImageGrab
import json
import csv
import time
import re
import cv2
import numpy as np
import win32gui
from ctypes import windll, Structure, c_long, byref
from rich.console import Console
from rich.table import Table
from rich.progress import track
from rich import print


#init rich console
console = Console()

#class for locating mouse position
class POINT(Structure):
    _fields_ = [("x", c_long), ("y", c_long)]

#define mouse position
def queryMousePosition():
    pt = POINT()
    windll.user32.GetCursorPos(byref(pt))
    return { "x": pt.x, "y": pt.y}

def wait_key():
    ''' Wait for a key press on the console and return it. '''
    result = None
    if os.name == 'nt':
        import msvcrt
        result = msvcrt.getch()
    else:
        import termios
        fd = sys.stdin.fileno()

        oldterm = termios.tcgetattr(fd)
        newattr = termios.tcgetattr(fd)
        newattr[3] = newattr[3] & ~termios.ICANON & ~termios.ECHO
        termios.tcsetattr(fd, termios.TCSANOW, newattr)

        try:
            result = sys.stdin.read(1)
        except IOError:
            pass
        finally:
            termios.tcsetattr(fd, termios.TCSAFLUSH, oldterm)

    return result

#query mouse position and define XY position for manually screen capture
def create_config():

    print('Move mouse to point X1,Y1 and press any key')
    wait_key()
    pos = queryMousePosition()
    x1 = pos['x']
    y1 = pos['y']
    print('X1: {}, Y1: {}'.format(x1,y1))
    time.sleep(0.3)

    print('Move mouse to point X2,Y2 and press any key')
    wait_key()
    pos = queryMousePosition()
    x2 = pos['x']
    y2 = pos['y']
    print('X2: {}, Y2: {}'.format(x2,y2))
    time.sleep(0.3)

    #write tess_config file
    f = open("tess_config.json", "w+")
    f.write('{\n')
    f.write('    "method" : "{}",\n'.format(method))
    f.write('    "tesseract_config" : "{}",\n'.format(tess_option))
    f.write('    "loc_x1" : "{}",\n'.format(x1))
    f.write('    "loc_y1" : "{}",\n'.format(y1))
    f.write('    "loc_x2" : "{}",\n'.format(x2))
    f.write('    "loc_y2" : "{}"\n'.format(y2))
    f.write('}')
    f.close()

#check and clean up data
def data_check(data_list):

    # add . if not exist
    for c in range(len(data_list)):
        d = data_list[c]
        if "." not in d:
            data_list[c] = f'{d[:len(d)-2]}.{d[len(d)-2:]}'

    #check data if range, if not return zero
    try:
        if float(data_list[1]) >= 100 or float(data_list[2]) >= 360:
            data_list = ['9.99','9.99','9.99']
    except:
        pass

    return data_list

def parse_survey_line(raw_text, delimiter):
    """Parse [depth, inc, azi] using delimiter; falls back to positional decimal
    extraction if the delimiter was misread or dropped by OCR.

    When the crop is slightly wide and OCR picks up the TVD column, the depth
    part may look like "229.55954.12" (depth concatenated with TVD).  We extract
    only the first number with 1-2 decimal places from each split segment so
    that the extra trailing digits are silently discarded.
    """
    parts = raw_text.split(delimiter)
    if len(parts) == 3:
        cleaned = []
        for p in parts:
            m = re.match(r'(\d+\.\d{1,2})', p.lstrip())
            cleaned.append(m.group(1) if m else p.strip())
        return cleaned
    # Fallback: find all decimal numbers in order
    numbers = re.findall(r'\d+\.\d+', raw_text)
    if len(numbers) >= 3:
        return numbers[:3]
    return ['0.00', '0.00', '0.00']

#declear version
print('[bold purple4]Maxwell Read Screen Utility for Real Time Survey Calculation[/bold purple4]')
print('[bold purple4]Version: 1.4.4 Date: 10-Mar-26[/bold purple4]\n')
print('[blue]This Python script based on Tesseract-OCR open source[/blue]')
print('[blue]Copyright (c) 2021 under Apache License, version 2.0[/blue]')
print('[blue]Develop by Ronnarong Wongmalasit (rwongmalasit@slb.com)[/blue]\n')


print('Initializing environment')
print('Check Tesseract version >>> ',end = '')
time.sleep(1)

#check if Tesseract engine is present and get version of it
try:
    tess_version = tess.get_tesseract_version()
except:
    print('[red]Tesseract-OCR Not Found[/red]')
    print('[yellow]Please install Tesseract OCR engine before continue[/yellow]')
    print('[yellow]https://github.com/tesseract-ocr/tesseract[/yellow]')
else:
    print(f'[green]{tess_version}[/green]')
    #if engine installed, then check the train data file size to determine train data set available
    print('Check Tesseract train data file >>> ',end = '')
    time.sleep(1)
    try:
        # TESSDATA_PREFIX points directly to the tessdata folder (Tesseract 5 behaviour).
        # For the system install fallback, use the standard location.
        if 'TESSDATA_PREFIX' in os.environ:
            _tessdata_path = os.path.join(os.environ['TESSDATA_PREFIX'], 'eng.traineddata')
        else:
            _tessdata_path = r'C:\Program Files\Tesseract-OCR\tessdata\eng.traineddata'
        tessdata_size = os.path.getsize(_tessdata_path)
    except:
        print('[red]train data not found[/red]')
    else:
        if tessdata_size < 6000000:
            print('[yellow]basic[/yellow]')
            print('[yellow]To improve OCR result, copy eng.traineddata to C:\\Program Files\\Tesseract-OCR\\tessdata[/yellow]')
            print('[yellow]or run replace-traindata.exe[/yellow]')
        elif tessdata_size > 15000000:
            print('[green]best[/green]')
        elif tessdata_size > 22000000:
            print('[green]LSTM + Legacy[/green]')

#reading config file
print('Reading config file >>> ',end = '')
time.sleep(1)
try:
    with open('tess_config.json') as file:
        tess_config = json.load(file)
    print('[green]completed[/green]')
except:
    #config file not found or error, create defult config file
    f = open("tess_config.json", "w+")
    f.write('{\n')
    f.write('    "method" : "replace",\n')
    f.write('    "tesseract_config" : "--psm 6 --oem 1",\n')
    f.write('    "loc_x1" : "10",\n')
    f.write('    "loc_y1" : "10",\n')
    f.write('    "loc_x2" : "100",\n')
    f.write('    "loc_y2" : "100"\n')
    f.write('}')
    f.close()
    print('[red]not found, create default config file[/red]')
    with open('tess_config.json') as file:
        tess_config = json.load(file)

#method = original/replace/threshold from config file
method = tess_config['method']

#tess_config option from config file
tess_option = tess_config['tesseract_config']

# Ensure character whitelist is active to reduce OCR misreads.
# Restricting Tesseract to only the characters that can appear in survey data
# (digits, decimal point, M/R tool markers) dramatically cuts misidentification.
if 'tessedit_char_whitelist' not in tess_option:
    tess_option += ' -c tessedit_char_whitelist=0123456789.MR'

#print OCR enginer mode
print('OCR engine modes >>> ',end = '')
time.sleep(0.5)
if '--oem 0' in tess_option:
    print('[green]legacy[/green]')
elif '--oem 1' in tess_option:
    print('[green]Neural nets LSTM[/green]')
elif '--oem 2' in tess_option:
    print('[green]Legacy + LSTM[/green]')
else:
    print('[green]default[/green]')

#print processing method
print('Image processing method >>> ',end = '')
time.sleep(0.5)
print(f'[green]{method}[/green]')

#ask input for resize image
scale_factor = input('\nWhat is image resize scale (1-10)? ')
if int(scale_factor) in range(1,10):
    scale_factor = int(scale_factor)
else:
    print("Use default scale = 1")
    scale_factor = 1

#ask what tool in BHA
tool_run = input('What is tool in BHA?\n1. RSS\n2. Motor\nChoose (1/2) = ')
if tool_run == "2":
    tool_run = "motor"
else:
    tool_run = "rss"

#ask if want to manually detemine XY?
auto_screen_locate = input('Using auto locate XY? (yes/no) ')
auto_screen_locate = auto_screen_locate.lower()
if auto_screen_locate == 'no':

    #run create config XY
    create_config()
    with open('tess_config.json') as file:
        tess_config = json.load(file)
else:

    #find windows name to get screenshot
    print('\n[red]Please Click on Rig Floor Console Window[/red]')
    found_windows_title = False
    while not found_windows_title:
        win_name  = win32gui.GetWindowText (win32gui.GetForegroundWindow())
        #match windows name that contain with Rig Floor Console -
        if "Rig Floor Console -" in win_name:
            found_windows_title = True
            print('Found Rig Floor Console Windows')
            print(f'Windows name = {win_name}')
        time.sleep(1)

#resize console screen size — try three methods in order:
# 1) Windows console API (works in classic conhost.exe)
# 2) mode con  (works in cmd.exe / older terminals)
# 3) XTerm resize escape sequence (works in Windows Terminal / ConPTY)
try:
    class _SMALL_RECT(ctypes.Structure):
        _fields_ = [('Left',  ctypes.c_short), ('Top',    ctypes.c_short),
                    ('Right', ctypes.c_short), ('Bottom', ctypes.c_short)]
    _k32 = ctypes.windll.kernel32
    _h   = _k32.GetStdHandle(-11)          # STD_OUTPUT_HANDLE
    _k32.SetConsoleWindowInfo(_h, True, ctypes.byref(_SMALL_RECT(0, 0, 39, 7)))
    _k32.SetConsoleScreenBufferSize(_h, ctypes.c_int32(40 | (500 << 16)))
except Exception:
    pass
os.system('mode con cols=40 lines=8 2>nul')
try:
    ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004
    _k32 = ctypes.windll.kernel32
    _h   = _k32.GetStdHandle(-11)
    _m   = ctypes.c_ulong(0)
    _k32.GetConsoleMode(_h, ctypes.byref(_m))
    _k32.SetConsoleMode(_h, _m.value | ENABLE_VIRTUAL_TERMINAL_PROCESSING)
    sys.stdout.write('\033[8;8;40t')
    sys.stdout.flush()
except Exception:
    pass

# last-known-good survey readings — written to CSV when the current cycle fails,
# so the Excel spreadsheet always receives the most recent valid survey.
last_mwd_list = ['0.00', '0.00', '0.00']
last_rss_list = ['0.00', '0.00', '0.00']

# sharpening kernel applied after upscale to restore edge definition
sharpen_kernel = np.array([[-1, -1, -1],
                            [-1,  9, -1],
                            [-1, -1, -1]])

try:
    #loop the program
    while True:

        #if manual XY, use the XY from config file
        if auto_screen_locate == 'no':

            loc_x1 = int(tess_config['loc_x1'])
            loc_x2 = int(tess_config['loc_x2'])
            loc_y1 = int(tess_config['loc_y1'])
            loc_y2 = int(tess_config['loc_y2'])

            bbox_adj = (loc_x1,loc_y1,loc_x2,loc_y2)

            #grab screenshot
            try:
                img = ImageGrab.grab(bbox_adj,all_screens=True)
                img.save('screenshot.png')
            except:
                print('Error grab screenshot and save')
            else:
                pass

        #else use XY from window name
        else:

            try:
                hwnd = win32gui.FindWindow(None, win_name)
                bbox = win32gui.GetWindowRect(hwnd)
            except:
                print('Missing Rig Floor Console')
            else:
                pass

            #adjust bbox due to windows default invisible border (7,0,7,7)
            bbox_adj = (bbox[0]+7,bbox[1],bbox[2]-7,bbox[3]-7)

            #grab screenshot
            img = ImageGrab.grab(bbox_adj,all_screens=True)

            #find size of screenshot
            width, height = img.size

            #crop screenshot based on tool selected
            if tool_run == "rss":
                # MD/INC/AZI cols only: x=45.5-78%, y=92.0-99.9%
                left = width - round(width*0.545)
                top = round(height*0.920)
                right = width - round(width*0.22)
                bottom = round(height*0.999)
            elif tool_run == "motor":
                # MD/INC/AZI cols only: x=40-76%, y=94.9-99.9%
                left = width - round(width*0.60)
                top = round(height*0.949)
                right = width - round(width*0.24)
                bottom = round(height*0.999)

            #crop image
            img_crop = img.crop((left,top,right,bottom))
            img_crop.save('screenshot.png')


        #read text out of image
        text = ''   # ensure text is always defined even if all retries fail
        for i in range(10):
            try:

                # --- Resize with LANCZOS4 (sharper for text than INTER_CUBIC) ---
                img = cv2.imread('screenshot.png')
                if img is None:
                    raise FileNotFoundError('cv2 could not read screenshot.png')
                img = cv2.resize(img, None, fx=scale_factor, fy=scale_factor,
                                 interpolation=cv2.INTER_LANCZOS4)
                # Sharpen after upscale to restore edge crispness lost during interpolation
                img = cv2.filter2D(img, -1, sharpen_kernel)
                cv2.imwrite('screenshot_resize.png', img)

                # --- replace method: HSV colour-space mask (robust to brightness changes) ---
                # Auto-detects display type via average V channel:
                #   avg_v > 150 → Motor/TeleScope (bright green bg, black text)
                #   avg_v ≤ 150 → RSS (dark bg, bright green text)
                if method == 'replace':

                    img_bgr = cv2.imread('screenshot_resize.png')
                    if img_bgr is None:
                        raise FileNotFoundError('cv2 could not read screenshot_resize.png')
                    img_hsv = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV)
                    avg_v = float(np.mean(img_hsv[:, :, 2]))
                    lower_green = np.array([30,  80,  80])
                    upper_green = np.array([100, 255, 255])
                    mask = cv2.inRange(img_hsv, lower_green, upper_green)
                    # Start with white canvas; green pixels → black
                    result = np.full_like(img_bgr, 255)
                    result[mask > 0] = 0
                    if avg_v > 150:
                        # Motor: bright green bg → was set to black, black text → was white.
                        # Invert so bg becomes white and text becomes black (Tesseract-preferred).
                        result = cv2.bitwise_not(result)
                    # Add white border — Tesseract reads better with surrounding whitespace
                    result = cv2.copyMakeBorder(result, 20, 20, 20, 20,
                                                cv2.BORDER_CONSTANT, value=[255, 255, 255])
                    cv2.imwrite('screenshot_replace.png', result)
                    # Convert to PIL Image for pytesseract compatibility
                    pil_img = Image.fromarray(cv2.cvtColor(result, cv2.COLOR_BGR2RGB))
                    text = tess.image_to_string(pil_img, config=tess_option)

                # --- threshold method: denoise → Otsu → morphological cleanup ---
                # Auto-detects threshold direction via average V channel:
                #   avg_v > 150 → Motor (bright bg): BINARY → black text on white bg
                #   avg_v ≤ 150 → RSS (dark bg): BINARY_INV → black text on white bg
                elif method == 'threshold':

                    img_bgr = cv2.imread('screenshot_resize.png')
                    if img_bgr is None:
                        raise FileNotFoundError('cv2 could not read screenshot_resize.png')
                    # Bilateral filter denoises while preserving character edges
                    img_bgr = cv2.bilateralFilter(img_bgr, 5, 75, 75)
                    gry = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
                    img_hsv_t = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV)
                    avg_v_t = float(np.mean(img_hsv_t[:, :, 2]))
                    thresh_type = (cv2.THRESH_BINARY if avg_v_t > 150
                                   else cv2.THRESH_BINARY_INV) + cv2.THRESH_OTSU
                    thr = cv2.threshold(gry, 0, 255, thresh_type)[1]
                    # Morphological opening removes isolated noise pixels without
                    # eroding the character strokes
                    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
                    thr = cv2.morphologyEx(thr, cv2.MORPH_OPEN, kernel)
                    # Add white border padding
                    thr = cv2.copyMakeBorder(thr, 20, 20, 20, 20,
                                             cv2.BORDER_CONSTANT, value=255)
                    cv2.imwrite('screenshot_resize_thr.png', thr)
                    # Convert to PIL Image for pytesseract compatibility
                    pil_img = Image.fromarray(thr)
                    text = tess.image_to_string(pil_img, config=tess_option)

                else:
                    # original: no colour transformation, just add border padding
                    img_bgr = cv2.imread('screenshot_resize.png')
                    if img_bgr is None:
                        raise FileNotFoundError('cv2 could not read screenshot_resize.png')
                    img_bgr = cv2.copyMakeBorder(img_bgr, 20, 20, 20, 20,
                                                 cv2.BORDER_CONSTANT, value=[0, 0, 0])
                    # Convert to PIL Image for pytesseract compatibility
                    pil_img = Image.fromarray(cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGB))
                    text = tess.image_to_string(pil_img, config=tess_option)

            except Exception as e:
                #if error, print reason and try again
                print(f'Read Image Error ({e}) -- Retrying...')
                continue
            else:
                break

        # --- clean string ---
        text = re.sub(r'[^\d\.MR\n]', '', text)
        text = text.replace('.M', 'M')
        text = text.replace('.R', 'R')
        text = text.replace('R.', 'M')
        # Remove blank lines left after stripping
        list_text = [l for l in text.split('\n') if l.strip()]

        # --- parse using delimiter with positional fallback ---
        mwd_list = parse_survey_line(list_text[0], 'M') if len(list_text) > 0 else ['0.00', '0.00', '0.00']
        # Only parse RSS row when tool is RSS; motor has only one data row
        if tool_run == 'rss':
            rss_list = parse_survey_line(list_text[1], 'R') if len(list_text) > 1 else ['0.00', '0.00', '0.00']
        else:
            rss_list = ['0.00', '0.00', '0.00']

        #perform data check
        mwd_list = data_check(mwd_list)
        if tool_run == 'rss':
            rss_list = data_check(rss_list)

        #if data is not make sense, them give error into list
        if mwd_list[0] == '9.99':
            mwd_out = ['Out','Of','Range']
            mwd_list = last_mwd_list.copy()
        elif mwd_list[0] == '':
            mwd_out = ['No','Data','Found']
            mwd_list = last_mwd_list.copy()
        elif mwd_list[0] == '0.00':
            mwd_out = ['Reading','Error','!!']
            mwd_list = last_mwd_list.copy()
        else:
            try:
                float(mwd_list[0])
                float(mwd_list[1])
                float(mwd_list[2])
                mwd_out = mwd_list
                last_mwd_list = mwd_list.copy()   # save as last known good
            except:
                mwd_out = ['Not','a','Number']
                mwd_list = last_mwd_list.copy()

        if tool_run == 'rss':
            if rss_list[0] == '9.99':
                rss_out = ['Out','Of','Range']
                rss_list = last_rss_list.copy()
            elif rss_list[0] == '':
                rss_out = ['No','Data','Found']
                rss_list = last_rss_list.copy()
            elif rss_list[0] == '0.00':
                rss_out = ['Reading','Error','!!']
                rss_list = last_rss_list.copy()
            else:
                try:
                    float(rss_list[0])
                    float(rss_list[1])
                    float(rss_list[2])
                    rss_out = rss_list
                    last_rss_list = rss_list.copy()   # save as last known good
                except:
                    rss_out = ['Not','a','Number']
                    rss_list = last_rss_list.copy()
        else:
            rss_out = ['N/A', 'N/A', 'N/A']


        #create rich table
        table = Table(show_header=True, header_style="bold green")
        table.add_column("TOOL")
        table.add_column("DEPTH", justify = 'right')
        table.add_column("INC", justify = 'right')
        table.add_column("AZI", justify = 'right')

        #clear
        console.clear()

        #add row data to rich table
        table.add_row("MWD",mwd_out[0],mwd_out[1],mwd_out[2])
        if tool_run == 'rss':
            table.add_row("RSS",rss_out[0],rss_out[1],rss_out[2])
        console.print(table)

        #output  csv file
        for i in range(10):
            try:
                with open('output.csv', mode='w', newline='') as output_file:
                    output_writer = csv.writer(output_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                    output_writer.writerow(mwd_list)
                    if tool_run == 'rss':
                        output_writer.writerow(rss_list)
            except:
                print('Output to file error -- Retrying...')
                continue
            else:
                break

        #print progress bar
        for n in track(range(20), description="Processing"):
            time.sleep(0.1)

except KeyboardInterrupt:
    pass
