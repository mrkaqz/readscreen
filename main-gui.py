import os
import sys
import re
import cv2
import json
import csv
import time
import queue
import threading
import numpy as np
import tkinter as tk
from tkinter import scrolledtext, messagebox
from PIL import Image, ImageGrab
import pytesseract as tess
import win32gui

# ── Tesseract auto-detection (mirrors main-auto.py) ───────────────────────────
_script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
_local_tess  = os.path.join(_script_dir, 'tesseract', 'tesseract.exe')
_system_tess = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
if os.path.exists(_local_tess):
    tess.pytesseract.tesseract_cmd = _local_tess
    os.environ['TESSDATA_PREFIX'] = os.path.join(_script_dir, 'tesseract', 'tessdata')
else:
    tess.pytesseract.tesseract_cmd = _system_tess

# ── Constants ─────────────────────────────────────────────────────────────────
VERSION     = '1.4.2 GUI'
DATE        = '10-Mar-26'
CONFIG_FILE = 'tess_config.json'
INTERVAL    = 2

_sharpen_kernel = np.array([[-1, -1, -1],
                             [-1,  9, -1],
                             [-1, -1, -1]])

# ── Pure helpers (identical logic to main-auto.py) ────────────────────────────
def data_check(data_list):
    for c in range(len(data_list)):
        d = data_list[c]
        if '.' not in d and d != '':
            data_list[c] = f'{d[:len(d)-2]}.{d[len(d)-2:]}'
    try:
        if float(data_list[1]) >= 100 or float(data_list[2]) >= 360:
            data_list = ['9.99', '9.99', '9.99']
    except Exception:
        pass
    return data_list


def parse_survey_line(raw_text, delimiter):
    """Parse [depth, inc, azi] using delimiter; falls back to positional decimal
    extraction.  Extracts only the first 1-2 decimal place number from each
    split segment to guard against TVD column concatenation."""
    parts = raw_text.split(delimiter)
    if len(parts) == 3:
        cleaned = []
        for p in parts:
            m = re.match(r'(\d+\.\d{1,2})', p.lstrip())
            cleaned.append(m.group(1) if m else p.strip())
        return cleaned
    numbers = re.findall(r'\d+\.\d+', raw_text)
    if len(numbers) >= 3:
        return numbers[:3]
    return ['0.00', '0.00', '0.00']


def fmt_val(v):
    try:
        return f'{float(v):.2f}'
    except Exception:
        return str(v)


# ── Application ───────────────────────────────────────────────────────────────
class App(tk.Tk):

    # Catppuccin Mocha palette
    P = {
        'crust':    '#11111b',
        'mantle':   '#181825',
        'base':     '#1e1e2e',
        'surface0': '#313244',
        'surface1': '#45475a',
        'overlay0': '#6c7086',
        'subtext0': '#a6adc8',
        'text':     '#cdd6f4',
        'lavender': '#b4befe',
        'blue':     '#89b4fa',
        'green':    '#a6e3a1',
        'red':      '#f38ba8',
        'yellow':   '#f9e2af',
        'peach':    '#fab387',
        'teal':     '#94e2d5',
        'mauve':    '#cba6f7',
    }

    def __init__(self):
        super().__init__()
        self.title(f'Read Screen  v{VERSION}')
        self.configure(bg=self.P['base'])
        self.geometry('420x420')
        self.minsize(360, 320)
        self.resizable(True, True)

        self._stop_event = threading.Event()
        self._worker     = None
        self._queue      = queue.Queue()

        self.var_method   = tk.StringVar(value='replace')
        self.var_tool     = tk.StringVar(value='rss')
        self.var_locate   = tk.StringVar(value='auto')
        self.var_scale    = tk.StringVar(value='3')
        self.var_interval = tk.StringVar(value='2')
        self.var_x1     = tk.StringVar(value='10')
        self.var_y1     = tk.StringVar(value='10')
        self.var_x2     = tk.StringVar(value='100')
        self.var_y2     = tk.StringVar(value='100')

        self._build_ui()
        self._load_config()
        self._poll_queue()
        self.protocol('WM_DELETE_WINDOW', self._on_close)

    # ── Widget factories ──────────────────────────────────────────────────────
    @staticmethod
    def _lighten(hex_color, amount=25):
        r = min(255, int(hex_color[1:3], 16) + amount)
        g = min(255, int(hex_color[3:5], 16) + amount)
        b = min(255, int(hex_color[5:7], 16) + amount)
        return f'#{r:02x}{g:02x}{b:02x}'

    def _lbl(self, parent, text, fg=None, font=None, **kw):
        return tk.Label(parent, text=text,
                        bg=parent['bg'], fg=fg or self.P['text'],
                        font=font or ('Segoe UI', 9), **kw)

    def _radio(self, parent, text, var, value, command=None, font=None):
        return tk.Radiobutton(parent, text=text, variable=var, value=value,
                              bg=parent['bg'], fg=self.P['text'],
                              selectcolor=self.P['surface0'],
                              activebackground=parent['bg'],
                              activeforeground=self.P['text'],
                              font=font or ('Segoe UI', 9), command=command)

    def _entry(self, parent, var, width=6):
        return tk.Entry(parent, textvariable=var, width=width,
                        bg=self.P['surface0'], fg=self.P['text'],
                        insertbackground=self.P['text'],
                        relief='flat', bd=5,
                        highlightthickness=1,
                        highlightcolor=self.P['lavender'],
                        highlightbackground=self.P['surface1'],
                        font=('Consolas', 9))

    def _flat_btn(self, parent, text, bg, fg, command, font=None, pady=None, **kw):
        return tk.Button(parent, text=text, bg=bg, fg=fg,
                         activebackground=self._lighten(bg),
                         activeforeground=fg,
                         font=font or ('Segoe UI', 10, 'bold'),
                         relief='flat', bd=0, padx=16,
                         pady=pady if pady is not None else 9,
                         cursor='hand2', command=command, **kw)

    def _section_label(self, text):
        tk.Label(self, text=text, bg=self.P['base'],
                 fg=self.P['overlay0'], font=('Segoe UI', 7, 'bold'),
                 padx=16, pady=0, anchor='w').pack(fill='x', pady=(10, 2))

    def _hsep(self, color=None):
        tk.Frame(self, bg=color or self.P['surface0'], height=1).pack(fill='x')

    # ── Build UI ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        self._build_header()
        self._build_tab_bar()
        self._build_setup_tab()
        self._build_data_tab()
        self._build_statusbar()
        self._build_log()
        self._show_tab('Setup')

    # ── Tab switching ──────────────────────────────────────────────────────────
    def _show_tab(self, name):
        for n, frame in self._tab_frames.items():
            if n == name:
                frame.pack(fill='x')
                self._tab_btns[n].configure(
                    bg=self.P['lavender'], fg=self.P['crust'])
            else:
                frame.pack_forget()
                self._tab_btns[n].configure(
                    bg=self.P['surface0'], fg=self.P['subtext0'])

    # ── Header ────────────────────────────────────────────────────────────────
    def _build_header(self):
        hdr = tk.Frame(self, bg=self.P['crust'], height=30)
        hdr.pack(fill='x')
        hdr.pack_propagate(False)

        left = tk.Frame(hdr, bg=self.P['crust'])
        left.pack(side='left', fill='y', padx=(10, 0))

        dot = tk.Canvas(left, width=7, height=7,
                        bg=self.P['crust'], highlightthickness=0)
        dot.pack(side='left', pady=11)
        dot.create_oval(0, 0, 7, 7, fill=self.P['green'], outline='')

        tk.Label(left, text='Maxwell Read Screen',
                 bg=self.P['crust'], fg=self.P['text'],
                 font=('Segoe UI', 9, 'bold')).pack(side='left', padx=(6, 0))

        tk.Label(hdr, text=f'v{VERSION}',
                 bg=self.P['crust'], fg=self.P['overlay0'],
                 font=('Segoe UI', 8)).pack(side='right', padx=10)

    # ── Tab bar ───────────────────────────────────────────────────────────────
    def _build_tab_bar(self):
        bar = tk.Frame(self, bg=self.P['surface0'], height=26)
        bar.pack(fill='x')
        bar.pack_propagate(False)

        self._tab_frames = {}
        self._tab_btns   = {}

        for name in ('Setup', 'Data'):
            btn = tk.Button(bar, text=name,
                            bg=self.P['surface0'], fg=self.P['subtext0'],
                            activebackground=self.P['lavender'],
                            activeforeground=self.P['crust'],
                            font=('Segoe UI', 8, 'bold'),
                            relief='flat', bd=0, padx=14, pady=4,
                            cursor='hand2',
                            command=lambda n=name: self._show_tab(n))
            btn.pack(side='left')
            self._tab_btns[name] = btn

    # ── Setup tab ─────────────────────────────────────────────────────────────
    def _build_setup_tab(self):
        F = tk.Frame(self, bg=self.P['mantle'])
        self._tab_frames['Setup'] = F

        inner = tk.Frame(F, bg=self.P['mantle'])
        inner.pack(fill='x', padx=10, pady=6)

        # Row 0: Method + Tool
        r0 = tk.Frame(inner, bg=self.P['mantle'])
        r0.pack(fill='x', pady=(0, 3))

        self._lbl(r0, 'Method', fg=self.P['subtext0'],
                  font=('Segoe UI', 8)).pack(side='left')
        tk.Frame(r0, width=4, bg=self.P['mantle']).pack(side='left')
        om = tk.OptionMenu(r0, self.var_method,
                           'replace', 'replace', 'threshold', 'original')
        om.configure(bg=self.P['surface0'], fg=self.P['text'],
                     activebackground=self.P['surface1'],
                     activeforeground=self.P['text'],
                     highlightthickness=0, relief='flat',
                     font=('Segoe UI', 8), bd=0, padx=5, pady=1)
        om['menu'].configure(bg=self.P['surface0'], fg=self.P['text'],
                             activebackground=self.P['lavender'],
                             activeforeground=self.P['crust'],
                             font=('Segoe UI', 8), bd=0)
        om.pack(side='left', padx=(0, 16))

        self._lbl(r0, 'Tool', fg=self.P['subtext0'],
                  font=('Segoe UI', 8)).pack(side='left', padx=(0, 4))
        self._radio(r0, 'RSS',   self.var_tool, 'rss',
                    font=('Segoe UI', 8)).pack(side='left')
        self._radio(r0, 'Motor', self.var_tool, 'motor',
                    font=('Segoe UI', 8)).pack(side='left', padx=(4, 0))

        # Row 1: Scale + Locate
        r1 = tk.Frame(inner, bg=self.P['mantle'])
        r1.pack(fill='x', pady=(0, 3))

        self._lbl(r1, 'Scale', fg=self.P['subtext0'],
                  font=('Segoe UI', 8)).pack(side='left')
        tk.Frame(r1, width=4, bg=self.P['mantle']).pack(side='left')
        self._entry(r1, self.var_scale, width=3).pack(side='left', padx=(0, 16))

        self._lbl(r1, 'Interval (s)', fg=self.P['subtext0'],
                  font=('Segoe UI', 8)).pack(side='left')
        tk.Frame(r1, width=4, bg=self.P['mantle']).pack(side='left')
        self._entry(r1, self.var_interval, width=3).pack(side='left', padx=(0, 16))

        self._lbl(r1, 'Locate', fg=self.P['subtext0'],
                  font=('Segoe UI', 8)).pack(side='left', padx=(0, 4))
        self._radio(r1, 'Auto',   self.var_locate, 'auto',
                    command=self._toggle_xy, font=('Segoe UI', 8)).pack(side='left')
        self._radio(r1, 'Manual', self.var_locate, 'manual',
                    command=self._toggle_xy, font=('Segoe UI', 8)).pack(side='left', padx=(4, 0))

        # Row 2: XY coords + Pick
        self._xy_frame = tk.Frame(inner, bg=self.P['mantle'])
        self._xy_frame.pack(fill='x', pady=(0, 3))

        self._xy_entries = []
        for lbl_text, var in [('X1', self.var_x1), ('Y1', self.var_y1),
                               ('X2', self.var_x2), ('Y2', self.var_y2)]:
            pair = tk.Frame(self._xy_frame, bg=self.P['mantle'])
            pair.pack(side='left', padx=(0, 6))
            self._lbl(pair, lbl_text, fg=self.P['subtext0'],
                      font=('Segoe UI', 7)).pack(side='left', padx=(0, 2))
            e = self._entry(pair, var, width=5)
            e.pack(side='left')
            self._xy_entries.append(e)

        tk.Button(self._xy_frame, text='+ Pick',
                  bg=self.P['surface0'], fg=self.P['lavender'],
                  activebackground=self.P['surface1'],
                  activeforeground=self.P['lavender'],
                  font=('Segoe UI', 7, 'bold'),
                  relief='flat', bd=0, padx=5, pady=2,
                  cursor='hand2', command=self._pick_coords
                  ).pack(side='left', padx=(4, 0))

        self._pick_overlay = None
        self._pick_points  = []
        self._toggle_xy()

        # Rows 3-4: START/STOP + Save/Load — unified 2×2 grid, equal sizing
        btns = tk.Frame(inner, bg=self.P['mantle'])
        btns.pack(fill='x', pady=(4, 0))
        btns.columnconfigure(0, weight=1)
        btns.columnconfigure(1, weight=1)

        _BF = ('Segoe UI', 9, 'bold')   # shared font for all 4 buttons
        _PY = 5                          # shared pady

        self._btn_start = self._flat_btn(
            btns, '▶  START', self.P['green'], self.P['crust'], self._start,
            font=_BF, pady=_PY)
        self._btn_start.grid(row=0, column=0, sticky='ew',
                              padx=(0, 2), pady=(0, 2))

        self._btn_stop = self._flat_btn(
            btns, '■  STOP', self.P['surface0'], self.P['overlay0'], self._stop,
            font=_BF, pady=_PY)
        self._btn_stop.configure(state='disabled',
                                  activebackground=self.P['surface0'])
        self._btn_stop.grid(row=0, column=1, sticky='ew',
                             padx=(2, 0), pady=(0, 2))

        for col, (label, cmd) in enumerate([('Save Config', self._save_config),
                                             ('Load Config', self._load_config)]):
            tk.Button(btns, text=label, command=cmd,
                      bg=self.P['surface0'], fg=self.P['subtext0'],
                      activebackground=self.P['surface1'],
                      activeforeground=self.P['text'],
                      font=_BF, relief='flat', bd=0,
                      padx=0, pady=_PY, cursor='hand2'
                      ).grid(row=1, column=col, sticky='ew',
                             padx=(0, 2) if col == 0 else (2, 0))

    # ── Data tab ──────────────────────────────────────────────────────────────
    def _build_data_tab(self):
        F = tk.Frame(self, bg=self.P['base'])
        self._tab_frames['Data'] = F

        _, self._lbl_mwd = self._build_sensor_card(F, 'MWD', self.P['green'])
        self._rss_badge, self._lbl_rss = self._build_sensor_card(
            F, 'RSS', self.P['blue'])

    # ── Sensor card (compact single-row) ──────────────────────────────────────
    def _build_sensor_card(self, parent, tool_name, accent):
        outer = tk.Frame(parent, bg=accent)
        outer.pack(fill='x', padx=6, pady=2)

        body = tk.Frame(outer, bg=self.P['mantle'])
        body.pack(fill='both', expand=True, padx=(3, 0))

        row = tk.Frame(body, bg=self.P['mantle'])
        row.pack(fill='x', padx=8, pady=5)

        badge_bg = tk.Frame(row, bg=accent)
        badge_bg.pack(side='left', padx=(0, 10))
        badge_lbl = tk.Label(badge_bg, text=tool_name,
                             bg=accent, fg=self.P['crust'],
                             font=('Segoe UI', 8, 'bold'), padx=7, pady=2)
        badge_lbl.pack()

        value_labels = []
        for i, title in enumerate(['DEPTH', 'INC', 'AZI']):
            if i > 0:
                tk.Frame(row, bg=self.P['surface0'],
                         width=1).pack(side='left', fill='y', padx=5)
            col = tk.Frame(row, bg=self.P['mantle'])
            col.pack(side='left', fill='x', expand=True)
            tk.Label(col, text=title,
                     bg=self.P['mantle'], fg=self.P['overlay0'],
                     font=('Segoe UI', 7, 'bold'), anchor='center').pack(fill='x')
            val = tk.Label(col, text='---',
                           bg=self.P['mantle'], fg=accent,
                           font=('Consolas', 14, 'bold'), anchor='center')
            val.pack(fill='x')
            value_labels.append(val)

        return badge_lbl, value_labels

    # ── Status bar ────────────────────────────────────────────────────────────
    def _build_statusbar(self):
        self._hsep()
        sb = tk.Frame(self, bg=self.P['mantle'])
        sb.pack(fill='x')

        self._dot_canvas = tk.Canvas(sb, width=8, height=8,
                                     bg=self.P['mantle'], highlightthickness=0)
        self._dot_canvas.pack(side='left', padx=(10, 5), pady=5)
        self._dot = self._dot_canvas.create_oval(
            0, 0, 8, 8, fill=self.P['overlay0'], outline='')

        self._status_var = tk.StringVar(value='Ready')
        tk.Label(sb, textvariable=self._status_var,
                 bg=self.P['mantle'], fg=self.P['subtext0'],
                 font=('Segoe UI', 8), anchor='w').pack(side='left')

    # ── Log ───────────────────────────────────────────────────────────────────
    def _build_log(self):
        self._hsep()
        log_outer = tk.Frame(self, bg=self.P['base'])
        log_outer.pack(fill='both', expand=True, padx=6, pady=(0, 4))

        self._log = scrolledtext.ScrolledText(
            log_outer, height=6, font=('Consolas', 8),
            bg=self.P['crust'], fg=self.P['subtext0'],
            insertbackground=self.P['text'],
            state='disabled', wrap='word',
            relief='flat', bd=0, padx=10, pady=8)
        self._log.pack(fill='both', expand=True)

        self._log_append(f'Read Screen Utility  v{VERSION}  |  {DATE}\n')
        self._log_append('-' * 38 + '\n')

    # ── UI helpers ────────────────────────────────────────────────────────────
    def _toggle_xy(self):
        state = 'normal' if self.var_locate.get() == 'manual' else 'disabled'
        for e in self._xy_entries:
            e.configure(state=state)

    # ── Coordinate picker ─────────────────────────────────────────────────────
    def _pick_coords(self):
        """Minimise window then open fullscreen click-to-capture overlay."""
        self._pick_points = []
        self.iconify()
        self.after(300, self._pick_overlay_start)

    def _pick_overlay_start(self):
        import ctypes
        # Cover the entire virtual desktop (multi-monitor safe)
        vx = ctypes.windll.user32.GetSystemMetrics(76)   # SM_XVIRTUALSCREEN
        vy = ctypes.windll.user32.GetSystemMetrics(77)   # SM_YVIRTUALSCREEN
        vw = ctypes.windll.user32.GetSystemMetrics(78)   # SM_CXVIRTUALSCREEN
        vh = ctypes.windll.user32.GetSystemMetrics(79)   # SM_CYVIRTUALSCREEN

        ov = tk.Toplevel(self)
        ov.overrideredirect(True)
        ov.geometry(f'{vw}x{vh}+{vx}+{vy}')
        ov.attributes('-topmost', True)
        ov.attributes('-alpha', 0.18)
        ov.configure(bg='black', cursor='crosshair')
        self._pick_overlay = ov

        # Banner shown at top of screen
        banner = tk.Frame(ov, bg=self.P['base'])
        banner.pack(fill='x')

        self._pick_label = tk.Label(
            banner,
            text='Click TOP-LEFT corner of survey data area    ESC = cancel',
            bg=self.P['base'], fg=self.P['green'],
            font=('Consolas', 13, 'bold'), pady=10, padx=20
        )
        self._pick_label.pack(side='left')

        self._pick_xy_label = tk.Label(
            banner, text='X: ---   Y: ---',
            bg=self.P['base'], fg=self.P['blue'],
            font=('Consolas', 13), pady=10, padx=20
        )
        self._pick_xy_label.pack(side='right')

        ov.bind('<Motion>',   self._on_pick_motion)
        ov.bind('<Button-1>', self._on_pick_click)
        ov.bind('<Escape>',   lambda e: self._pick_cancel())
        ov.focus_force()

    def _on_pick_motion(self, event):
        self._pick_xy_label.config(
            text=f'X: {event.x_root}   Y: {event.y_root}'
        )

    def _on_pick_click(self, event):
        self._pick_points.append((event.x_root, event.y_root))
        if len(self._pick_points) == 1:
            self._pick_label.config(
                text='Click BOTTOM-RIGHT corner of survey data area    ESC = cancel',
                fg=self.P['peach']
            )
        elif len(self._pick_points) == 2:
            self._pick_finish()

    def _pick_cancel(self):
        if self._pick_overlay:
            self._pick_overlay.destroy()
            self._pick_overlay = None
        self.deiconify()
        self._log_append('Coordinate pick cancelled.\n')

    def _pick_finish(self):
        if self._pick_overlay:
            self._pick_overlay.destroy()
            self._pick_overlay = None

        (x1, y1), (x2, y2) = self._pick_points
        # Normalise so x1,y1 is always the smaller corner
        x1, x2 = min(x1, x2), max(x1, x2)
        y1, y2 = min(y1, y2), max(y1, y2)

        self.var_x1.set(str(x1))
        self.var_y1.set(str(y1))
        self.var_x2.set(str(x2))
        self.var_y2.set(str(y2))

        # Auto-switch to Manual mode and enable entries
        self.var_locate.set('manual')
        self._toggle_xy()

        self.deiconify()
        self._log_append(
            f'Coordinates captured: X1={x1} Y1={y1}  X2={x2} Y2={y2}\n'
        )

    def _log_append(self, text):
        self._log.configure(state='normal')
        self._log.insert('end', text)
        try:
            if int(self._log.index('end-1c').split('.')[0]) > 400:
                self._log.delete('1.0', '80.0')
        except Exception:
            pass
        self._log.yview('end')
        self._log.configure(state='disabled')

    def _set_dot(self, color):
        self._dot_canvas.itemconfigure(self._dot, fill=color)

    def _update_data(self, mwd, rss):
        for lbl, val in zip(self._lbl_mwd, mwd):
            lbl.configure(text=fmt_val(val))
        for lbl, val in zip(self._lbl_rss, rss):
            lbl.configure(text=fmt_val(val))
        # Flash white briefly on update
        for lbl in self._lbl_mwd + self._lbl_rss:
            lbl.configure(fg='#ffffff')
        self.after(80, self._restore_value_colors)

    def _restore_value_colors(self):
        for lbl in self._lbl_mwd:
            lbl.configure(fg=self.P['green'])
        for lbl in self._lbl_rss:
            lbl.configure(fg=self.P['blue'])

    # ── Config ────────────────────────────────────────────────────────────────
    def _save_config(self):
        cfg = {
            'method':           self.var_method.get(),
            'tesseract_config': '--psm 6 --oem 1',
            'scale_factor':     self.var_scale.get(),
            'interval':         self.var_interval.get(),
            'tool':             self.var_tool.get(),
            'locate':           self.var_locate.get(),
            'loc_x1':           self.var_x1.get(),
            'loc_y1':           self.var_y1.get(),
            'loc_x2':           self.var_x2.get(),
            'loc_y2':           self.var_y2.get(),
        }
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(cfg, f, indent=4)
            self._log_append(f'Config saved: {CONFIG_FILE}\n')
        except Exception as e:
            messagebox.showerror('Save Error', str(e))

    def _load_config(self):
        try:
            with open(CONFIG_FILE) as f:
                cfg = json.load(f)
            self.var_method.set(  cfg.get('method',       'replace'))
            self.var_scale.set(   cfg.get('scale_factor', '3'))
            self.var_interval.set(cfg.get('interval',     '2'))
            self.var_tool.set(    cfg.get('tool',         'rss'))
            self.var_locate.set(  cfg.get('locate',       'auto'))
            self.var_x1.set(    cfg.get('loc_x1',       '10'))
            self.var_y1.set(    cfg.get('loc_y1',       '10'))
            self.var_x2.set(    cfg.get('loc_x2',       '100'))
            self.var_y2.set(    cfg.get('loc_y2',       '100'))
            self._toggle_xy()
            self._log_append(
                f'Config: method={cfg.get("method")}  '
                f'tool={cfg.get("tool","rss")}  '
                f'scale={cfg.get("scale_factor","3")}\n')
        except FileNotFoundError:
            self._log_append(f'{CONFIG_FILE} not found — using defaults\n')
        except Exception as e:
            self._log_append(f'Load config error: {e}\n')

    # ── Start / Stop ──────────────────────────────────────────────────────────
    def _start(self):
        if self._worker and self._worker.is_alive():
            return

        def _int(v, default):
            try:    return int(v)
            except: return default

        config = {
            'method':    self.var_method.get(),
            'tool':      self.var_tool.get(),
            'locate':    self.var_locate.get(),
            'scale':     max(1, min(10, _int(self.var_scale.get(), 3))),
            'interval':  max(1, min(60, _int(self.var_interval.get(), 2))),
            'x1':        _int(self.var_x1.get(), 10),
            'y1':        _int(self.var_y1.get(), 10),
            'x2':        _int(self.var_x2.get(), 100),
            'y2':        _int(self.var_y2.get(), 100),
        }

        self._stop_event.clear()
        self._btn_start.configure(state='disabled',
                                   bg=self.P['surface0'], fg=self.P['overlay0'],
                                   activebackground=self.P['surface0'])
        self._btn_stop.configure(state='normal',
                                  bg=self.P['red'], fg=self.P['crust'],
                                  activebackground=self._lighten(self.P['red']))
        self._set_dot(self.P['yellow'])
        self._show_tab('Data')
        self._worker = threading.Thread(target=self._worker_loop,
                                        args=(config,), daemon=True)
        self._worker.start()

    def _stop(self):
        self._stop_event.set()
        self._btn_stop.configure(state='disabled',
                                  bg=self.P['surface0'], fg=self.P['overlay0'],
                                  activebackground=self.P['surface0'])
        self._status_var.set('Stopping...')
        self._set_dot(self.P['yellow'])

    def _on_close(self):
        self._stop_event.set()
        self.destroy()

    # ── Queue polling (main thread) ───────────────────────────────────────────
    def _poll_queue(self):
        try:
            while True:
                msg = self._queue.get_nowait()
                kind = msg[0]
                if kind == 'log':
                    self._log_append(msg[1])
                elif kind == 'status':
                    self._status_var.set(msg[1])
                    if 'update' in msg[1]:
                        self._set_dot(self.P['green'])
                    elif 'Error' in msg[1]:
                        self._set_dot(self.P['red'])
                    elif 'Waiting' in msg[1] or 'Click' in msg[1]:
                        self._set_dot(self.P['yellow'])
                elif kind == 'data':
                    self._update_data(msg[1], msg[2])
                elif kind == 'stopped':
                    self._btn_start.configure(
                        state='normal',
                        bg=self.P['green'], fg=self.P['crust'],
                        activebackground=self._lighten(self.P['green']))
                    self._btn_stop.configure(
                        state='disabled',
                        bg=self.P['surface0'], fg=self.P['overlay0'],
                        activebackground=self.P['surface0'])
                    self._status_var.set('Stopped')
                    self._set_dot(self.P['overlay0'])
                    self._show_tab('Setup')
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _q(self, *args):
        self._queue.put(args)

    # ── Worker thread ─────────────────────────────────────────────────────────
    def _worker_loop(self, config):
        method    = config['method']
        tool      = config['tool']
        locate    = config['locate']
        scale     = config['scale']
        interval  = config['interval']
        tess_opt  = '--psm 6 --oem 1 -c tessedit_char_whitelist=0123456789.MR'
        rss_label = 'RSS' if tool == 'rss' else 'MTR'

        last_mwd = ['0.00', '0.00', '0.00']
        last_rss = ['0.00', '0.00', '0.00']

        # ── Auto window detection ──
        win_name = None
        if locate == 'auto':
            self._q('status', 'Click on the Rig Floor Console window...')
            self._q('log', 'Waiting for Rig Floor Console window...\n')
            while not self._stop_event.is_set():
                name = win32gui.GetWindowText(win32gui.GetForegroundWindow())
                if 'Rig Floor Console -' in name:
                    win_name = name
                    self._q('log', f'Window found: {win_name}\n')
                    break
                time.sleep(0.5)
            if self._stop_event.is_set():
                self._q('stopped')
                return

        self._q('status', 'Running...')
        self._q('log',
                f'Started | method={method}  tool={tool}  '
                f'scale={scale}  locate={locate}\n')
        self._q('log', '-' * 38 + '\n')

        # ── Main capture loop ──
        while not self._stop_event.is_set():
            try:
                # Capture
                if locate == 'manual':
                    img = ImageGrab.grab(
                        (config['x1'], config['y1'],
                         config['x2'], config['y2']),
                        all_screens=True)
                else:
                    hwnd = win32gui.FindWindow(None, win_name)
                    if not hwnd:
                        self._q('log', 'Window lost — retrying...\n')
                        time.sleep(2)
                        continue
                    b = win32gui.GetWindowRect(hwnd)
                    img = ImageGrab.grab(
                        (b[0]+7, b[1], b[2]-7, b[3]-7), all_screens=True)
                    w, h = img.size
                    if tool == 'rss':
                        # MD/INC/AZI cols only: x=45.5-78%, y=92.0-99.9%
                        crop = (w - round(w*0.545), round(h*0.920),
                                w - round(w*0.22),  round(h*0.999))
                    else:
                        # MD/INC/AZI cols only: x=40-76%, y=94.8-99.9%
                        crop = (w - round(w*0.60),  round(h*0.948),
                                w - round(w*0.24),  round(h*0.999))
                    img = img.crop(crop)

                img.save('screenshot.png')

                # Preprocess (same as main-auto.py)
                cv_img = cv2.imread('screenshot.png')
                if cv_img is None:
                    raise FileNotFoundError('Cannot read screenshot.png')
                cv_img = cv2.resize(cv_img, None, fx=scale, fy=scale,
                                    interpolation=cv2.INTER_LANCZOS4)
                cv_img = cv2.filter2D(cv_img, -1, _sharpen_kernel)

                if method == 'replace':
                    hsv  = cv2.cvtColor(cv_img, cv2.COLOR_BGR2HSV)
                    mask = cv2.inRange(hsv,
                                       np.array([30,  80,  80]),
                                       np.array([100, 255, 255]))
                    out  = np.full_like(cv_img, 255)
                    out[mask > 0] = 0
                    out  = cv2.copyMakeBorder(out, 20, 20, 20, 20,
                                              cv2.BORDER_CONSTANT, value=[255, 255, 255])
                    pil  = Image.fromarray(cv2.cvtColor(out, cv2.COLOR_BGR2RGB))

                elif method == 'threshold':
                    cv_img = cv2.bilateralFilter(cv_img, 5, 75, 75)
                    gry    = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
                    thr    = cv2.threshold(gry, 0, 255,
                                           cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
                    thr    = cv2.morphologyEx(
                        thr, cv2.MORPH_OPEN,
                        cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2)))
                    thr    = cv2.copyMakeBorder(thr, 20, 20, 20, 20,
                                                cv2.BORDER_CONSTANT, value=255)
                    pil    = Image.fromarray(thr)

                else:
                    cv_img = cv2.copyMakeBorder(cv_img, 20, 20, 20, 20,
                                                cv2.BORDER_CONSTANT, value=[0, 0, 0])
                    pil    = Image.fromarray(cv2.cvtColor(cv_img, cv2.COLOR_BGR2RGB))

                # OCR
                raw = tess.image_to_string(pil, config=tess_opt)
                raw = re.sub(r'[^\d\.MR\n]', '', raw)
                raw = raw.replace('.M', 'M').replace('.R', 'R').replace('R.', 'M')
                lines = [l for l in raw.split('\n') if l.strip()]

                mwd = (parse_survey_line(lines[0], 'M')
                       if len(lines) > 0 else ['0.00', '0.00', '0.00'])
                # Only parse RSS row when tool is RSS; motor has only one data row
                if tool == 'rss':
                    rss = (parse_survey_line(lines[1], 'R')
                           if len(lines) > 1 else ['0.00', '0.00', '0.00'])
                else:
                    rss = ['0.00', '0.00', '0.00']

                mwd = data_check(mwd)
                if tool == 'rss':
                    rss = data_check(rss)

                # Last-known-good
                def resolve(lst, last):
                    if lst[0] in ('9.99', '', '0.00'):
                        return last[:], False
                    try:
                        float(lst[0]); float(lst[1]); float(lst[2])
                        return lst[:], True
                    except Exception:
                        return last[:], False

                mwd_out, mwd_ok = resolve(mwd, last_mwd)
                if mwd_ok: last_mwd = mwd_out[:]

                if tool == 'rss':
                    rss_out, rss_ok = resolve(rss, last_rss)
                    if rss_ok: last_rss = rss_out[:]
                else:
                    rss_out = ['--', '--', '--']

                # Write CSV
                for _ in range(10):
                    try:
                        with open('output.csv', 'w', newline='') as f:
                            wr = csv.writer(f)
                            wr.writerow(mwd_out)
                            if tool == 'rss':
                                wr.writerow(rss_out)
                        break
                    except Exception:
                        time.sleep(0.1)

                # Update GUI
                ts      = time.strftime('%H:%M:%S')
                mwd_str = ' / '.join(fmt_val(v) for v in mwd_out)
                rss_str = ' / '.join(fmt_val(v) for v in rss_out)
                self._q('data', mwd_out, rss_out)
                self._q('status', f'Last update: {ts}  |  interval: {interval}s')
                if tool == 'rss':
                    self._q('log',
                            f'[{ts}]\nMWD {mwd_str}\n{rss_label} {rss_str}\n')
                else:
                    self._q('log', f'[{ts}]\nMWD {mwd_str}\n')

            except Exception as e:
                ts = time.strftime('%H:%M:%S')
                self._q('log',   f'[{ts}] Error: {e}\n')
                self._q('status', 'Error — retrying...')

            # Responsive wait
            for _ in range(interval * 10):
                if self._stop_event.is_set():
                    break
                time.sleep(0.1)

        self._q('stopped')


if __name__ == '__main__':
    App().mainloop()
