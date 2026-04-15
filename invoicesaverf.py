"""
╔══════════════════════════════════════════════════════════════════════╗
║           STYLUXE INVOICE PRO — Complete Single File                ║
║           OCR • Smart Date Extract • Analytics • Export             ║
║                                                                      ║
║  INSTALL REQUIREMENTS FIRST:                                         ║
║    pip install opencv-python pillow pytesseract pyperclip            ║
║                                                                      ║
║  ALSO INSTALL TESSERACT ENGINE:                                      ║
║    https://github.com/UB-Mannheim/tesseract/wiki                     ║
║    Default install path: C:\\Program Files\\Tesseract-OCR\\tesseract.exe ║
╚══════════════════════════════════════════════════════════════════════╝
"""

# ═══════════════════════════════════════════════════════════════════════
#  IMPORTS
# ═══════════════════════════════════════════════════════════════════════
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import cv2
import numpy as np
from PIL import Image, ImageTk
import pytesseract
import sqlite3
import re
import csv
import os
import threading
import pyperclip
from datetime import datetime

# ═══════════════════════════════════════════════════════════════════════
#  CONFIGURATION  —  change Tesseract path here if needed
# ═══════════════════════════════════════════════════════════════════════
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH  = os.path.join(BASE_DIR, "styluxe_invoices.db")
SAVE_DIR = os.path.join(BASE_DIR, "saved_invoices")
CFG_PATH = os.path.join(BASE_DIR, "styluxe_config.txt")

# ═══════════════════════════════════════════════════════════════════════
#  COLOUR PALETTE
# ═══════════════════════════════════════════════════════════════════════
C = {
    "bg_dark"   : "#0d1117",
    "bg_mid"    : "#161b22",
    "bg_card"   : "#21262d",
    "border"    : "#30363d",
    "accent1"   : "#58a6ff",
    "accent2"   : "#7ee787",
    "accent3"   : "#f78166",
    "accent4"   : "#d2a8ff",
    "accent5"   : "#e3b341",
    "text_main" : "#e6edf3",
    "text_muted": "#8b949e",
    "white"     : "#ffffff",
}

# ═══════════════════════════════════════════════════════════════════════
#  SMART DATE EXTRACTOR  (handles all OCR noise)
# ═══════════════════════════════════════════════════════════════════════

MONTH_NAMES = (
    r'(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|'
    r'May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|'
    r'Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)'
)

_DATE_PATTERNS = [
    # DD-Mon-YYYY  e.g. 08-Mar-2026
    (r'\b(\d{1,2}[.\/\-]' + MONTH_NAMES + r'[.\/\-]\d{2,4})\b', 1),
    # Month DD, YYYY  e.g. March 31, 2026
    (r'\b(' + MONTH_NAMES + r'\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{2,4})\b', 1),
    # DD Month YYYY  e.g. 31 March 2026
    (r'\b(\d{1,2}(?:st|nd|rd|th)?\s+' + MONTH_NAMES + r'\s+\d{2,4})\b', 1),
    # YYYY-MM-DD  ISO format
    (r'\b(\d{4}[-\/]\d{2}[-\/]\d{2})\b', 1),
    # Date keyword + DD/MM/YYYY same line
    (r'(?:invoice\s*)?[Dd]ate\s*[:\-\u2014]?\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})', 1),
    # Date keyword + YYYY-MM-DD
    (r'(?:invoice\s*)?[Dd]ate\s*[:\-]?\s*(\d{4}[-\/]\d{2}[-\/]\d{2})', 1),
    # Date keyword + DD-Mon-YYYY
    (r'(?:invoice\s*)?[Dd]ate\s*[:\-]?\s*(\d{1,2}[.\/\-]' + MONTH_NAMES + r'[.\/\-]\d{2,4})', 1),
    # Order Date: Month DD, YYYY  (Daraz style)
    (r'[Oo]rder\s+[Dd]ate\s*[:\-]?\s*(' + MONTH_NAMES + r'\s+\d{1,2},?\s+\d{4})', 1),
    # Bill/Due/Issue Date
    (r'(?:[Bb]ill|[Dd]ue|[Ii]ssue)\s*[Dd]ate\s*[:\-]?\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})', 1),
    # Bare DD/MM/YYYY or DD-MM-YY (last resort, most generic)
    (r'\b(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})\b', 1),
]

def _fix_ocr_noise(text):
    """Fix common Tesseract character misreads near date tokens."""
    lines_out = []
    for line in text.splitlines():
        if re.search(r'\d', line):
            # Collapse spaces between digits  e.g. "3 1 / 0 3" → "31/03"
            line = re.sub(r'(\d)\s+(?=\d)', r'\1', line)
            # Fix O→0, l/I→1, S→5, B→8, backslash→slash near digit separators
            line = re.sub(
                r'[OolISBG\|](?=[\d\/\-\.])|(?<=[\d\/\-\.])[OolISBG\|]',
                lambda m: m.group(0).translate(
                    str.maketrans('OolISBG\\|', '001558601')),
                line
            )
            line = line.replace('\\', '/')
        lines_out.append(line)
    return '\n'.join(lines_out)

def _window_around_keyword(text, keyword, window=3):
    """Return ±window lines around any line containing keyword."""
    lines = text.splitlines()
    result = []
    for i, line in enumerate(lines):
        if re.search(keyword, line, re.IGNORECASE):
            result.extend(lines[max(0, i-1): min(len(lines), i+window+1)])
    return '\n'.join(result)

def _validate_date(s):
    """Reject strings that look like phone numbers, NTN codes, prices etc."""
    if not s:
        return False
    has_sep   = bool(re.search(r'[\/\-\.]', s))
    has_month = bool(re.search(MONTH_NAMES, s, re.IGNORECASE))
    if not (has_sep or has_month):
        return False
    digits = re.findall(r'\d+', s)
    if len(digits) < 2:
        return False
    for d in digits:
        if len(d) == 4:
            yr = int(d)
            if not (1990 <= yr <= 2099):
                return False
    return True

def _run_patterns(text):
    for pat, grp in _DATE_PATTERNS:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(grp).strip()
    return None

def extract_invoice_date(raw_text):
    """
    Master function — tries 4 strategies in order, returns best date found.
    Strategy 1: context window around 'date' keyword in original text
    Strategy 2: context window around 'date' keyword in OCR-cleaned text
    Strategy 3: full original text
    Strategy 4: full OCR-cleaned text
    """
    if not raw_text or not raw_text.strip():
        return "Not Found"

    cleaned = _fix_ocr_noise(raw_text)
    date_kw  = r'(?:invoice\s*)?date|order\s*date|bill\s*date|due\s*date|issue\s*date'

    for search_text in [
        _window_around_keyword(raw_text, date_kw),
        _window_around_keyword(cleaned,  date_kw),
        raw_text,
        cleaned,
    ]:
        if not search_text.strip():
            continue
        result = _run_patterns(search_text)
        if result and _validate_date(result):
            return result

    return "Not Found"

# ═══════════════════════════════════════════════════════════════════════
#  TOOLTIP HELPER
# ═══════════════════════════════════════════════════════════════════════
class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text   = text
        self.tip    = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, _=None):
        x = self.widget.winfo_rootx() + 40
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self.tip, text=self.text, background="#2d333b",
                 foreground=C["text_muted"], relief="flat",
                 font=("Consolas", 9), padx=8, pady=4).pack()

    def hide(self, _=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None

# ═══════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ═══════════════════════════════════════════════════════════════════════
class InvoiceApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Styluxe Invoice Pro")
        self.root.geometry("1100x700")
        self.root.configure(bg=C["bg_dark"])
        self.root.resizable(True, True)

        self.current_image   = None
        self.processed_image = None
        self.last_ocr_text   = ""
        self.rotation_angle  = 0

        os.makedirs(SAVE_DIR, exist_ok=True)
        self._load_config()
        self._init_db()
        self._setup_styles()
        self._build_ui()
        self._bind_shortcuts()
    
    # ─────────────────────── CONFIG ───────────────────────────────────
    def _load_config(self):
        if os.path.exists(CFG_PATH):
            with open(CFG_PATH, "r") as f:
                for line in f:
                    if line.startswith("TESSERACT_PATH="):
                        p = line.split("=", 1)[1].strip()
                        if p:
                            pytesseract.pytesseract.tesseract_cmd = p

    def _save_config(self):
        with open(CFG_PATH, "w") as f:
            f.write(f"TESSERACT_PATH={pytesseract.pytesseract.tesseract_cmd}\n")

    # ─────────────────────── DATABASE ─────────────────────────────────
    def _init_db(self):
        conn = sqlite3.connect(DB_PATH)
        try:
            cur = conn.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS invoices (
                    id           INTEGER PRIMARY KEY AUTOINCREMENT,
                    scan_date    DATETIME DEFAULT CURRENT_TIMESTAMP,
                    invoice_date TEXT,
                    total_amount TEXT,
                    raw_text     TEXT,
                    image_path   TEXT
                )""")
            try:
                cur.execute("ALTER TABLE invoices ADD COLUMN image_path TEXT")
            except sqlite3.OperationalError:
                pass
            conn.commit()
        finally:
            conn.close()

    # ─────────────────────── TTK STYLES ───────────────────────────────
    def _setup_styles(self):
        s = ttk.Style()
        s.theme_use("clam")
        s.configure("Treeview",
                    background=C["bg_card"], foreground=C["text_main"],
                    fieldbackground=C["bg_card"], rowheight=28,
                    font=("Consolas", 10))
        s.configure("Treeview.Heading",
                    background=C["bg_mid"], foreground=C["accent1"],
                    font=("Consolas", 10, "bold"), relief="flat")
        s.map("Treeview",
              background=[("selected", C["accent1"])],
              foreground=[("selected", C["bg_dark"])])
        s.configure("Horizontal.TProgressbar",
                    troughcolor=C["bg_card"], background=C["accent1"],
                    lightcolor=C["accent1"], darkcolor=C["accent1"],
                    bordercolor=C["border"])

    # ─────────────────────── MAIN UI ──────────────────────────────────
    def _build_ui(self):
        # Header
        hdr = tk.Frame(self.root, bg=C["bg_mid"], height=52)
        hdr.pack(fill="x")
        tk.Label(hdr, text="  ◈  STYLUXE  INVOICE  PRO",
                 font=("Courier New", 15, "bold"),
                 bg=C["bg_mid"], fg=C["accent1"]).pack(side="left", pady=12)
        tk.Label(hdr, text="v2.0  •  OCR + Smart Date + Analytics",
                 font=("Courier New", 9),
                 bg=C["bg_mid"], fg=C["text_muted"]).pack(side="left", padx=8)
        tk.Button(hdr, text="⚙", command=self.open_settings,
                  bg=C["bg_mid"], fg=C["text_muted"],
                  font=("Arial", 14), relief="flat", bd=0,
                  cursor="hand2", activebackground=C["bg_card"],
                  activeforeground=C["accent4"]).pack(side="right", padx=12)

        # Body
        body = tk.Frame(self.root, bg=C["bg_dark"])
        body.pack(fill="both", expand=True)
        self._build_sidebar(body)

        right = tk.Frame(body, bg=C["bg_dark"])
        right.pack(side="right", fill="both", expand=True, padx=10, pady=10)
        self._build_stats_bar(right)
        self._build_canvas_area(right)
        self._build_progress(right)

        # Status bar
        self.status_var = tk.StringVar(value="  ◉  System Ready. Awaiting input…")
        tk.Label(self.root, textvariable=self.status_var,
                 bg=C["bg_mid"], fg=C["text_muted"],
                 font=("Consolas", 9), anchor="w").pack(
                     side="bottom", fill="x", ipady=4)

    def _build_sidebar(self, parent):
        sb = tk.Frame(parent, bg=C["bg_mid"], width=235)
        sb.pack(side="left", fill="y", padx=(8, 0), pady=8)
        sb.pack_propagate(False)

        tk.Label(sb, text="ACTIONS", font=("Consolas", 9, "bold"),
                 bg=C["bg_mid"], fg=C["text_muted"]).pack(
                     pady=(18, 6), padx=14, anchor="w")

        buttons = [
            ("📸  Capture from Camera",  self.capture_image,         C["accent1"], "Use your webcam to capture invoice"),
            ("📁  Upload Invoice Image",  self.upload_image,          C["accent1"], "Ctrl+O  ·  Pick an image file"),
            ("🛠   Tune & Enhance Image",  self.open_enhancement_tool, C["accent4"], "Adjust contrast, brightness, rotation"),
            ("⚙   Extract & Save Data",   self.process_and_save,      C["accent2"], "Ctrl+S  ·  Run OCR and save to database"),
            ("📋  Copy Last OCR Text",    self.copy_ocr_text,         C["accent5"], "Copy extracted text to clipboard"),
            ("📊  View & Export Records", self.view_database,         C["accent5"], "Ctrl+D  ·  Browse all saved invoices"),
        ]

        for label, cmd, color, tip in buttons:
            btn = tk.Button(sb, text=label, command=cmd,
                            bg=C["bg_card"], fg=color,
                            font=("Consolas", 10, "bold"),
                            relief="flat", bd=0, pady=10,
                            cursor="hand2", activebackground=C["border"],
                            activeforeground=C["white"],
                            anchor="w", padx=14)
            btn.pack(fill="x", padx=10, pady=3)
            Tooltip(btn, tip)
            btn.bind("<Enter>", lambda e, b=btn: b.config(bg=C["border"]))
            btn.bind("<Leave>", lambda e, b=btn: b.config(bg=C["bg_card"]))

        ttk.Separator(sb, orient="horizontal").pack(fill="x", padx=14, pady=14)

        tk.Label(sb, text="SHORTCUTS", font=("Consolas", 9, "bold"),
                 bg=C["bg_mid"], fg=C["text_muted"]).pack(
                     pady=(0, 4), padx=14, anchor="w")
        for key, desc in [("Ctrl+O","Upload"), ("Ctrl+S","Extract"),
                           ("Ctrl+D","Database"), ("Ctrl+Q","Quit")]:
            row = tk.Frame(sb, bg=C["bg_mid"])
            row.pack(fill="x", padx=14, pady=1)
            tk.Label(row, text=key, font=("Consolas", 8, "bold"),
                     bg=C["border"], fg=C["accent1"],
                     padx=4, pady=1).pack(side="left")
            tk.Label(row, text=f"  {desc}", font=("Consolas", 8),
                     bg=C["bg_mid"], fg=C["text_muted"]).pack(side="left")

    def _build_stats_bar(self, parent):
        sf = tk.Frame(parent, bg=C["bg_mid"],
                      highlightbackground=C["border"], highlightthickness=1)
        sf.pack(fill="x", pady=(0, 8))
        self.stat_vars = {
            "total_invoices": tk.StringVar(value="—"),
            "total_amount"  : tk.StringVar(value="—"),
            "last_scan"     : tk.StringVar(value="—"),
        }
        for title, key, color in [
            ("📄  Total Invoices", "total_invoices", C["accent1"]),
            ("💰  Sum of Totals",  "total_amount",   C["accent2"]),
            ("🕐  Last Scan",      "last_scan",       C["accent4"]),
        ]:
            cell = tk.Frame(sf, bg=C["bg_mid"])
            cell.pack(side="left", expand=True, fill="both", padx=2, pady=2)
            tk.Label(cell, text=title, font=("Consolas", 8),
                     bg=C["bg_mid"], fg=C["text_muted"]).pack(
                         anchor="w", padx=10, pady=(6, 0))
            tk.Label(cell, textvariable=self.stat_vars[key],
                     font=("Courier New", 13, "bold"),
                     bg=C["bg_mid"], fg=color).pack(anchor="w", padx=10, pady=(0, 6))
        self._refresh_stats()

    def _refresh_stats(self):
        conn = sqlite3.connect(DB_PATH)
        try:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*), MAX(scan_date) FROM invoices")
            count, last = cur.fetchone()
            cur.execute("SELECT total_amount FROM invoices")
            totals = cur.fetchall()
        finally:
            conn.close()
        total_sum = 0.0
        for (t,) in totals:
            if t and t != "Not Found":
                try:
                    total_sum += float(t.replace(",", "").strip())
                except ValueError:
                    pass
        self.stat_vars["total_invoices"].set(str(count or 0))
        self.stat_vars["total_amount"].set(
            f"PKR {total_sum:,.0f}" if total_sum else "—")
        self.stat_vars["last_scan"].set(last[:16] if last else "—")

    def _build_canvas_area(self, parent):
        card = tk.Frame(parent, bg=C["bg_card"],
                        highlightbackground=C["border"], highlightthickness=1)
        card.pack(fill="both", expand=True)
        self._placeholder = tk.Label(
            card, text="[ Upload or Capture an Invoice Image ]",
            font=("Consolas", 13), bg=C["bg_card"], fg=C["border"])
        self._placeholder.place(relx=0.5, rely=0.5, anchor="center")
        self.canvas = tk.Canvas(card, bg=C["bg_card"], highlightthickness=0)
        self.canvas.pack(fill="both", expand=True, padx=4, pady=4)

    def _build_progress(self, parent):
        pf = tk.Frame(parent, bg=C["bg_dark"])
        pf.pack(fill="x", pady=(4, 0))
        self.progress_var = tk.DoubleVar(value=0)
        ttk.Progressbar(pf, variable=self.progress_var,
                         style="Horizontal.TProgressbar",
                         mode="determinate").pack(fill="x")

    def _bind_shortcuts(self):
        self.root.bind("<Control-o>", lambda e: self.upload_image())
        self.root.bind("<Control-s>", lambda e: self.process_and_save())
        self.root.bind("<Control-d>", lambda e: self.view_database())
        self.root.bind("<Control-q>", lambda e: self.root.quit())

    # ─────────────────────── IMAGE HANDLING ───────────────────────────
    def upload_image(self):
        path = filedialog.askopenfilename(
            filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp *.tiff *.webp")])
        if path:
            self.current_image   = cv2.imread(path)
            self.processed_image = None
            self.rotation_angle  = 0
            self._show_image(self.current_image)
            self._set_status(f"  ◉  Loaded: {os.path.basename(path)}", C["accent1"])

    def capture_image(self):
        self._set_status("  ◎  Opening camera…", C["accent4"])
        self.root.update()
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            messagebox.showerror("Camera Error",
                "No camera detected or it is already in use.")
            self._set_status("  ✗  No camera detected.", C["accent3"])
            return
        ret, frame = cap.read()
        cap.release()
        if ret:
            self.current_image   = frame
            self.processed_image = None
            self.rotation_angle  = 0
            self._show_image(frame)
            self._set_status("  ◉  Camera capture successful!", C["accent2"])
        else:
            messagebox.showerror("Camera Error", "Failed to read frame from camera.")
            self._set_status("  ✗  Could not capture frame.", C["accent3"])

    def _show_image(self, img):
        """Display BGR or grayscale image on main canvas, scaled to fit."""
        self._placeholder.place_forget()
        w = self.canvas.winfo_width()  or 800
        h = self.canvas.winfo_height() or 440
        rgb = cv2.cvtColor(img, cv2.COLOR_GRAY2RGB) \
              if len(img.shape) == 2 \
              else cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        pil = Image.fromarray(rgb)
        pil.thumbnail((w, h), Image.LANCZOS)
        self._tk_img = ImageTk.PhotoImage(pil)
        self.canvas.delete("all")
        self.canvas.create_image(w // 2, h // 2, image=self._tk_img)

    # ─────────────────── ENHANCEMENT STUDIO ──────────────────────────
    def open_enhancement_tool(self):
        if self.current_image is None:
            messagebox.showwarning("No Image", "Please load an image first!")
            return

        self.enh_win = tk.Toplevel(self.root)
        self.enh_win.title("Enhancement Studio")
        self.enh_win.geometry("640x800")
        self.enh_win.configure(bg=C["bg_mid"])
        self.enh_win.grab_set()

        tk.Label(self.enh_win, text="ENHANCEMENT STUDIO",
                 font=("Courier New", 13, "bold"),
                 bg=C["bg_mid"], fg=C["accent4"]).pack(pady=14)

        self.enh_canvas = tk.Canvas(self.enh_win, width=580, height=360,
                                    bg=C["bg_dark"], highlightthickness=0)
        self.enh_canvas.pack(pady=6, padx=10)

        self._enh_img_cache = None

        cf = tk.Frame(self.enh_win, bg=C["bg_mid"])
        cf.pack(fill="x", padx=30, pady=6)

        self.val_contrast = tk.DoubleVar(value=1.0)
        self.val_bright   = tk.IntVar(value=0)
        self.val_thresh   = tk.IntVar(value=0)

        sliders = [
            ("Contrast  (Clarity)",         self.val_contrast, 0.5, 3.0, 0.1),
            ("Brightness",                  self.val_bright,  -100, 100, 1),
            ("B/W Threshold  (0 = off)",    self.val_thresh,     0, 255, 1),
        ]
        for i, (lbl, var, lo, hi, res) in enumerate(sliders):
            tk.Label(cf, text=lbl, bg=C["bg_mid"], fg=C["text_main"],
                     font=("Consolas", 10), width=26, anchor="w").grid(
                         row=i, column=0, pady=7, sticky="w")
            tk.Scale(cf, from_=lo, to=hi, resolution=res,
                     orient="horizontal", variable=var,
                     command=self._update_enh_preview,
                     bg=C["bg_mid"], fg=C["text_main"],
                     activebackground=C["accent1"],
                     highlightthickness=0, troughcolor=C["bg_card"],
                     length=350).grid(row=i, column=1, pady=7)

        # Rotation row
        rot = tk.Frame(self.enh_win, bg=C["bg_mid"])
        rot.pack(pady=6)
        tk.Label(rot, text="Rotate Image:",
                 bg=C["bg_mid"], fg=C["text_main"],
                 font=("Consolas", 10)).pack(side="left", padx=8)
        for txt, angle in [("↺  −90°", -90), ("↻  +90°", 90)]:
            tk.Button(rot, text=txt,
                      command=lambda a=angle: self._rotate_image(a),
                      bg=C["bg_card"], fg=C["accent5"],
                      font=("Consolas", 10, "bold"),
                      relief="flat", bd=0, padx=14, pady=6,
                      cursor="hand2").pack(side="left", padx=8)

        tk.Button(self.enh_win, text="✅  Apply & Use This Image",
                  command=self._apply_enhancements,
                  bg=C["accent2"], fg=C["bg_dark"],
                  font=("Courier New", 12, "bold"),
                  relief="flat", bd=0, pady=10,
                  cursor="hand2").pack(pady=16, padx=30, fill="x")

        self._update_enh_preview()

    def _get_enhanced_preview(self):
        img  = cv2.convertScaleAbs(self.current_image,
                                   alpha=self.val_contrast.get(),
                                   beta=self.val_bright.get())
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY) \
               if len(img.shape) == 3 else img
        tv = self.val_thresh.get()
        if tv > 0:
            _, gray = cv2.threshold(gray, tv, 255, cv2.THRESH_BINARY)
        return gray

    def _update_enh_preview(self, *_):
        preview = self._get_enhanced_preview()
        self._enh_img_cache = preview
        rgb = cv2.cvtColor(preview, cv2.COLOR_GRAY2RGB)
        pil = Image.fromarray(rgb)
        pil.thumbnail((580, 360), Image.LANCZOS)
        self._enh_tk = ImageTk.PhotoImage(pil)
        self.enh_canvas.delete("all")
        self.enh_canvas.create_image(290, 180, image=self._enh_tk)

    def _rotate_image(self, angle):
        self.rotation_angle = (self.rotation_angle + angle) % 360
        h, w = self.current_image.shape[:2]
        M = cv2.getRotationMatrix2D((w // 2, h // 2), -angle, 1.0)
        self.current_image = cv2.warpAffine(self.current_image, M, (w, h))
        self._update_enh_preview()

    def _apply_enhancements(self):
        if self._enh_img_cache is None:
            self._enh_img_cache = self._get_enhanced_preview()
        self.processed_image = self._enh_img_cache
        self._show_image(self.processed_image)
        self._set_status("  ◉  Enhancements applied. Ready for OCR.", C["accent2"])
        self.enh_win.destroy()

    # ─────────────────── OCR PREPROCESSING ───────────────────────────
    def _preprocess_for_ocr(self, img):
        h, w = img.shape[:2]
        if w < 1000:
            img = cv2.resize(img, None, fx=1000/w, fy=1000/w,
                             interpolation=cv2.INTER_CUBIC)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY) \
               if len(img.shape) == 3 else img
        gray = cv2.GaussianBlur(gray, (5, 5), 0)
        gray = cv2.fastNlMeansDenoising(gray, h=10)
        _, gray = cv2.threshold(gray, 0, 255,
                                cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        return gray

    # ─────────────────── EXTRACT & SAVE  (threaded) ───────────────────
    def process_and_save(self):
        if self.current_image is None:
            messagebox.showwarning("No Image",
                "Please capture or upload an invoice first.")
            return
        threading.Thread(target=self._ocr_worker, daemon=True).start()

    def _ocr_worker(self):
        self._set_status("  ◎  Preprocessing image…", C["accent4"])
        self._set_progress(15)
        try:
            # Save original colour image
            ts             = datetime.now().strftime("%Y%m%d_%H%M%S")
            image_filename = os.path.join(SAVE_DIR, f"invoice_{ts}.jpg")
            cv2.imwrite(image_filename, self.current_image)

            img_to_ocr = self.processed_image \
                         if self.processed_image is not None \
                         else self._preprocess_for_ocr(self.current_image)

            self._set_status("  ◎  Running Tesseract OCR…", C["accent4"])
            self._set_progress(45)

            text = pytesseract.image_to_string(
                img_to_ocr, config=r'--oem 3 --psm 6')
            self.last_ocr_text = text

            self._set_progress(70)

            if not text.strip():
                self.root.after(0, lambda: messagebox.showerror(
                    "OCR Failed",
                    "No text detected in image.\n\n"
                    "Tips:\n"
                    "  • Make sure Tesseract is installed\n"
                    "  • Try enhancing the image first (increase contrast)\n"
                    "  • Ensure image is not blurry or too dark"))
                self._set_status("  ✗  OCR failed — no text found.", C["accent3"])
                self._set_progress(0)
                return

            # ── SMART DATE EXTRACTION ──────────────────────────────
            self._set_status("  ◎  Extracting date & total…", C["accent4"])
            self._set_progress(85)

            invoice_date = extract_invoice_date(text)   # uses the smart extractor

            # ── TOTAL EXTRACTION ───────────────────────────────────
            total_match = re.search(
                r'(?:grand\s+total|net\s+total|total\s+amount|amount\s+due|total)'
                r'[^\d\n]*([\d,]+(?:\.\d{1,2})?)',
                text, re.IGNORECASE)
            total_amount = total_match.group(1) if total_match else "Not Found"

            # ── SAVE TO DATABASE ───────────────────────────────────
            conn = sqlite3.connect(DB_PATH)
            try:
                conn.execute(
                    "INSERT INTO invoices "
                    "(invoice_date, total_amount, raw_text, image_path) "
                    "VALUES (?, ?, ?, ?)",
                    (invoice_date, total_amount, text, image_filename))
                conn.commit()
            finally:
                conn.close()

            self._set_progress(100)
            self._set_status(
                f"  ✓  Saved  ·  Date: {invoice_date}  ·  Total: {total_amount}",
                C["accent2"])
            self.root.after(0, self._refresh_stats)

            def _show_result():
                ans = messagebox.askyesno(
                    "Extraction Complete ✓",
                    f"Date Found   :  {invoice_date}\n"
                    f"Total Found  :  {total_amount}\n\n"
                    f"Image saved to:\n{image_filename}\n\n"
                    "Copy full OCR text to clipboard?")
                if ans:
                    self.copy_ocr_text()
            self.root.after(0, _show_result)
            self.root.after(3000, lambda: self._set_progress(0))

        except pytesseract.TesseractNotFoundError:
            self.root.after(0, lambda: messagebox.showerror(
                "Tesseract Not Found",
                "Tesseract OCR engine is not installed or path is wrong.\n\n"
                "Install from:\nhttps://github.com/UB-Mannheim/tesseract/wiki\n\n"
                f"Current path:\n{pytesseract.pytesseract.tesseract_cmd}\n\n"
                "Click ⚙ in the header to update the path."))
            self._set_status("  ✗  Tesseract not found.", C["accent3"])
            self._set_progress(0)
        except Exception as exc:
            self.root.after(0,
                lambda e=exc: messagebox.showerror("Error", str(e)))
            self._set_status(f"  ✗  Error: {exc}", C["accent3"])
            self._set_progress(0)

    # ─────────────── COPY OCR TEXT TO CLIPBOARD ───────────────────────
    def copy_ocr_text(self):
        if not self.last_ocr_text.strip():
            messagebox.showinfo("No Text",
                "No OCR text yet. Run Extract & Save first.")
            return
        try:
            pyperclip.copy(self.last_ocr_text)
        except Exception:
            self.root.clipboard_clear()
            self.root.clipboard_append(self.last_ocr_text)
        self._set_status("  ✓  OCR text copied to clipboard.", C["accent2"])

    # ─────────────────── DATABASE VIEWER ──────────────────────────────
    def view_database(self):
        win = tk.Toplevel(self.root)
        win.title("Styluxe — Financial Records")
        win.geometry("980x550")
        win.configure(bg=C["bg_mid"])

        # Toolbar
        tb = tk.Frame(win, bg=C["bg_mid"])
        tb.pack(fill="x", padx=10, pady=8)
        tk.Label(tb, text="🔍  Search:",
                 bg=C["bg_mid"], fg=C["text_muted"],
                 font=("Consolas", 10)).pack(side="left")
        search_var = tk.StringVar()
        tk.Entry(tb, textvariable=search_var,
                 bg=C["bg_card"], fg=C["text_main"],
                 insertbackground=C["accent1"],
                 font=("Consolas", 10), relief="flat", width=32).pack(
                     side="left", padx=8, ipady=4)

        tk.Label(win,
                 text="💡  Double-click any row to view raw OCR text & open image.",
                 bg=C["bg_mid"], fg=C["accent5"],
                 font=("Consolas", 9, "italic")).pack(anchor="w", padx=12)

        # Treeview
        cols   = ("ID", "Scan Time", "Invoice Date", "Total", "Image File")
        widths = (50, 160, 120, 110, 340)
        tree   = ttk.Treeview(win, columns=cols, show="headings",
                              selectmode="browse")
        for col, w in zip(cols, widths):
            tree.heading(col, text=col)
            tree.column(col, width=w,
                        anchor="w" if col == "Image File" else "center")

        vsb = ttk.Scrollbar(win, orient="vertical", command=tree.yview)
        tree.configure(yscroll=vsb.set)
        vsb.pack(side="right", fill="y", padx=(0, 6))
        tree.pack(fill="both", expand=True, padx=10, pady=4)

        def load_records(filter_text=""):
            tree.delete(*tree.get_children())
            conn = sqlite3.connect(DB_PATH)
            try:
                cur = conn.cursor()
                if filter_text:
                    q = f"%{filter_text}%"
                    cur.execute(
                        "SELECT id, scan_date, invoice_date, total_amount, image_path "
                        "FROM invoices "
                        "WHERE invoice_date LIKE ? OR total_amount LIKE ? OR image_path LIKE ? "
                        "ORDER BY id DESC", (q, q, q))
                else:
                    cur.execute(
                        "SELECT id, scan_date, invoice_date, total_amount, image_path "
                        "FROM invoices ORDER BY id DESC")
                for row in cur.fetchall():
                    tree.insert("", "end", values=row)
            finally:
                conn.close()

        load_records()
        search_var.trace_add("write", lambda *_: load_records(search_var.get()))

        # Double-click → detail window
        def on_double_click(_):
            sel = tree.selection()
            if not sel:
                return
            row_id = tree.item(sel[0], "values")[0]
            conn = sqlite3.connect(DB_PATH)
            try:
                cur = conn.cursor()
                cur.execute(
                    "SELECT raw_text, image_path FROM invoices WHERE id=?",
                    (row_id,))
                result = cur.fetchone()
            finally:
                conn.close()
            if not result:
                return
            raw_text, img_path = result

            dw = tk.Toplevel(win)
            dw.title(f"Record #{row_id}  ·  Raw OCR Text")
            dw.geometry("640x520")
            dw.configure(bg=C["bg_mid"])
            tk.Label(dw, text=f"  Record #{row_id}  ·  Extracted Date:  "
                     f"{tree.item(sel[0], 'values')[2]}",
                     font=("Consolas", 11, "bold"),
                     bg=C["bg_mid"], fg=C["accent1"]).pack(
                         anchor="w", padx=12, pady=10)
            txt = scrolledtext.ScrolledText(
                dw, font=("Consolas", 10),
                bg=C["bg_card"], fg=C["text_main"],
                insertbackground=C["accent1"],
                relief="flat", wrap="word")
            txt.insert("1.0", raw_text or "(No text stored)")
            txt.config(state="disabled")
            txt.pack(fill="both", expand=True, padx=12, pady=4)

            br = tk.Frame(dw, bg=C["bg_mid"])
            br.pack(fill="x", padx=12, pady=10)

            def _copy():
                dw.clipboard_clear()
                dw.clipboard_append(raw_text or "")
                messagebox.showinfo("Copied", "Text copied!", parent=dw)

            def _open_img():
                if img_path and os.path.exists(img_path):
                    Image.open(img_path).show()
                else:
                    messagebox.showerror("Not Found",
                        "Image file not found on disk.", parent=dw)

            for txt_b, cmd_b, col_b in [
                ("📋  Copy Text",   _copy,    C["accent1"]),
                ("🖼   Open Image",  _open_img, C["accent4"]),
            ]:
                tk.Button(br, text=txt_b, command=cmd_b,
                          bg=C["bg_card"], fg=col_b,
                          font=("Consolas", 10), relief="flat",
                          padx=12, pady=6).pack(side="left", padx=4)

        tree.bind("<Double-1>", on_double_click)

        # Delete selected
        def delete_selected():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("No Selection",
                    "Select a record first.", parent=win)
                return
            row_id = tree.item(sel[0], "values")[0]
            if not messagebox.askyesno(
                    "Confirm Delete",
                    f"Delete record #{row_id}? This cannot be undone.",
                    parent=win):
                return
            conn = sqlite3.connect(DB_PATH)
            try:
                conn.execute("DELETE FROM invoices WHERE id=?", (row_id,))
                conn.commit()
            finally:
                conn.close()
            load_records(search_var.get())
            self.root.after(0, self._refresh_stats)
            self._set_status(f"  ✓  Record #{row_id} deleted.", C["accent5"])

        # Export CSV
        def export_csv():
            fp = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")],
                parent=win)
            if not fp:
                return
            conn = sqlite3.connect(DB_PATH)
            try:
                cur = conn.cursor()
                cur.execute(
                    "SELECT id, scan_date, invoice_date, total_amount, image_path "
                    "FROM invoices ORDER BY id DESC")
                rows = cur.fetchall()
            finally:
                conn.close()
            with open(fp, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["ID", "Scan Time", "Invoice Date",
                             "Total Amount", "Image Path"])
                w.writerows(rows)
            messagebox.showinfo("Exported",
                f"✓  {len(rows)} records exported to:\n{fp}", parent=win)

        # Bottom buttons
        bot = tk.Frame(win, bg=C["bg_mid"])
        bot.pack(fill="x", padx=10, pady=8)
        for txt_b, cmd_b, col_b in [
            ("📥  Export to CSV",   export_csv,      C["accent1"]),
            ("🗑   Delete Selected", delete_selected, C["accent3"]),
        ]:
            tk.Button(bot, text=txt_b, command=cmd_b,
                      bg=C["bg_card"], fg=col_b,
                      font=("Consolas", 10, "bold"),
                      relief="flat", padx=14, pady=7,
                      cursor="hand2").pack(side="left", padx=6)

    # ─────────────────────── SETTINGS ─────────────────────────────────
    def open_settings(self):
        win = tk.Toplevel(self.root)
        win.title("Settings")
        win.geometry("560x240")
        win.configure(bg=C["bg_mid"])
        win.grab_set()

        tk.Label(win, text="  ⚙  Settings",
                 font=("Courier New", 13, "bold"),
                 bg=C["bg_mid"], fg=C["accent4"]).pack(
                     anchor="w", padx=16, pady=14)

        row = tk.Frame(win, bg=C["bg_mid"])
        row.pack(fill="x", padx=16, pady=4)
        tk.Label(row, text="Tesseract Path:",
                 bg=C["bg_mid"], fg=C["text_main"],
                 font=("Consolas", 10), width=16, anchor="w").pack(side="left")
        path_var = tk.StringVar(value=pytesseract.pytesseract.tesseract_cmd)
        tk.Entry(row, textvariable=path_var,
                 bg=C["bg_card"], fg=C["text_main"],
                 insertbackground=C["accent1"],
                 font=("Consolas", 10), relief="flat", width=36).pack(
                     side="left", padx=6, ipady=4)

        def _browse():
            p = filedialog.askopenfilename(
                filetypes=[("Executables", "*.exe"), ("All", "*.*")],
                parent=win)
            if p:
                path_var.set(p)

        tk.Button(row, text="Browse", command=_browse,
                  bg=C["bg_card"], fg=C["accent1"],
                  font=("Consolas", 9), relief="flat",
                  padx=8, pady=4).pack(side="left")

        def _save():
            pytesseract.pytesseract.tesseract_cmd = path_var.get().strip()
            self._save_config()
            self._set_status("  ✓  Settings saved.", C["accent2"])
            win.destroy()

        tk.Button(win, text="💾  Save Settings", command=_save,
                  bg=C["accent2"], fg=C["bg_dark"],
                  font=("Courier New", 11, "bold"),
                  relief="flat", pady=8, cursor="hand2").pack(
                      padx=16, pady=20, fill="x")

    # ─────────────────────── HELPERS ──────────────────────────────────
    def _set_status(self, msg, color=C["text_muted"]):
        self.root.after(0, lambda: self.status_var.set(msg))

    def _set_progress(self, val):
        self.root.after(0, lambda: self.progress_var.set(val))


# ═══════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    root = tk.Tk()
    try:
        root.iconbitmap(os.path.join(BASE_DIR, "icon.ico"))
    except Exception:
        pass
    app = InvoiceApp(root)
    root.mainloop()