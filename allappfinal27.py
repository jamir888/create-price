# path: create_price_labels.py
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Create Price Labels — parity macOS/Windows
Fixes:
- Single Theme block (no duplicates)
- Qt stylesheet uses explicit font props (no CSS shorthands)
- HiDPI flags for Windows
- Remove unreachable duplicate manual block
- Guard double commit in inline editors
- Excel popup hides on focus-out too

Extra stability (Windows bundle parity):
- Template tiles use Fixed size policies (no layout over-expansion)
- Prevent templates area from pushing Generate/Table out of view
- Renderer: case-insensitive align + per-side clipping (no cross-midline bleed)

Performance + Quick Import fixes:
- Fast Excel import via pandas (no heavy COM roundtrips)
- Quick Import fallback: if DB has no rows for a saved source, import from its file path
"""
import os, sys, re, hashlib, tempfile, subprocess, pathlib, glob, json, base64
import csv
import pathlib
import stat 
# move this import to the top of your file, BEFORE _first_run_seed() is called
import shutil
import pandas as pd
pd.set_option('future.no_silent_downcasting', True)

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module=r"openpyxl(\.|$)")


from datetime import datetime, date
from typing import Dict, List, Optional, Tuple, Callable
from PySide6 import QtWidgets, QtGui, QtCore
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFrame, QLabel, QLineEdit, QPushButton, QCheckBox,
    QTextEdit, QVBoxLayout, QHBoxLayout, QGridLayout, QDialog, QInputDialog, QMessageBox,
    QTableWidget, QTableWidgetItem, QMenu, QAbstractItemView, QHeaderView, QListWidget,
    QListWidgetItem, QSizePolicy, QToolButton, QScrollArea,  QFileDialog, 
)
from PySide6.QtCore import Qt, QEvent, QTimer, QRect, QPoint, QPropertyAnimation, QEasingCurve, QSize
from PySide6.QtGui import (
    QFont, QColor, QPalette, QCursor, QKeyEvent, QMouseEvent, QResizeEvent, QFocusEvent, QClipboard,
    QGuiApplication, QIcon, QPixmap
)
from PySide6.QtWidgets import QGraphicsDropShadowEffect
from PySide6.QtWidgets import QWidget, QHBoxLayout, QSizePolicy
from PySide6.QtWidgets import QLayout, QStyle
from PySide6.QtGui import QGuiApplication
import sys, os, shutil
from pathlib import Path
from shiboken6 import Shiboken
# 1) ADD import (near your other QtCore imports)
from PySide6.QtCore import QObject, QEvent
from PySide6.QtCore import Signal, QThread
from PySide6.QtCore import QObject, QTimer
from pathlib import Path
import os, stat, time
# --- add once inside _ensure_reportlab() ---
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pathlib import Path
from pathlib import Path
import sys, platform
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import platform
from pathlib import Path
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont




_ARABIC_FONT_NAME = "SysArabic"  # keep consistent everywhere

# --- Arabic font resolver (bundled-first, Windows-friendly) ---

def _find_system_arabic_font() -> tuple[str | None, int | None]:
    """
    Returns (path, subfontIndex) for a good Arabic-capable font.
    Order of preference:
      1) Bundled fonts in ./fonts/
      2) Windows fonts (Dubai / Traditional Arabic / Tahoma / Arial)
      3) macOS system Arabic fonts (Geeza, Damascus, AlNile, Baghdad)
      4) Linux common Noto Arabic
    """
    base = os.path.abspath(os.path.dirname(__file__))
    bundled = [
        os.path.join(base, "fonts", "Dubai-Regular.ttf"),
        os.path.join(base, "fonts", "Amiri-Regular.ttf"),
        os.path.join(base, "fonts", "NotoNaskhArabic-Regular.ttf"),
        os.path.join(base, "fonts", "NotoKufiArabic-Regular.ttf"),
    ]
    for p in bundled:
        try:
            if Path(p).exists():
                return (p, None)
        except Exception:
            pass

    sysname = platform.system()
    if sysname == "Windows":
        # Prefer Dubai on Win10/11, then Traditional Arabic, then Tahoma/Arial
        candidates = [
            r"C:\Windows\Fonts\Dubai-Regular.ttf",
            r"C:\Windows\Fonts\trado.ttf",              # Traditional Arabic
            r"C:\Windows\Fonts\arabtype.ttf",           # Arabic Typesetting (old)
            r"C:\Windows\Fonts\tahoma.ttf",
            r"C:\Windows\Fonts\arial.ttf",
        ]
    elif sysname == "Darwin":  # macOS
        candidates = [
            "/System/Library/Fonts/Supplemental/GeezaPro.ttc",
            "/System/Library/Fonts/Supplemental/Damascus.ttc",
            "/System/Library/Fonts/Supplemental/AlNile.ttc",
            "/System/Library/Fonts/Supplemental/Baghdad.ttc",
            "/Library/Fonts/GeezaPro.ttc",
        ]
    else:  # Linux
        candidates = [
            "/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf",
            "/usr/share/fonts/truetype/noto/NotoSansArabic-Regular.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ]

    for p in candidates:
        try:
            if Path(p).exists():
                if p.lower().endswith(".ttc"):
                    return (p, 0)  # first face
                return (p, None)
        except Exception:
            continue
    return (None, None)


def _ensure_arabic_font() -> str:
    """
    Register the Arabic font with ReportLab (TTF is embedded).
    Returns the face name to use with canvas.setFont(...).
    Silent if already registered.
    """
    try:
        # Already registered?
        pdfmetrics.getFont(_ARABIC_FONT_NAME)
        return _ARABIC_FONT_NAME
    except Exception:
        pass

    path, sub_idx = _find_system_arabic_font()
    if not path:
        # No font found; caller can still fall back to system family name
        return _ARABIC_FONT_NAME

    try:
        if path.lower().endswith(".ttc"):
            pdfmetrics.registerFont(TTFont(_ARABIC_FONT_NAME, path, subfontIndex=(sub_idx or 0)))
        else:
            pdfmetrics.registerFont(TTFont(_ARABIC_FONT_NAME, path))
        # Registered successfully; ReportLab will embed this TTF in PDFs
        return _ARABIC_FONT_NAME
    except Exception:
        # On failure, just return the logical name; Windows will still render with system fonts
        return _ARABIC_FONT_NAME


def _is_alive(obj):
    return obj is not None and Shiboken.isValid(obj)

# --- ADD near the top, right after _is_alive --------------------------------
def alive(obj) -> bool:
    """Safe validity check for Qt wrappers."""
    return obj is not None and Shiboken.isValid(obj)

def _safe_widget(owner, name: str):
    """Return a live widget attribute or None."""
    w = getattr(owner, name, None)
    return w if alive(w) else None    

def _stop_thread(t):
    try:
        if t and t.isRunning():
            t.requestInterruption()
            t.quit()
            t.wait(2000)
    except Exception:
        pass


def resource_path(rel: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
    return (base / rel).resolve()

# ---------- Core headers + app data (single source of truth) ----------
CORE_HEADERS = ["BARCODE","BRAND","ITEM","REG","PROMO","START_DATE","END_DATE","SECTION","COOP"]

def _db_dir()->pathlib.Path:
    p = pathlib.Path.home() / "Documents" / "PriceLabelsDB"
    p.mkdir(parents=True, exist_ok=True)
    return p

def _db_path()->str: return str(_db_dir() / "labels_db.csv")
def _settings_path()->str: return str(_db_dir() / "app_settings.json")
def _headers_cfg_path()->str: return str(_db_dir() / "headers_config.json")
def _excel_sources_path()->str: return str(_db_dir() / "excel_sources.json")

TEMPLATES_DIR = str(_db_dir() / "templates")
os.makedirs(TEMPLATES_DIR, exist_ok=True)

def _first_run_seed():
    src_tpl = resource_path("templates")
    dest_tpl = pathlib.Path(TEMPLATES_DIR)
    try:
        if src_tpl.exists() and not any(dest_tpl.iterdir()):
            shutil.copytree(src_tpl, dest_tpl, dirs_exist_ok=True)
    except Exception:
        pass

    src_hdr = resource_path("headers.json")
    dest_hdr = pathlib.Path(_headers_cfg_path())
    try:
        if not dest_hdr.exists():
            if src_hdr.exists():
                shutil.copy(src_hdr, dest_hdr)
            else:
                dest_hdr.write_text("{}", encoding="utf-8")
    except Exception:
        pass

_first_run_seed()



# ---------- Lazy heavy deps ----------
def _pd():
    import importlib
    return importlib.import_module("pandas")

def _ensure_reportlab():
    global pdfgen_canvas, A4, mm, black, Color, pdfmetrics
    if "pdfgen_canvas" in globals():
        return
    from reportlab.pdfgen import canvas as pdfgen_canvas  # type: ignore
    from reportlab.lib.pagesizes import A4               # type: ignore
    from reportlab.lib.units import mm                   # type: ignore
    from reportlab.lib.colors import black, Color        # type: ignore
    from reportlab.pdfbase import pdfmetrics             # type: ignore




# ---------- Theme (single source of truth) ----------
CLR_BG            = "#F6F1EA"
CLR_TEXT          = "#262626"
CLR_CARD          = "#FCFAF7"
CLR_PRIMARY       = "#E9DFD4"
CLR_PRIMARY_HOVER = "#E4D6C7"
CLR_PRIMARY_ACTIVE= "#DFCAB5"
CLR_ACCENT        = "#3B82F6"
CLR_BORDER        = "#D8D2CA"
CLR_MUTED_TEXT    = "#6B6B6B"
CLR_INPUT_BG      = "#FFFFFF"
CLR_INPUT_BORDER  = "#D7D1C9"
CLR_INPUT_FOCUS   = "#BDAF9E"

CLR_GEN_OFF       = "#E9DFD4"
CLR_GEN_ON        = "#D7C7B5"


APP_TITLE = "Create Price Labels"
SMALL_GEOM = (620, 460, 200, 140)
LARGE_GEOM = (1200, 820, 160, 80)

# === Fresh Section (toggleable; default OFF) ===
FRESH_SECTION_ACTIVE = False  # runtime-only (not persisted yet)
FRESH_HEADERS = (
    "PLU",
    "ARABIC_DESCRIPTION",
    "ENGLISH_DESCRIPTION",
    "REGULAR_PRICE",
    "PROMO_PRICE",
)

BRAND_PREFIX_MIN_COMPARABLES = 20   # how many rows must be comparable
BRAND_PREFIX_MIN_RATIO = 0.10       # 10% exact-match to first 1–2 words         
# === [PDF FONT/ARABIC HELPERS — unified] ===

def _ensure_unicode_fonts():
    """Register Noto + Arabic fonts and prepare Arabic shaping/BiDi."""
    global _fonts_ready, _arabic_reshape, _bidi_get_display
    if globals().get("_fonts_ready", False):
        return
    _ensure_reportlab()

    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from pathlib import Path

    def _font_path(fname: str) -> Path:
        for base in [Path(resource_path("fonts")), Path(__file__).parent / "fonts"]:
            p = base / fname
            if p.exists():
                return p
        return Path(fname)

    def _try_register(name: str, fname: str) -> bool:
        try:
            p = _font_path(fname)
            if not p.exists():
                return False
            pdfmetrics.registerFont(TTFont(name, str(p)))
            return True
        except Exception:
            return False

    latin_ok_r = _try_register("NotoSans-Regular", "NotoSans-Regular.ttf")
    latin_ok_b = _try_register("NotoSans-Bold", "NotoSans-Bold.ttf")
    arab_ok_r  = _try_register("NotoNaskhArabic-Regular", "NotoNaskhArabic-Regular.ttf")
    arab_ok_b  = _try_register("NotoNaskhArabic-Bold", "NotoNaskhArabic-Bold.ttf")

    try:
        _rn = set(pdfmetrics.getRegisteredFontNames())
        if "NotoSans-Regular" in _rn:
            pdfmetrics.registerFontFamily(
                "NotoSans",
                normal="NotoSans-Regular",
                bold="NotoSans-Bold" if "NotoSans-Bold" in _rn else "NotoSans-Regular",
            )
        if "NotoNaskhArabic-Regular" in _rn:
            pdfmetrics.registerFontFamily(
                "NotoNaskhArabic",
                normal="NotoNaskhArabic-Regular",
                bold="NotoNaskhArabic-Bold" if "NotoNaskhArabic-Bold" in _rn else "NotoNaskhArabic-Regular",
            )
    except Exception:
        pass

    if not (arab_ok_r or arab_ok_b):
        _ensure_arabic_font()

    try:
        import arabic_reshaper as _ar
        _arabic_reshape = _ar.ArabicReshaper(_ar.config_for_true_type())
    except Exception:
        _arabic_reshape = None
    try:
        from bidi.algorithm import get_display as _gd
        _bidi_get_display = _gd
    except Exception:
        _bidi_get_display = None

    globals()["_fonts_ready"] = True



_ARABIC_RANGES = [
    (0x0600, 0x06FF),  # Arabic
    (0x0750, 0x077F),  # Arabic Supplement
    (0x08A0, 0x08FF),  # Arabic Extended-A
    (0xFB50, 0xFDFF),  # Arabic Presentation Forms-A
    (0xFE70, 0xFEFF),  # Arabic Presentation Forms-B
]

def _contains_arabic(text: str) -> bool:
    if not text:
        return False
    for ch in text:
        cp = ord(ch)
        for a, b in _ARABIC_RANGES:
            if a <= cp <= b:
                return True
    return False

def _sanitize_text(text: str) -> str:
    """Normalize user text before shaping. Why: avoid PDF gaps, strip control chars, and drop NaN."""
    import math

    # Normalize None / NaN (pandas, Excel) to empty
    if text is None:
        return ""
    try:
        # handles float('nan') and numpy.nan
        if isinstance(text, float) and math.isnan(text):
            return ""
    except Exception:
        pass

    s = str(text).strip()

    # Common textual NaN markers -> empty
    if s.lower() in {"nan", "none", "null", "na", "n/a", "-"}:
        return ""

    # strip zero-width & bidi control characters that break shaping/width
    zw = {
        "\u200b", "\u200c", "\u200d", "\ufeff", "\u202a", "\u202b", "\u202c",
        "\u202d", "\u202e", "\u2066", "\u2067", "\u2068", "\u2069", "\u200e", "\u200f"
    }
    if any(ch in s for ch in zw):
        for ch in zw:
            s = s.replace(ch, "")

    # collapse whitespace for stable width measurement
    s = " ".join(s.split())
    return s


def _shape_for_pdf(text: str) -> str:
    """
    Returns a display-ready string after:
      1) sanitize (strip zero-width & normalize spaces)
      2) Arabic shaping (arabic_reshaper) if available
      3) BiDi reordering (python-bidi) if available
    Non-Arabic text is returned sanitized-as-is.
    """
    raw = "" if text is None else str(text)
    raw = _sanitize_text(raw)
    if not raw:
        return ""
    if not _contains_arabic(raw):
        return raw
    try:
        s = _arabic_reshape.reshape(raw) if _arabic_reshape else raw
        s = _bidi_get_display(s) if _bidi_get_display else s
        return s
    except Exception:
        return raw
# === [end helpers] ===



def apply_styles(app):
    app.setStyle("Fusion")
    pal = QPalette()
    pal.setColor(QPalette.Window,       QColor(CLR_BG))
    pal.setColor(QPalette.Base,         QColor(CLR_INPUT_BG))
    pal.setColor(QPalette.AlternateBase,QColor(CLR_CARD))
    pal.setColor(QPalette.Button,       QColor(CLR_PRIMARY))
    pal.setColor(QPalette.ButtonText,   QColor(CLR_TEXT))
    pal.setColor(QPalette.Text,         QColor(CLR_TEXT))
    pal.setColor(QPalette.WindowText,   QColor(CLR_TEXT))
    pal.setColor(QPalette.ToolTipBase,  QColor("#111111"))
    pal.setColor(QPalette.ToolTipText,  QColor("#FFFFFF"))
    pal.setColor(QPalette.Highlight,    QColor(CLR_PRIMARY_ACTIVE))
    pal.setColor(QPalette.Disabled, QPalette.Text, QColor("#9E9E9E"))
    app.setPalette(pal)

    base_font = "Arial"

    stylesheet = f"""
    * {{
        font-family: {base_font};
        color: {CLR_TEXT};
    }}
    QFrame[objectName="Card"] {{
        background-color: {CLR_CARD};
        border: 1px solid {CLR_BORDER};
        border-radius: 10px;
    }}
    QLabel[objectName="Title"] {{
        font-weight: 700; font-size: 22px; font-family: {base_font};
        letter-spacing:.3px; padding: 6px 8px;
    }}
    QLabel[objectName="Back"]  {{
        font-weight: 600; font-size: 13px; font-family: {base_font};
        color:{CLR_MUTED_TEXT}; padding: 4px 8px;
    }}
    QLabel[objectName="Small"] {{
        font-weight: 500; font-size: 12px; font-family: {base_font};
        color:{CLR_MUTED_TEXT}; padding: 2px 2px;
    }}
    QLabel[objectName="FormLabel"] {{
        font-weight: 600; font-size: 12px; font-family: {base_font};
        color:{CLR_TEXT}; padding: 0 2px 2px 2px; letter-spacing:.2px;
    }}
    QLineEdit {{
        background:{CLR_INPUT_BG}; border:1px solid {CLR_INPUT_BORDER}; border-radius:10px; padding:8px 10px;
        selection-background-color:{CLR_PRIMARY_ACTIVE};
    }}
    QLineEdit:focus {{ border:1px solid {CLR_INPUT_FOCUS}; outline:none; }}
    QLineEdit:disabled {{ background:#F2F2F2; color:#9A9A9A; }}
    QTextEdit {{
        background:{CLR_INPUT_BG}; border:1px solid {CLR_INPUT_BORDER}; border-radius:10px; padding:8px;
    }}
    QTextEdit:focus {{ border:1px solid {CLR_INPUT_FOCUS}; }}
    QCheckBox, QRadioButton {{ spacing:6px; }}
    QPushButton {{
        border:1px solid {CLR_BORDER}; border-radius:12px; padding:2px 6px; background-color:{CLR_PRIMARY};
        font-weight:600; font-size:12px; font-family:{base_font}; qproperty-iconSize:16px;
    }}
    QPushButton:hover  {{ background-color:{CLR_PRIMARY_HOVER}; }}
    QPushButton:pressed{{ background-color:{CLR_PRIMARY_ACTIVE}; }}
    QPushButton:disabled {{ background-color:#EEE8E1; color:#9A8F83; border-color:#E0D8CF; }}
    QPushButton::menu-indicator {{ image:none; }}
    QPushButton[objectName="Primary"],
    QPushButton[objectName="GenerateOn"],
    QPushButton[objectName="GenerateOff"],
    QPushButton[objectName="Template"],
    QPushButton#SmartAI {{ padding:8px 14px; }}
    QTableWidget {{
        background:{CLR_INPUT_BG}; gridline-color:{CLR_BORDER};
        border:1px solid {CLR_BORDER}; border-radius:10px; font-size:12px; font-family:{base_font};
    }}
    QTableWidget::item {{ padding:3px; }}
    QTableWidget::item:selected {{ background:{CLR_PRIMARY_ACTIVE}; color:{CLR_TEXT}; }}
    QHeaderView::section {{
        background:#EEE7DF; color:{CLR_TEXT}; border:0; border-right:1px solid {CLR_BORDER};
        padding:6px 6px; font-weight:600; font-size:12px; font-family:{base_font};
    }}
    QHeaderView::section:horizontal {{ border-bottom:1px solid {CLR_BORDER}; }}
    QTableCornerButton::section {{ background:#EEE7DF; border:none; }}
    QTableView::item:hover {{ background:#F6EFE7; }}
        QTableWidget#StageTable {{ font-size:12px; font-family:{base_font}; }}
    QTableWidget#StageTable QHeaderView::section {{
        padding:6px 6px; font-weight:600; font-size:12px; font-family:{base_font};
    }}
    QTableWidget#ExcelTable {{ font-size:12px; font-family:{base_font}; }}
    QTableWidget#ExcelTable QHeaderView::section {{
        padding:6px 6px; font-weight:600; font-size:12px; font-family:{base_font};
    }}

    QTableWidget#LiveHits {{ font-size:12px; font-family:{base_font}; }}
    QTableWidget#LiveHits QHeaderView::section {{ padding:2px 6px; font-weight:600; font-size:12px; font-family:{base_font}; }}
    QListWidget {{
        background:{CLR_INPUT_BG}; border:1px solid {CLR_INPUT_BORDER}; border-radius:8px;
    }}
    QListWidget::item {{ padding:6px 8px; }}
    QListWidget::item:selected {{ background:{CLR_PRIMARY_ACTIVE}; color:{CLR_TEXT}; }}
    QScrollBar:vertical {{ background:transparent; width:10px; margin:4px 0; }}
    QScrollBar::handle:vertical {{ background:#D5CDC4; border-radius:6px; min-height:24px; }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height:0; background:transparent; border:none; }}
    QScrollBar:horizontal {{ background:transparent; height:10px; margin:0 4px; }}
    QScrollBar::handle:horizontal {{ background:#D5CDC4; border-radius:6px; min-width:24px; }}
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width:0; background:transparent; border:none; }}
    QToolTip {{
        background-color:rgba(34,34,34,0.96); color:#FFFFFF; border:1px solid #161616; padding:6px 8px;
        border-radius:6px; font-weight:600; font-size:11px; font-family:{base_font};
    }}
     QPushButton#FreshToggle {{
         border: 0; border-radius: 10px; padding: 8px 12px;
         font-weight: 700; font-size: 12px; font-family: {base_font};
         background-color: #16A34A; /* green-600 */
         color: #FFFFFF;
     }}
     QPushButton#FreshToggle:hover {{ background-color: #22C55E; }}  /* green-500 */
     QPushButton#FreshToggle:pressed {{ background-color: #15803D; }} /* green-700 */
     QPushButton#FreshToggle[active="false"] {{
         background-color: #E2E8F0; color: #0F172A;  /* slate-100 / slate-900 */
         border: 1px solid {CLR_BORDER};
     }}


    """
    app.setStyleSheet(stylesheet)



# ---------- small IO helpers ----------
def _read_json(path: str):
    try:
        with open(path, "r", encoding="utf-8") as f: return json.load(f)
    except Exception: return None

def _write_json(path: str, data: dict):
    with open(path, "w", encoding="utf-8") as f: json.dump(data, f, indent=2)

# path: create_price_labels.py
# --- REPLACE the whole pasted block (from def _list_template_files to before "settings") with this ---

def _list_template_files():
    files = sorted(glob.glob(os.path.join(TEMPLATES_DIR, "*.json")))
    return [(os.path.splitext(os.path.basename(p))[0], p) for p in files]


# --- Recent Downloads helpers (Excel) ---
def _candidate_recent_roots() -> List[Path]:
    """Common save locations: Downloads, Desktop, Documents, OneDrive (personal/business)."""
    dirs: List[Path] = []
    seen: set[str] = set()

    def add(p: Path):
        try:
            if not p or not p.exists():
                return
            rp = str(p.resolve())
        except Exception:
            rp = str(p)
        if rp and rp not in seen:
            seen.add(rp)
            dirs.append(p)

    home = Path.home()
    userprofile = Path(os.environ.get("USERPROFILE", "")) if os.environ.get("USERPROFILE", "") else home

    # Core
    for base in (home, userprofile):
        add(base / "Downloads")
        add(base / "Desktop")
        add(base / "Documents")

    # OneDrive env + detected variants
    od_env = os.environ.get("OneDrive", "")
    if od_env:
        od = Path(od_env)
        add(od); add(od / "Downloads"); add(od / "Desktop"); add(od / "Documents")

    try:
        for p in home.glob("OneDrive*"):  # e.g. OneDrive - Company
            add(p); add(p / "Downloads"); add(p / "Desktop"); add(p / "Documents")
    except Exception:
        pass

    return dirs


def _recent_excels_anywhere(limit: int = 5, *, max_depth: int = 3, per_root_cap: int = 300) -> List[Path]:
    """
    Return most-recent Excel files across common roots.
    Depth-limited for speed; per-root cap to avoid huge walks; de-duped by resolved path.
    """
    exts = {".xlsx", ".xls", ".xlsb", ".xlsm"}
    ranked: List[Tuple[float, Path]] = []

    for root in _candidate_recent_roots():
        taken = 0
        for dirpath, dirnames, filenames in os.walk(root):
            # depth-limit
            try:
                rel = Path(dirpath).resolve().relative_to(root.resolve())
                depth = len(rel.parts)
            except Exception:
                depth = 0
            if depth >= max_depth:
                dirnames[:] = []
            for fn in filenames:
                if taken >= per_root_cap:
                    break
                if os.path.splitext(fn)[1].lower() in exts:
                    p = Path(dirpath) / fn
                    try:
                        st = p.stat()
                        if stat.S_ISREG(st.st_mode):
                            ranked.append((st.st_mtime, p))
                            taken += 1
                    except Exception:
                        continue
            if taken >= per_root_cap:
                break

    ranked.sort(key=lambda t: (t[0], t[1].name.lower()), reverse=True)

    out: List[Path] = []
    seen: set[str] = set()
    for _, p in ranked:
        try:
            rp = str(p.resolve())
        except Exception:
            rp = str(p)
        if rp in seen:
            continue
        seen.add(rp)
        out.append(p)
        if len(out) >= int(limit):
            break

    return out

    # === [CHUNK 2] Date-based Excel search (entire PC) + popup selection ===
    # Paste after: def _recent_excels_anywhere(...)
    # Requires: PySide6, QtCore/QtWidgets imports already present in your file.

    # ---- Date parsing helpers (DD/MM/YYYY priority; accept many separators; 2-digit years = 2000-2099) ----
    _DATE_SEPS_RE = re.compile(r"[.\-_/\\]+")
    _DIGITS_ONLY_RE = re.compile(r"^\d{6,8}$")

    def _parse_user_date(user_input: str) -> Optional[date]:
        """
        Parse flexible date input with DD/MM/YYYY default.
        Accepts: 10-12-2025, 10/12/2025, 10-12/25, 10/12\2025, 10122025, 011225, 1/12/25, etc.
        Two-digit years map to 2000–2099.
        """
        if not user_input:
            return None
        s = str(user_input).strip()
        s = _DATE_SEPS_RE.sub("/", s)

        def _mk_date(d: int, m: int, y: int) -> Optional[date]:
            if 0 <= y <= 99:
                y = 2000 + y
            try:
                return date(y, m, d)
            except Exception:
                return None

        # digits-only like DDMMYYYY or DDMMYY
        if _DIGITS_ONLY_RE.match(s):
            try:
                if len(s) == 8:   # DDMMYYYY
                    d, m, y = int(s[0:2]), int(s[2:4]), int(s[4:8])
                    return _mk_date(d, m, y)
                if len(s) == 6:   # DDMMYY
                    d, m, y = int(s[0:2]), int(s[2:4]), int(s[4:6])
                    return _mk_date(d, m, y)
            except Exception:
                return None

        # split by separators (prefer DD/MM/YYYY)
        parts = [p for p in s.split("/") if p != ""]
        if len(parts) in (2, 3):
            try:
                d = int(parts[0])
                m = int(parts[1])
                y = int(parts[2]) if len(parts) == 3 else datetime.now().year
                return _mk_date(d, m, y)
            except Exception:
                return None
        return None


    def _date_filename_tokens(d: date) -> List[str]:
        """
        Build tokens that might appear in filenames for this date.
        We generate multiple common patterns to maximize matches.
        """
        dd = f"{d.day:02d}"
        mm = f"{d.month:02d}"
        yyyy = f"{d.year:04d}"
        yy = f"{d.year % 100:02d}"
        return [
            f"{dd}-{mm}-{yyyy}",
            f"{dd}-{mm}-{yy}",
            f"{dd}_{mm}_{yyyy}",
            f"{dd}_{mm}_{yy}",
            f"{dd}.{mm}.{yyyy}",
            f"{dd}.{mm}.{yy}",
            f"{yyyy}-{mm}-{dd}",
            f"{yyyy}_{mm}_{dd}",
            f"{yyyy}.{mm}.{dd}",
            f"{yyyy}{mm}{dd}",
            f"{dd}{mm}{yyyy}",
            f"{dd}{mm}{yy}",
        ]


    def _excel_ext_ok(path: str) -> bool:
        e = os.path.splitext(path)[1].lower()
        return e in {".xlsx", ".xls", ".xlsm", ".xlsb"}


    def _all_roots() -> List[Path]:
        """
        Enumerate all available roots to search:
        - Windows: all mounted drive letters.
        - macOS/Linux: '/', plus home and common user dirs.
        """
        roots: List[Path] = []
        try:
            if os.name == "nt":
                import string
                from ctypes import windll
                bitmask = windll.kernel32.GetLogicalDrives()
                for i, ch in enumerate(string.ascii_uppercase):
                    if bitmask & (1 << i):
                        roots.append(Path(f"{ch}:/"))
            else:
                roots.append(Path("/"))
        except Exception:
            pass

        # Add common user places (duplicates deduped later)
        home = Path.home()
        for p in [home, home / "Downloads", home / "Desktop", home / "Documents"]:
            roots.append(p)

        # OneDrive variants if present
        od_env = os.environ.get("OneDrive", "")
        if od_env:
            roots.append(Path(od_env))
        try:
            for p in home.glob("OneDrive*"):
                roots.append(p)
        except Exception:
            pass

        # De-dup by resolved path
        out: List[Path] = []
        seen: set[str] = set()
        for p in roots:
            try:
                rp = str(p.resolve())
            except Exception:
                rp = str(p)
            if rp not in seen and p.exists():
                seen.add(rp)
                out.append(p)
        return out


    # ---- Scanner thread: walks entire PC, searches filenames that contain any of the date tokens ----
    class DateExcelScanner(QThread):
        progress = Signal(str)          # current directory
        found = Signal(str)             # path string
        finished_with = Signal(list)    # list[str]

        def __init__(self, tokens: List[str], parent: Optional[QObject] = None):
            super().__init__(parent)
            self.tokens = [t.lower() for t in tokens if t]
            self._stop = False
            self._hits: List[str] = []

        def stop(self):
            self._stop = True

        def run(self):
            try:
                for root in _all_roots():
                    if self._stop:
                        break
                    for dirpath, dirnames, filenames in os.walk(root, topdown=True):
                        if self._stop:
                            break
                        self.progress.emit(dirpath)
                        # Optional throttling to keep UI responsive
                        QtCore.QThread.msleep(1)
                        for fn in filenames:
                            if self._stop:
                                break
                            if not _excel_ext_ok(fn):
                                continue
                            lower_name = fn.lower()
                            if any(tok in lower_name for tok in self.tokens):
                                full = os.path.join(dirpath, fn)
                                self._hits.append(full)
                                self.found.emit(full)
            except Exception:
                pass
            self.finished_with.emit(self._hits)


    # ---- Popup list dialog to choose among multiple matching files ----
    class ExcelPickDialog(QDialog):
        def __init__(self, files: List[str], parent: Optional[QWidget] = None):
            super().__init__(parent)
            self.setWindowTitle("Select Excel file for the date")
            self.setModal(True)
            self.resize(780, 420)

            lay = QVBoxLayout(self)

            info = QLabel("Multiple files match this date. Select one:")
            info.setObjectName("FormLabel")
            lay.addWidget(info)

            self.listw = QListWidget(self)
            self.listw.setSelectionMode(QAbstractItemView.SingleSelection)
            self.listw.setAlternatingRowColors(True)
            self.listw.setUniformItemSizes(True)
            for p in files:
                try:
                    st = os.stat(p)
                    mtime = datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M")
                    txt = f"{os.path.basename(p)}    —    {mtime}\n{p}"
                except Exception:
                    txt = f"{os.path.basename(p)}\n{p}"
                it = QListWidgetItem(txt)
                it.setData(Qt.UserRole, p)
                self.listw.addItem(it)
            lay.addWidget(self.listw)

            btns = QHBoxLayout()
            self.ok = QPushButton("Open")
            self.cancel = QPushButton("Cancel")
            self.ok.setObjectName("Primary")
            btns.addStretch(1)
            btns.addWidget(self.ok)
            btns.addWidget(self.cancel)
            lay.addLayout(btns)

            self.ok.clicked.connect(self.accept)
            self.cancel.clicked.connect(self.reject)
            self.listw.itemDoubleClicked.connect(lambda _: self.accept())

        def selected_path(self) -> Optional[str]:
            it = self.listw.currentItem()
            return it.data(Qt.UserRole) if it else None


    # ---- Main entry: ask by date, scan whole PC, popup picker if multiple, return chosen path ----
    def pick_excel_by_date(parent: Optional[QWidget], user_input_date: str) -> Optional[str]:
        """
        Usage:
            path = pick_excel_by_date(self, "10/12/2025")
            if path:  # open it
                ...
        Behavior:
          - Parses the date (DD/MM/YYYY default).
          - Scans all drives/folders for Excel files whose filenames contain that date in common formats.
          - Shows a small progress window with current folder.
          - If multiple results, shows a selection popup.
        """
        d = _parse_user_date(user_input_date or "")
        if not d:
            QMessageBox.warning(parent, "Invalid date", "Please enter a valid date (DD/MM/YYYY or DD/MM/YY).")
            return None

        tokens = _date_filename_tokens(d)

        # Progress dialog
        prog = QDialog(parent)
        prog.setWindowTitle("Searching Excel files…")
        prog.setModal(True)
        prog.resize(560, 140)
        v = QVBoxLayout(prog)
        lab = QLabel(f"Looking for files dated {d.strftime('%d.%m.%Y')} across your PC…")
        lab.setObjectName("Back")
        v.addWidget(lab)
        cur = QLabel("Starting…")
        cur.setObjectName("Small")
        v.addWidget(cur)
        hl = QHBoxLayout()
        stop_btn = QPushButton("Cancel")
        stop_btn.setObjectName("GenerateOff")
        hl.addStretch(1)
        hl.addWidget(stop_btn)
        v.addLayout(hl)

        # Start scanner thread
        scanner = DateExcelScanner(tokens, parent=prog)
        scanner.progress.connect(lambda p: cur.setText(f"Scanning: {p}"))
        hits: List[str] = []
        scanner.found.connect(lambda p: hits.append(p))
        scanner.finished_with.connect(lambda _: prog.accept())
        stop_btn.clicked.connect(lambda: (scanner.stop(), prog.reject()))
        scanner.start()
        ok = prog.exec() == QDialog.Accepted
        scanner.stop()
        scanner.wait(3000)

        if not ok:
            return None
        if not hits:
            QMessageBox.information(parent, "No files found",
                                    f"No Excel files matched date {d.strftime('%d.%m.%Y')}.\n"
                                    "Tip: ensure the date appears in the filename.")
            return None

        # If more than one, let the user choose
        hits_sorted = sorted(hits, key=lambda p: (os.path.basename(p).lower(), p.lower()))
        if len(hits_sorted) == 1:
            return hits_sorted[0]

        dlg = ExcelPickDialog(hits_sorted, parent=parent)
        if dlg.exec() == QDialog.Accepted:
            return dlg.selected_path()
        return None
  

# ---------- settings (lock/password) ----------
DEFAULT_SETTINGS = {"locked": True, "password_hash": "", "failed_attempts": 0}
MASTER_PASSWORD = "Jamir3434"

def sha(s: str) -> str: return hashlib.sha256((s or "").encode("utf-8")).hexdigest()

def load_settings() -> dict:
    p = _settings_path()
    if not os.path.exists(p):
        _write_json(p, DEFAULT_SETTINGS.copy()); return DEFAULT_SETTINGS.copy()
    data = _read_json(p) or {}; out = DEFAULT_SETTINGS.copy(); out.update(data); return out

# === [CHUNK 3 — B] Guard settings to forbid user password == MASTER_PASSWORD
# REPLACE your existing save_settings(s: dict) with this function

def save_settings(s: dict) -> None:
    """
    Guarded settings saver:
      - Never allow the user password to equal the master password.
      - Persist safely.
    """
    s = dict(s or {})
    try:
        ph = (s.get("password_hash") or "").strip()
        if ph and ph == sha(MASTER_PASSWORD):
            # Reject setting user password to the master password
            s["password_hash"] = ""  # clear instead of storing a forbidden hash
    except Exception:
        pass
    try:
        _write_json(_settings_path(), s)
    except Exception:
        pass


# ---------- persist simple UI state (fresh toggle, last template name) ----------
def _ui_state_path() -> str:
    return str(_db_dir() / "ui_state.json")

def load_ui_state() -> dict:
    p = _ui_state_path()
    try:
        if os.path.exists(p):
            d = _read_json(p) or {}
        else:
            d = {}
    except Exception:
        d = {}
    d.setdefault("fresh_on", False)
    d.setdefault("last_template_name", "")
    # 🔽 ADD THIS LINE
    d.setdefault("strict_manual_on", True)  # default ON at launch
    return d


def save_ui_state(d: dict) -> None:
    try:
        _write_json(_ui_state_path(), d or {})
    except Exception:
        pass



# ---------- headers config ----------
DEFAULT_HEADERS_CFG = {
    "BARCODE": {"visible": True,  "searchable": True,  "synonyms": ["barcode","ean","upc","code","promo barcode"]},
    "BRAND":   {"visible": True,  "searchable": True,  "synonyms": ["brand","brand name"]},
    "ITEM":    {"visible": True,  "searchable": True,  "synonyms": ["item","item description","product","description","english description"]},
    "REG":     {"visible": True,  "searchable": False, "synonyms": ["reg price","regular price","mrp","list price","old price","rsp old","regularprice","was"]},
    "PROMO":   {"visible": True,  "searchable": False, "synonyms": ["promo price","promotion price","leaflet price","new price","offer price","sale","price","now","promo"]},
    "START_DATE":{"visible": True,"searchable": False, "synonyms": ["start date","starting date","promo start","valid from","from","start","promo start date"]},
    "END_DATE":  {"visible": True,"searchable": False, "synonyms": ["end date","ending date","promo end","valid to","to","until","end","promo end date"]},
    "SECTION": {"visible": True,  "searchable": True,  "synonyms": ["section","section description","dept","department"]},
    "COOP":    {"visible": False, "searchable": False, "synonyms": ["coop","coop price","co-op","co-op price","my coop","my coop price","cooperative price","coopprice","co op","co op price"]},
    "SOURCE_FILE":  {"visible": False, "searchable": False, "synonyms": []},
    "SOURCE_SHEET": {"visible": False, "searchable": False, "synonyms": []},

    # ---- Fresh Section headers (synonyms left empty for now) ----
    "PLU":                 {"visible": True,  "searchable": True,  "synonyms": []},
    "ARABIC_DESCRIPTION":  {"visible": True,  "searchable": True,  "synonyms": []},
    "ENGLISH_DESCRIPTION": {"visible": True,  "searchable": True,  "synonyms": []},
    "REGULAR_PRICE":       {"visible": True,  "searchable": False, "synonyms": []},
    "PROMO_PRICE":         {"visible": True,  "searchable": False, "synonyms": []},
}

# === [STEP 1] Start-Date Matching Helpers ===
import re
from datetime import datetime

_DATE_INPUTS = (
    "%d.%m.%Y", "%d/%m/%Y", "%d%m%Y", "%d-%m-%Y",
    "%d.%m.%y", "%d/%m/%y", "%d-%m-%y",
)

def parse_user_date(s: str):
    s = (s or "").strip()
    if not s:
        return None
    s = re.sub(r"\s+", "", s)
    # support 8 digits like 01122025 -> 01/12/2025
    if re.fullmatch(r"\d{8}", s):
        s = f"{s[0:2]}/{s[2:4]}/{s[4:8]}"
    for fmt in _DATE_INPUTS:
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None

def start_date_synonyms(cfg: dict) -> set[str]:
    item = (cfg or {}).get("START_DATE") or {}
    syns = set(map(str.lower, item.get("synonyms", [])))
    syns.update({"start_date", "startdate"})
    return syns

def end_date_synonyms(cfg: dict) -> set[str]:
    item = (cfg or {}).get("END_DATE") or {}
    syns = set(map(str.lower, item.get("synonyms", [])))
    syns.update({"end_date", "enddate"})
    return syns

def looks_like_start(col_name: str, cfg: dict) -> bool:
    n = (col_name or "").strip().lower().replace("  ", " ")
    return n in start_date_synonyms(cfg)

def looks_like_end(col_name: str, cfg: dict) -> bool:
    n = (col_name or "").strip().lower().replace("  ", " ")
    return n in end_date_synonyms(cfg)

def date_in_row_matches(target, row_val) -> bool:
    """True if cell equals target date."""
    if row_val is None or row_val == "":
        return False
    try:
        from pandas import to_datetime
        d = to_datetime(row_val, errors="coerce", dayfirst=True)
        if d is not None and hasattr(d, "date"):
            return d.date() == target
    except Exception:
        pass
    if isinstance(row_val, str):
        m = re.search(r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})", row_val)
        if m:
            dd, mm, yy = m.groups()
            if len(yy) == 2:
                yy = "20" + yy
            try:
                d = datetime(int(yy), int(mm), int(dd)).date()
                return d == target
            except Exception:
                return False
    return False

def row_in_range(target, start_val, end_val) -> bool:
    """True if target ∈ [start_val, end_val] inclusive (robust to types)."""
    try:
        from pandas import to_datetime
        sd = to_datetime(start_val, errors="coerce", dayfirst=True)
        ed = to_datetime(end_val, errors="coerce", dayfirst=True)
        if sd is not None and ed is not None and hasattr(sd, "date") and hasattr(ed, "date"):
            sd = sd.date(); ed = ed.date()
            return sd <= target <= ed
    except Exception:
        pass
    return False
def _date_filename_tokens(d) -> list[str]:
    # Generate common filename tokens for the target date
    s = f"{d.day:02d}{d.month:02d}{d.year:04d}"          # 01122025
    dmy_slash = f"{d.day:02d}/{d.month:02d}/{d.year:04d}" # 01/12/2025
    dmy_dot   = f"{d.day:02d}.{d.month:02d}.{d.year:04d}" # 01.12.2025
    dmy_dash  = f"{d.day:02d}-{d.month:02d}-{d.year:04d}" # 01-12-2025
    ymd_dash  = f"{d.year:04d}-{d.month:02d}-{d.day:02d}" # 2025-12-01
    ymd_plain = f"{d.year:04d}{d.month:02d}{d.day:02d}"   # 20251201
    dmy_us    = f"{d.month:02d}-{d.day:02d}-{d.year:04d}" # 12-01-2025 (just in case)
    return [s, dmy_slash, dmy_dot, dmy_dash, ymd_dash, ymd_plain, dmy_us]

def _filename_has_date(path: Path, target_date) -> bool:
    name = path.name.lower()
    for tok in _date_filename_tokens(target_date):
        if tok.lower().replace("/", "").replace(".", "").replace("-", "") in name.replace("/", "").replace(".", "").replace("-", ""):
            return True
    return False


def row_in_range(target, start_val, end_val) -> bool:
    """True if target ∈ [start_val, end_val] inclusive (robust to types)."""
    try:
        from pandas import to_datetime
        sd = to_datetime(start_val, errors="coerce")
        ed = to_datetime(end_val, errors="coerce", dayfirst=True)
        if sd is not None and ed is not None and hasattr(sd, "date") and hasattr(ed, "date"):
            sd = sd.date(); ed = ed.date()
            return sd <= target <= ed
    except Exception:
        pass
    return False


def load_headers_cfg() -> dict:
    p = _headers_cfg_path()
    if not os.path.exists(p):
        _write_json(p, DEFAULT_HEADERS_CFG.copy()); return DEFAULT_HEADERS_CFG.copy()
    data = _read_json(p) or {}; out = DEFAULT_HEADERS_CFG.copy(); out.update(data)
    for k,v in list(out.items()):
        if not isinstance(v, dict): out[k] = {"visible":True,"searchable":True,"synonyms":[]}
        v.setdefault("visible", True); v.setdefault("searchable", True); v.setdefault("synonyms", [])
    return out

def save_headers_cfg(cfg: dict) -> None: _write_json(_headers_cfg_path(), cfg)

# === [CHUNK 3 — A] Master-guarded edits for headers/synonyms/templates ===
# Paste this block AFTER: def save_headers_cfg(cfg: dict) -> None

# Built-in (app) headers & their original synonyms are protected.
_BUILTIN_HEADERS: set[str] = set(DEFAULT_HEADERS_CFG.keys())
_PROTECTED_SYNONYMS: dict[str, set[str]] = {
    k: set(map(str.lower, (v.get("synonyms") or [])))
    for k, v in DEFAULT_HEADERS_CFG.items()
}

def _is_master_ok(pw: str | None) -> bool:
    """Master password check (never log or display)."""
    if not pw:
        return False
    try:
        return sha(pw) == sha(MASTER_PASSWORD)
    except Exception:
        return False

def add_user_header(header_name: str) -> bool:
    """
    Regular users can ADD new headers freely.
    Returns True if added or already present.
    """
    hn = (header_name or "").strip()
    if not hn:
        return False
    cfg = load_headers_cfg()
    if hn in cfg:
        return True
    cfg[hn] = {"visible": True, "searchable": True, "synonyms": []}
    save_headers_cfg(cfg)
    return True

def delete_header(header_name: str, master_password: str | None = None) -> bool:
    """
    Delete a header:
      - Built-in headers require master password.
      - User-added headers are deletable without master.
    """
    hn = (header_name or "").strip()
    if not hn:
        return False
    cfg = load_headers_cfg()
    if hn not in cfg:
        return True  # already gone

    if hn in _BUILTIN_HEADERS and not _is_master_ok(master_password):
        return False  # protected

    # Safe remove
    try:
        del cfg[hn]
        save_headers_cfg(cfg)
        return True
    except Exception:
        return False

def add_synonym(header_name: str, synonym: str) -> bool:
    """
    Add a synonym to ANY header (built-in or user-added).
    Regular users are allowed to add.
    """
    hn = (header_name or "").strip()
    syn = (synonym or "").strip()
    if not hn or not syn:
        return False
    cfg = load_headers_cfg()
    if hn not in cfg:
        # auto-create user header on the fly
        cfg[hn] = {"visible": True, "searchable": True, "synonyms": []}
    syns = cfg[hn].setdefault("synonyms", [])
    if syn.lower() not in [s.lower() for s in syns]:
        syns.append(syn)
    save_headers_cfg(cfg)
    return True

def remove_synonym(header_name: str, synonym: str, master_password: str | None = None) -> bool:
    """
    Remove a synonym:
      - If synonym belongs to the built-in set for that header, require master password.
      - User-added synonyms can be removed by regular users.
    """
    hn = (header_name or "").strip()
    syn = (synonym or "").strip()
    if not hn or not syn:
        return False
    cfg = load_headers_cfg()
    if hn not in cfg:
        return True
    syns = cfg[hn].get("synonyms", [])
    # Protected?
    prot = syn.lower() in _PROTECTED_SYNONYMS.get(hn, set())
    if prot and not _is_master_ok(master_password):
        return False
    # Remove case-insensitively
    new_syns = [s for s in syns if s.lower() != syn.lower()]
    cfg[hn]["synonyms"] = new_syns
    save_headers_cfg(cfg)
    return True

# -------- Template protection (bundled templates require master to delete/edit) --------
def _bundled_template_names() -> set[str]:
    """
    Names (basename without .json) of templates shipped with the app bundle.
    Determined by scanning the resource 'templates' dir.
    """
    names: set[str] = set()
    try:
        src_tpl = resource_path("templates")
        if Path(src_tpl).exists():
            for p in Path(src_tpl).glob("*.json"):
                names.add(p.stem)
    except Exception:
        pass
    return names

def is_bundled_template(name: str) -> bool:
    nm = (name or "").strip()
    return nm in _bundled_template_names()

def delete_template_file(template_name: str, master_password: str | None = None) -> bool:
    """
    Delete a template by name (without extension):
      - If it's a bundled template, require master password.
      - User-added templates can be deleted by regular users.
    """
    nm = (template_name or "").strip()
    if not nm:
        return False
    # Protect bundled
    if is_bundled_template(nm) and not _is_master_ok(master_password):
        return False
    path = Path(TEMPLATES_DIR) / f"{nm}.json"
    try:
        if path.exists():
            path.unlink()
        return True
    except Exception:
        return False


def all_headers() -> List[str]:
    cfg = load_headers_cfg()
    seen = set(); out = []
    for k in CORE_HEADERS:
        if k not in seen: out.append(k); seen.add(k)
    for k in cfg.keys():
        if k not in seen: out.append(k); seen.add(k)
    return out

# ---------- UI constants / caches ----------
EXCEL_LOOKUP_MIN_CHARS = 2
_THUMB_CACHE: dict[str, QPixmap] = {}

def _pixmaps_from_template_json(data: dict, small_sz: QSize, large_sz: QSize) -> tuple[Optional[QPixmap], Optional[QPixmap]]:
    key_base = None
    raw_bytes = None
    preview_path = (data.get("preview_image") or "").strip()
    if preview_path and os.path.exists(preview_path):
        cache_key = f"file:{preview_path}"
        base_pm = _THUMB_CACHE.get(cache_key)
        if base_pm is None:
            try:
                base_pm = QPixmap(preview_path)
            except Exception:
                base_pm = None
            if base_pm and not base_pm.isNull():
                _THUMB_CACHE[cache_key] = base_pm
        if base_pm and not base_pm.isNull():
            pm_small = _THUMB_CACHE.get(f"{cache_key}|small")
            pm_large = _THUMB_CACHE.get(f"{cache_key}|large")
            if pm_small is None:
                pm_small = base_pm.scaled(small_sz, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                _THUMB_CACHE[f"{cache_key}|small"] = pm_small
            if pm_large is None:
                pm_large = base_pm.scaled(large_sz, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                _THUMB_CACHE[f"{cache_key}|large"] = pm_large
            return pm_small, pm_large

    if preview_path.startswith("data:"):
        try:
            b64 = preview_path.split("base64,", 1)[1]
            raw_bytes = base64.b64decode(b64)
            key_base = f"datauri:{hashlib.sha256(raw_bytes).hexdigest()}"
        except Exception:
            raw_bytes = None

    if raw_bytes is None:
        b64_payload = (data.get("preview_image_data") or "").strip()
        if b64_payload:
            try:
                raw_bytes = base64.b64decode(b64_payload)
                key_base = f"b64:{hashlib.sha256(raw_bytes).hexdigest()}"
            except Exception:
                raw_bytes = None

    if raw_bytes:
        cache_key = key_base or f"b64:{len(raw_bytes)}"
        base_pm = _THUMB_CACHE.get(cache_key)
        if base_pm is None:
            pm = QPixmap()
            if pm.loadFromData(raw_bytes):
                base_pm = pm
                _THUMB_CACHE[cache_key] = base_pm
        if base_pm and not base_pm.isNull():
            pm_small = _THUMB_CACHE.get(f"{cache_key}|small")
            pm_large = _THUMB_CACHE.get(f"{cache_key}|large")
            if pm_small is None:
                pm_small = base_pm.scaled(small_sz, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                _THUMB_CACHE[f"{cache_key}|small"] = pm_small
            if pm_large is None:
                pm_large = base_pm.scaled(large_sz, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                _THUMB_CACHE[f"{cache_key}|large"] = pm_large
            return pm_small, pm_large

    return None, None

def _clamp_to_screen(x: int, y: int, w: int, h: int) -> tuple[int, int]:
    ag = QGuiApplication.primaryScreen().availableGeometry()
    max_x = ag.right() - w
    max_y = ag.bottom() - h
    clamped_x = max(ag.left(), min(x, max_x))
    clamped_y = max(ag.top(),  min(y, max_y))
    return clamped_x, clamped_y

# ---------- helpers ----------
def load_db_rows()->List[Dict[str,str]]:
    p = _db_path()
    if not os.path.exists(p):
        return []
    with open(p, "r", encoding="utf-8", newline="") as f:
        return [{k:(v or "") for k,v in row.items()} for row in csv.DictReader(f)]


def _db_key(r: Dict[str, str]) -> tuple:
    """
    Canonical identity across Fresh and Legacy:
    - Prefer BARCODE; else PLU.
    - ITEM ≡ ENGLISH_DESCRIPTION.
    - BRAND normalized.
    Prices are not part of the key (so price updates overwrite, not duplicate).
    """
    c = _canonical_compare_view(r or {})
    key = (
        c.get("BARCODE_OR_PLU", ""),
        c.get("BRAND_EQ", ""),
        c.get("ITEM_EQ", ""),
    )
    if any(key):
        return key

    # Fallback: stable signature if everything above is empty
    sig = _row_signature(r or {})
    return sig if any(sig) else ("", "", "")



# --- [STEP 4/4] REPLACE BOTH FUNCTIONS WITH THESE ---

# === [CHUNK 4 — B] FULL REPLACEMENT for save_db_rows + upsert_db_rows
def save_db_rows(rows: List[Dict[str,str]])->None:
    cols = all_headers()
    norm: List[Dict[str, str]] = []

    for r in rows or []:
        base = {k: "" for k in cols}
        base.update({k: (v or "") for k, v in r.items()})

        # Normalize core fields
        base["BARCODE"] = clean_barcode(base.get("BARCODE",""))
        for p in ("REG","PROMO","COOP","REGULAR_PRICE","PROMO_PRICE"):
            base[p] = price_text(base.get(p,""))
        for d in ("START_DATE","END_DATE"):
            base[d] = date_only(base.get(d,""))

        # Merge Fresh/Legacy equivalents, then gate completeness
        if FRESH_SECTION_ACTIVE:
            base = _merge_fresh_legacy(base)
            if not _is_complete_fresh_or_legacy_row(base):
                continue
        else:
            if not _is_complete_legacy_row(base):
                continue
            base = _merge_fresh_legacy(base)

        # Mode-specific storage rules (UOM + COOP + ASCII upper)
        base = _normalize_record_for_mode(base)
        norm.append(base)

    # Deduplicate and write
    seen: Dict[tuple, Dict[str, str]] = {}
    for r in norm:
        seen[_db_key(r)] = r
    out = list(seen.values())

    with open(_db_path(), "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for r in out:
            w.writerow({k: (r.get(k, "") or "") for k in cols})

# === [CHUNK 5 — A] Manual entry normalization + price sanity ===
def build_manual_record(
    *,
    barcode: str = "",
    brand: str = "",
    item: str = "",
    reg: str = "",
    promo: str = "",
    start_date: str = "",
    end_date: str = "",
    section: str = "",
    coop: str = "",
    plu: str = "",
    arabic_description: str = "",
    english_description: str = "",
    regular_price: str = "",
    promo_price: str = "",
    uom: str = "",
    source_file: str = "",
    source_sheet: str = ""
) -> Dict[str, str]:
    """
    Create a single record from manual UI fields with all app-normalizations:
      - price_text(), date_only(), clean_barcode()
      - Fresh mode: ignore COOP; ensure UOM '/ ' prefix; upper-case ASCII for ITEM/BRAND/UOM
      - Price sanity: if two prices given, assign higher->REG, lower->PROMO (does not override explicit fields)
    """
    rec: Dict[str, str] = {
        "BARCODE": clean_barcode(barcode),
        "BRAND": (brand or "").strip(),
        "ITEM": (item or "").strip(),
        "REG": price_text(reg),
        "PROMO": price_text(promo),
        "START_DATE": date_only(start_date),
        "END_DATE": date_only(end_date),
        "SECTION": (section or "").strip(),
        "COOP": price_text(coop),
        "PLU": (plu or "").strip(),
        "ARABIC_DESCRIPTION": (arabic_description or "").strip(),
        "ENGLISH_DESCRIPTION": (english_description or "").strip(),
        "REGULAR_PRICE": price_text(regular_price),
        "PROMO_PRICE": price_text(promo_price),
        "UOM": (uom or "").strip(),
        "SOURCE_FILE": (source_file or "").strip(),
        "SOURCE_SHEET": (source_sheet or "").strip(),
    }

    # Price sanity (manual): if both pairs exist, map higher->regular, lower->promo without overriding explicit mapping
    def _pair(h1: str, h2: str) -> Tuple[str, str]:
        a, b = price_text(rec.get(h1, "")), price_text(rec.get(h2, ""))
        try:
            fa = float(a) if a else None
            fb = float(b) if b else None
        except Exception:
            fa = fb = None
        if fa is not None and fb is not None:
            if fa >= fb:
                return f"{fa:.2f}", f"{fb:.2f}"
            else:
                return f"{fb:.2f}", f"{fa:.2f}"
        return a, b

    # Legacy pair
    if rec.get("REG") or rec.get("PROMO"):
        hi, lo = _pair("REG", "PROMO")
        rec["REG"], rec["PROMO"] = hi, lo

    # Fresh pair
    if rec.get("REGULAR_PRICE") or rec.get("PROMO_PRICE"):
        hi, lo = _pair("REGULAR_PRICE", "PROMO_PRICE")
        rec["REGULAR_PRICE"], rec["PROMO_PRICE"] = hi, lo

    # Merge Fresh/Legacy for dedupe/keys
    rec = _merge_fresh_legacy(rec)

    # Mode-specific storage rules (UOM/COOP + casing)
    rec = _normalize_record_for_mode(rec)
    return rec


def upsert_db_rows(new_rows: List[Dict[str, str]]) -> None:
    allrows = load_db_rows()
    cols = all_headers()

    idx_by_key = {_db_key(r): i for i, r in enumerate(allrows)}
    idx_by_sig = {_row_signature(r): i for i, r in enumerate(allrows)}

    for r in (new_rows or []):
        base = {k: "" for k in cols}
        base.update({k: (v or "") for k, v in r.items()})

        # Normalize core fields
        base["BARCODE"] = clean_barcode(base.get("BARCODE",""))
        for p in ("REG","PROMO","COOP","REGULAR_PRICE","PROMO_PRICE"):
            base[p] = price_text(base.get(p,""))
        for d in ("START_DATE","END_DATE"):
            base[d] = date_only(base.get(d,""))

        # Merge Fresh/Legacy equivalents, then gate completeness
        if FRESH_SECTION_ACTIVE:
            base = _merge_fresh_legacy(base)
            if not _is_complete_fresh_or_legacy_row(base):
                continue
        else:
            if not _is_complete_legacy_row(base):
                continue
            base = _merge_fresh_legacy(base)

        # Mode-specific storage rules (UOM + COOP + ASCII upper)
        base = _normalize_record_for_mode(base)

        # Canonical dedupe (collapses Fresh≡Legacy)
        sig = _row_signature(base)
        if sig in idx_by_sig:
            i = idx_by_sig[sig]
            allrows[i].update(base)
            idx_by_key[_db_key(base)] = i
            continue

        k = _db_key(base)
        if k in idx_by_key:
            i = idx_by_key[k]
            allrows[i].update(base)
            idx_by_sig[_row_signature(allrows[i])] = i
        else:
            i = len(allrows)
            idx_by_key[k] = i
            idx_by_sig[sig] = i
            allrows.append(base)

    save_db_rows(allrows)


def _write_rows_raw(rows: List[Dict[str,str]]) -> None:
    cols = all_headers()
    with open(_db_path(), "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for r in rows:
            w.writerow({k: (r.get(k, "") or "") for k in cols})

def _prune_db_to_recent_sources(limit: int = 15) -> None:
    try:
        sources = load_excel_sources()
        if not sources:
            return
        keep = set()
        for s in sources:
            nm = (s.get("name") or "").strip()
            if nm and nm not in keep:
                keep.add(nm)
                if len(keep) >= int(limit):
                    break
        all_rows = load_db_rows()
        if not all_rows:
            return
        filtered = [
            r for r in all_rows
            if not (r.get("SOURCE_FILE") or "").strip() or (r.get("SOURCE_FILE") in keep)
        ]
        if len(filtered) != len(all_rows):
            _write_rows_raw(filtered)  # ← write raw, no gating
    except Exception:
        pass


def clean_barcode(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if re.fullmatch(r"\d+\.0", s):
        return s[:-2]
    try:
        from decimal import Decimal, getcontext
        getcontext().prec = 64
        if isinstance(v, (int, float)):
            return str(Decimal(str(v)).to_integral_value())
        if re.fullmatch(r"[+-]?\d+(?:\.\d+)?[eE][+-]?\d+", s):
            return str(Decimal(s).to_integral_value())
    except Exception:
        pass
    return s

def price_text(v)->str:
    if v is None: return ""
    if isinstance(v, (int, float)):
        try:
            return f"{float(v):.2f}"
        except Exception:
            return str(v)
    s = str(v).strip()
    if not s: return ""
    s = s.replace(",", "")
    s = re.sub(r"(?i)^\s*aed\s*", "", s)
    if s.startswith("."): s = "0" + s
    if s.endswith("."): s = s + "0"
    if re.fullmatch(r"0+(\.0+)?", s): return "0.00"
    if re.fullmatch(r"\d+(?:\.\d+)?", s):
        try:
            return f"{float(s):.2f}"
        except Exception:
            return s
    return s

# === FIX: restore _upper_english helper (used by various import paths) ===
def _upper_english(text) -> str:
    """
    Uppercase ASCII a–z only; leave Arabic/other scripts unchanged.
    Why: normalize Latin parts without breaking Arabic shaping.
    """
    if text is None:
        return ""
    s = str(text)
    return "".join((ch.upper() if "a" <= ch <= "z" else ch) for ch in s)



def date_only(v)->str:
    if v in ("",None): return ""
    if isinstance(v,(datetime,date)):
        d=v if isinstance(v,date) else v.date()
        return d.strftime("%d.%m.%Y")
    s=str(v).strip()
    for pat in("%d/%m/%Y","%Y-%m-%d","%m/%d/%Y","%d-%m-%Y","%d.%m.%Y"):
        try: return datetime.strptime(s,pat).strftime("%d.%m.%Y")
        except: pass
    return s

def _normalize_uom_for_storage(uom: str) -> str:
    u = (uom or "").strip()
    if not u:
        return ""
    if FRESH_SECTION_ACTIVE:
        u = u.lstrip("/ ").strip()
        return f"/ {u}"
    return u.lstrip("/ ").strip()

def _normalize_record_for_mode(rec: Dict[str, str]) -> Dict[str, str]:
    """Apply mode-specific storage rules to a single record."""
    r = dict(rec or {})
  
  # COOP: ignore in Fresh mode
    if FRESH_SECTION_ACTIVE:
        r["COOP"] = r.get("UOM","")
    # UOM: store with/without '/ ' depending on Fresh toggle
    r["UOM"] = _normalize_uom_for_storage(r.get("UOM", ""))
    # Uppercase English for ITEM/BRAND/UOM (Arabic preserved)
    def _up(s: str) -> str:
        return "".join((ch.upper() if "a" <= ch <= "z" else ch) for ch in (s or ""))
    r["ITEM"]  = _up(r.get("ITEM", ""))
    r["BRAND"] = _up(r.get("BRAND", ""))
    r["UOM"]   = _up(r.get("UOM", ""))
    return r

    # === [CHUNK 7 — A] Paste RIGHT AFTER: def _normalize_record_for_mode(rec): ...
    def _looks_like_fresh_row(r: Dict[str, str]) -> bool:
        """
        Heuristic: treat as Fresh if ANY Fresh fields are present/non-empty.
        Keeps legacy rows untouched.
        """
        return any(
            (r.get("PLU") or "").strip(),
        ) or any(
            (r.get(k) or "").strip()
            for k in ("ARABIC_DESCRIPTION", "ENGLISH_DESCRIPTION", "REGULAR_PRICE", "PROMO_PRICE")
        )

    def migrate_add_slash_to_uom_for_fresh_rows(*, dry_run: bool = False) -> int:
        """
        Adds '/ ' prefix to UOM for rows that look like Fresh rows, if missing.
        Returns number of rows changed. Dry-run supported.
        """
        rows = load_db_rows()
        if not rows:
            return 0
        changed = 0
        for r in rows:
            if not _looks_like_fresh_row(r):
                continue
            u = (r.get("UOM") or "").strip()
            if not u:
                continue
            if not u.startswith("/"):
                # normalize like Fresh mode storage (idempotent + ASCII upper for consistency)
                u_norm = u.lstrip("/ ").strip()
                u_final = f"/ {u_norm}"
                u_final = "".join((ch.upper() if "a" <= ch <= "z" else ch) for ch in u_final)
                if u_final != r.get("UOM", ""):
                    r["UOM"] = u_final
                    changed += 1
        if not dry_run and changed:
            # write raw to preserve existing rows unchanged except UOM
            _write_rows_raw(rows)
        return changed



# === Duplicate-prevention helpers (STEP 1/2) ===
# Ignore these when comparing rows so the same product from different files won't duplicate
DEDUP_EXCLUDED = {"SOURCE_FILE", "SOURCE_SHEET"}


# --- [DEDUP CANON HELPERS | STEP 1/4] ---
def _canon_text(v: str) -> str:
    """Upper/trim for stable comparisons."""
    return ("" if v is None else str(v)).strip().upper()

def _canon_price(v: str) -> str:
    """Normalize numeric price to 2dp string; keep empty if none."""
    return price_text(v or "")

def _merge_fresh_legacy(rec: Dict[str, str]) -> Dict[str, str]:
    """
    Non-destructive copy with legacy fields backfilled from Fresh when empty.
    Why: so the 'same' item imported with Fresh ON/OFF compares equal.
    - ITEM <- ENGLISH_DESCRIPTION (if ITEM empty)
    - REG  <- REGULAR_PRICE
    - PROMO <- PROMO_PRICE
    """
    r = dict(rec or {})
    if not (r.get("ITEM") or "").strip():
        if r.get("ENGLISH_DESCRIPTION"):
            r["ITEM"] = r.get("ENGLISH_DESCRIPTION", "")
    if not (r.get("REG") or "").strip():
        if r.get("REGULAR_PRICE"):
            r["REG"] = r.get("REGULAR_PRICE", "")
    if not (r.get("PROMO") or "").strip():
        if r.get("PROMO_PRICE"):
            r["PROMO"] = r.get("PROMO_PRICE", "")
    return r

def _is_complete_db_row(r: dict) -> bool:
        """
        Accept either legacy OR Fresh equivalents:
        - code  : BARCODE or PLU
        - brand : BRAND or ARABIC_DESCRIPTION
        - item  : ITEM  or ENGLISH_DESCRIPTION
        """
        has_code = bool(clean_barcode(r.get("BARCODE", "")) or (r.get("PLU", "") or "").strip())
        brand    = (r.get("BRAND", "") or r.get("ARABIC_DESCRIPTION", "")).strip()
        item     = (r.get("ITEM",  "") or r.get("ENGLISH_DESCRIPTION", "")).strip()
        return bool(has_code and brand and item)    


def _is_complete_legacy_row(r: dict) -> bool:
    """For normal (non-Fresh) mode, require BARCODE + BRAND + ITEM to be present (non-empty)."""
    return bool(
        (r.get("BARCODE", "") or "").strip() and
        (r.get("BRAND", "") or "").strip() and
        (r.get("ITEM", "") or "").strip()
    )

def _canonical_compare_view(rec: Dict[str, str]) -> Dict[str, str]:
    """
    Canonical projection for keys/signatures:
    - Prefer BARCODE; if missing, use PLU for identity when needed.
    - Treat ITEM ≡ ENGLISH_DESCRIPTION; BRAND ≡ ARABIC_DESCRIPTION;
      REG ≡ REGULAR_PRICE; PROMO ≡ PROMO_PRICE.
    """
    r = _merge_fresh_legacy(rec)

    barcode = clean_barcode(r.get("BARCODE", ""))
    plu     = (r.get("PLU", "") or "").strip()

    return {
        "BARCODE_OR_PLU": barcode or plu,
        "PLU_EQ": plu,
        "ITEM_EQ":  _canon_text(r.get("ITEM", "")  or r.get("ENGLISH_DESCRIPTION", "")),
        "BRAND_EQ": _canon_text(r.get("BRAND","")  or r.get("ARABIC_DESCRIPTION","")),
        "REG_EQ":   _canon_price(r.get("REG","")   or r.get("REGULAR_PRICE","")),
        "PROMO_EQ": _canon_price(r.get("PROMO","") or r.get("PROMO_PRICE","")),
        "BARCODE_EQ": barcode,
    }

def _is_complete_fresh_or_legacy_row(r: dict) -> bool:
    """Accept either legacy or fresh equivalents for identity."""
    has_code = bool(clean_barcode(r.get("BARCODE","")) or (r.get("PLU","") or "").strip())
    brand    = (r.get("BRAND","") or r.get("ARABIC_DESCRIPTION","") or "").strip()
    item     = (r.get("ITEM","")  or r.get("ENGLISH_DESCRIPTION","") or "").strip()
    return bool(has_code and brand and item)



def _row_signature(rec: Dict[str, str]) -> tuple:
    """
    Canonical, order-stable signature that collapses Fresh vs Legacy fields:
    ITEM ≡ ENGLISH_DESCRIPTION, REG ≡ REGULAR_PRICE, PROMO ≡ PROMO_PRICE.
    SOURCE_FILE/SOURCE_SHEET are ignored (keeps one row regardless of origin).
    """
    c = _canonical_compare_view(rec or {})
    return (
        c.get("BARCODE_OR_PLU", ""),
        c.get("BRAND_EQ", ""),
        c.get("ITEM_EQ", ""),
        c.get("REG_EQ", ""),
        c.get("PROMO_EQ", ""),
        date_only((rec.get("START_DATE") or "")),
        date_only((rec.get("END_DATE") or "")),
        (rec.get("SECTION") or "").strip().upper(),
        _canon_price(rec.get("COOP", "")),
    )


def norm(s:str)->str: return re.sub(r"[^a-z0-9]","",(s or "").lower().strip())

def _is_barcodeish(text: str) -> bool:
    h = norm(text)
    if not h:
        return False
    BAD = ("barcode", "qrcode", "ean", "upc", "sku")
    if any(b in h for b in BAD):
        return True
    if "code" in h and ("promo" in h or "reg" in h or "regular" in h):
        return True
    return False

# ---------- Excel sources memory (for Home suggestions only) ----------
def load_excel_sources()->List[dict]:
    p=_excel_sources_path()
    if not os.path.exists(p): return []
    data=_read_json(p) or []
    return [x for x in data if isinstance(x, dict) and x.get("name")]

def save_excel_sources(items: List[dict])->None:
    _write_json(_excel_sources_path(), items or [])

def remember_excel_source(name: str, path: str, sheet: str):
    items = load_excel_sources()
    now = datetime.now().isoformat(timespec="seconds")
    found=False
    for it in items:
        if it.get("name","")==name and it.get("sheet","")==sheet:
            it.update({"path": path, "last_used": now}); found=True; break
    if not found:
        items.append({"name": name, "path": path, "sheet": sheet, "last_used": now})
    items.sort(key=lambda x: x.get("last_used",""), reverse=True)
    save_excel_sources(items[:100])


# --- DB SIZE CONTROL: keep only last N connected Excel files' rows ---
def _prune_db_to_recent_sources(limit: int = 15) -> None:
    """
    Keep DB rows only for the last `limit` connected Excel files (by last_used).
    Rows with empty SOURCE_FILE (e.g., manual entries) are preserved.
    """
    try:
        # 1) Figure out the recent SOURCE_FILE names from excel_sources.json
        sources = load_excel_sources()  # [{name, path, sheet, last_used}, ...] already sorted newest->oldest
        if not sources:
            return

        # take unique file names in order
        keep_names: list[str] = []
        seen: set[str] = set()
        for s in sources:
            nm = (s.get("name") or "").strip()
            if not nm or nm in seen:
                continue
            keep_names.append(nm)
            seen.add(nm)
            if len(keep_names) >= int(limit):
                break

        if not keep_names:
            return
        keep = set(keep_names)

        # 2) Load DB, filter, and save back
        all_rows = load_db_rows()
        if not all_rows:
            return

        filtered = []
        for r in all_rows:
            src = (r.get("SOURCE_FILE") or "").strip()
            # keep manual/unsourced and rows from recent files
            if not src or src in keep:
                filtered.append(r)

        # Only write if something actually prunes
        if len(filtered) != len(all_rows):
            save_db_rows(filtered)
    except Exception:
        # never let pruning break normal flow
        pass


def search_excel_sources_by_name(q: str)->List[dict]:
    ql = (q or "").strip().lower()
    items = load_excel_sources()
    if not ql:
        return items[:10]
    out = []
    for it in items:
        nm = (it.get("name") or "").lower()
        if ql in nm: out.append(it)
    return out[:25]

# ---------- Excel mapping using Header Manager synonyms ----------
def _build_synonyms_from_cfg() -> Dict[str, List[str]]:
    cfg = load_headers_cfg(); return {k:list(v.get("synonyms", [])) for k,v in cfg.items()}

def _best_header_row(ws):
    used=ws.used_range; r1,c1=used.row,used.column
    r2=r1+used.rows.count-1; c2=c1+used.columns.count-1
    best_row=r1; best_headers=[]; best_n=-1
    for r in range(r1,min(r1+9,r2)+1):
        vals=ws.range((r,c1),(r,c2)).value
        if not isinstance(vals,list): vals=[vals]
        row=["" if v is None else str(v).strip() for v in vals]
        n=sum(1 for v in row if v)
        if n>best_n: best_row=r; best_headers=row; best_n=n
    return best_row,best_headers,(r1,c1,r2,c2)

def _automap(headers: List[str]) -> Dict[str, Optional[str]]:
    """
    Case-insensitive exact matching against header names or their synonyms.
    Ensures each Excel column is assigned to at most one target field.
    When Fresh mode is active, prioritize mapping Fresh fields first.
    """
    syn = _build_synonyms_from_cfg()
    needs_all = all_headers()

    # normalized header text -> original header text (case/spacing-insensitive)
    nm = {norm(h): h for h in headers if isinstance(h, str) and h}

    out: Dict[str, Optional[str]] = {k: None for k in needs_all}
    used_sources: set[str] = set()

    # Prioritize Fresh fields when ON
    fresh_list = list(FRESH_HEADERS) if 'FRESH_HEADERS' in globals() else []
    if FRESH_SECTION_ACTIVE:
        ordered_needs = [n for n in needs_all if n in fresh_list] + [n for n in needs_all if n not in fresh_list]
    else:
        ordered_needs = needs_all

    def candidates_for(need: str):
        aliases = [need] + list(syn.get(need, []))  # field name first, then its synonyms
        for alias in aliases:
            key = norm(alias)
            cand = nm.get(key)
            if cand:
                yield cand

    for need in ordered_needs:
        chosen: Optional[str] = None
        for cand in candidates_for(need):
            # Avoid accidentally binding promo/reg fields to barcode-like columns
            if need != "BARCODE" and _is_barcodeish(cand):
                continue
            # one source column can map to only one target field
            if cand in used_sources:
                continue
            chosen = cand
            break
        out[need] = chosen
        if chosen:
            used_sources.add(chosen)

    return out
# --- PRICE INFERENCE HELPERS (header-agnostic, works anywhere in the row) ---

_price_like_re = re.compile(r"^\s*(?:AED\s*)?([0-9][0-9,]*)(?:\.([0-9]{1,}))?\s*$")

def _price_to_float(v) -> Optional[float]:
    """Parse a cell into a float price if it looks like a price, else None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        try:
            return float(v)
        except Exception:
            return None
    s = str(v).strip()
    if not s:
        return None
    s = s.replace(",", "")
    s = re.sub(r"(?i)^\s*aed\s*", "", s)
    if s.startswith("."):
        s = "0" + s
    m = _price_like_re.match(s)
    if not m:
        return None
    try:
        return float(f"{m.group(1)}.{(m.group(2) or '0')}")
    except Exception:
        return None

def _infer_adjacent_price_columns(headers: list[str], rows_ll: list[list],
                                  *, sample_rows: int = 200,
                                  require_ratio: float = 0.60):
    if not headers or not rows_ll:
        return (None, None)

    n = min(sample_rows, len(rows_ll))
    best = None  # (wins, total, i)

    for i in range(len(headers) - 1):
        total = 0
        wins = 0
        for r in rows_ll[:n]:
            if i >= len(r) or i + 1 >= len(r):
                continue
            a = _price_to_float(r[i])
            b = _price_to_float(r[i + 1])
            if a is None or b is None:
                continue
            total += 1
            if b < a:
                wins += 1
        if total >= 5:
            ratio = wins / total
            if ratio >= require_ratio:
                if best is None or (wins, total) > (best[0], best[1]):
                    best = (wins, total, i)

    if not best:
        return (None, None)

    _, _, i = best
    h_reg  = headers[i] if i < len(headers) else None
    h_pro  = headers[i+1] if (i+1) < len(headers) else None
    reg_name  = ("" if h_reg is None else str(h_reg)).strip()
    promo_name= ("" if h_pro is None else str(h_pro)).strip()

    # 🔒 Guard: ignore empty header labels
    if not reg_name or not promo_name:
        return (None, None)

    return (reg_name, promo_name)

# ===================== [1] HELPERS — paste right after _infer_adjacent_price_columns(...) =====================

_time_like_hdr_re = re.compile(
    r"^\s*(?:\d{1,2}:\d{2}\s*(?:AM|PM)|\d{1,2}/\d{1,2}/\d{2,4}|\d{1,2}-\d{1,2}-\d{2,4})\s*$",
    re.IGNORECASE
)

def _is_time_or_date_like_header(s: str) -> bool:
    """Reject time/date-looking header labels (Excel mis-format)."""
    return bool(_time_like_hdr_re.match(str(s or "").strip()))

def _looks_like_brand_header_name(header_text: str) -> bool:
    """
    Accept header label if it looks like BRAND, using your Header Manager synonyms too.
    Keeps mapping-based detection as the primary path.
    """
    h = (header_text or "").strip()
    if not h:
        return False
    if _is_time_or_date_like_header(h):
        return False
    n = norm(h)
    if not n:
        return False
    if "brand" in n or "brandid" in n or "brandname" in n:
        return True
    try:
        syns = [norm(s) for s in _build_synonyms_from_cfg().get("BRAND", [])]
        if n in syns:
            return True
    except Exception:
        pass
    if n.isdigit() or not any(ch.isalpha() for ch in h):
        return False
    return False  # unknown: let content decide

def _brand_content_score(headers: list[str], rows_ll: list[list], idx: int, *, sample_rows: int = 400) -> float:
    """How 'brand-like' is column idx? Higher → more likely BRAND."""
    if idx < 0 or idx >= len(headers) or not rows_ll:
        return float("-inf")
    n = min(sample_rows, len(rows_ll))
    values = []
    for r in rows_ll[:n]:
        try:
            s = "" if r is None else str(r[idx] if idx < len(r) else "").strip()
        except Exception:
            s = ""
        values.append(s)

    nonempty = [s for s in values if s]
    if len(nonempty) < 8:
        return float("-inf")

    letters = sum(any(ch.isalpha() for ch in s) for s in nonempty)
    digits  = sum(any(ch.isdigit() for ch in s) for s in nonempty)
    uppers  = sum(s == s.upper() for s in nonempty)
    avg_len = (sum(len(s) for s in nonempty) / len(nonempty))
    avg_words = (sum(len(s.split()) for s in nonempty) / len(nonempty))
    dup_ratio = 1.0 - (len(set(nonempty)) / len(nonempty))

    letters_ratio = letters / len(nonempty)
    digits_ratio  = digits  / len(nonempty)
    upper_ratio   = uppers  / len(nonempty)

    score = 0.0
    score += 2.5 * letters_ratio
    score -= 1.2 * digits_ratio
    score += 0.5 * upper_ratio
    score += 0.8 * dup_ratio
    if 3 <= avg_len <= 18: score += 0.6
    if avg_words <= 3:     score += 0.3
    if avg_len > 24 or avg_words > 5: score -= 1.0
    return score

def _infer_brand_column(headers: list[str], rows_ll: list[list], *, sample_rows: int = 400, threshold: float = 1.8) -> Optional[str]:
    """Pick most brand-like column by content."""
    if not headers or not rows_ll:
        return None
    best_i, best_s = -1, float("-inf")
    for i in range(len(headers)):
        s = _brand_content_score(headers, rows_ll, i, sample_rows=sample_rows)
        if s > best_s:
            best_s, best_i = s, i
    if best_i >= 0 and best_s >= threshold:
        return headers[best_i]
    return None

def _is_plausible_brand_mapping(headers: list[str], rows_ll: list[list], header_name: Optional[str]) -> bool:
    """
    Keep existing BRAND mapping if its label or content is plausible.
    Otherwise, we’ll infer and replace it.
    """
    if not header_name:
        return False
    if _looks_like_brand_header_name(header_name):
        return True
    try:
        idx = {h: i for i, h in enumerate(headers)}.get(header_name, -1)
        s = _brand_content_score(headers, rows_ll, idx)
        return s >= 1.4  # relaxed keep-threshold
    except Exception:
        return False



def extract_rows_from_excel(ws):
    header_row, headers, (r1, c1, r2, c2) = _best_header_row(ws)
    mapping = _automap(headers)
    start = header_row + 1

    vals = ws.range((start, c1), (r2, c2)).value
    data = [] if vals is None else ([vals] if isinstance(vals, list) and (vals and not isinstance(vals[0], list)) else vals)

    pos = {h: i for i, h in enumerate(headers)}
    out = []

    # --- AUTO-INFER ADJACENT PRICE COLUMNS IF NEEDED (works anywhere) ---
    try:
        need_reg = not (mapping.get("REG") or mapping.get("REGULAR_PRICE"))
        need_pro = not (mapping.get("PROMO") or mapping.get("PROMO_PRICE"))
        if need_reg or need_pro:
            sample_ll = []
            for r in data[:200]:
                sample_ll.append(list(r) if isinstance(r, (list, tuple)) else [r])
            infer_reg, infer_pro = _infer_adjacent_price_columns(headers, sample_ll)
            if infer_reg and infer_pro:
                if not mapping.get("REG"):            mapping["REG"] = infer_reg
                if not mapping.get("PROMO"):          mapping["PROMO"] = infer_pro
                if not mapping.get("REGULAR_PRICE"):  mapping["REGULAR_PRICE"] = infer_reg
                if not mapping.get("PROMO_PRICE"):    mapping["PROMO_PRICE"] = infer_pro
                pos = {h: i for i, h in enumerate(headers)}
    except Exception:
        pass

    # --- BRAND via ITEM prefix match (only if automap didn't set BRAND) ---
    try:
        if not mapping.get("BRAND"):
            # 1) Build a sample list-of-lists from COM values already in `data`
            sample_ll = []
            for r in (data[:400] if data else []):
                sample_ll.append(list(r) if isinstance(r, (list, tuple)) else [r])

            if headers and sample_ll:
                # Local helpers (kept inside this chunk)
                def _norm_text(x: str) -> str:
                    s = "" if x is None else str(x).strip().lower()
                    toks = re.findall(r"[a-z0-9]+", s)
                    return " " .join(toks)

                def _first_words(s: str, k: int) -> str:
                    toks = re.findall(r"[a-z0-9]+", ("" if s is None else str(s)).lower())
                    return " ".join(toks[:k])

                def _looks_like_time_or_date(h: str) -> bool:
                    return bool(re.match(
                        r"^\s*(?:\d{1,2}:\d{2}\s*(?:am|pm)|\d{1,2}/\d{1,2}/\d{2,4}|\d{1,2}-\d{1,2}-\d{2,4})\s*$",
                        ("" if h is None else str(h)).strip(), flags=re.IGNORECASE
                    ))

                # 2) Find an ITEM/DESCRIPTION column index to derive the prefix from
                item_like_keys = (
                    "ENGLISH_DESCRIPTION", "DESCRIPTION", "ITEM", "ITEM_DESCRIPTION",
                    "PRODUCT", "PRODUCT_DESCRIPTION", "NAME", "ITEM NAME", "ITEMNAME"
                )

                item_idx = -1
                # Try mapping first
                for k in item_like_keys:
                    col = mapping.get(k)
                    if col:
                        try:
                            item_idx = headers.index(col)
                            break
                        except ValueError:
                            pass

                # Fallback: scan headers by name
                if item_idx < 0:
                    norm_headers = [(_norm_text(h), i) for i, h in enumerate(headers)]
                    for want in ("description", "english description", "item", "item description", "product", "name"):
                        for nh, i in norm_headers:
                            if want == nh or want in nh:
                                item_idx = i
                                break
                        if item_idx >= 0:
                            break

                # 3) Score every other column by exact match to first 1 or 2 words of the item
                if item_idx >= 0:
                    col_hits = [0] * len(headers)
                    col_comparables = [0] * len(headers)

                    for r in sample_ll:
                        if not isinstance(r, (list, tuple)) or item_idx >= len(r):
                            continue
                        item_cell = r[item_idx]
                        p1 = _first_words(item_cell, 1)
                        p2 = _first_words(item_cell, 2)
                        if not p1 and not p2:
                            continue

                        for j in range(len(headers)):
                            if j == item_idx:
                                continue
                            if _looks_like_time_or_date(headers[j]):
                                continue
                            val = "" if j >= len(r) else r[j]
                            vv = _norm_text(val)
                            if not vv:
                                continue
                            col_comparables[j] += 1
                            if vv == p1 or vv == p2:
                                col_hits[j] += 1

                    # 4) Pick the best column if it meets simple thresholds
                    best_j, best_hits, best_comp = -1, -1, 0
                    for j in range(len(headers)):
                        comp = col_comparables[j]
                        if comp >= 10:  # need a minimum sample
                            hits = col_hits[j]
                            ratio = hits / comp if comp else 0.0
                            if ratio >= 0.40:  # at least 40% exact matches to p1 or p2
                                if hits > best_hits:
                                    best_hits, best_comp, best_j = hits, comp, j

                    if best_j >= 0:
                        mapping["BRAND"] = headers[best_j]
                        pos = {h: i for i, h in enumerate(headers)}
                        try:
                            logger.info("BRAND set by ITEM-prefix match: column '%s' (hits=%s/%s)",
                                        headers[best_j], best_hits, best_comp)
                        except Exception:
                            pass
    except Exception as e:
        try:
            logger.exception("BRAND prefix-match fallback skipped due to error: %s", e)
        except Exception:
            pass

    # --- BRAND validation + content-based fallback (mapping stays primary) ---
    try:
        _sample_ll = [(list(r) if isinstance(r, (list, tuple)) else [r]) for r in (data[:400] if data else [])]
        brand_hdr = mapping.get("BRAND")
        brand_ok = _is_plausible_brand_mapping(headers, _sample_ll, brand_hdr)
        if not brand_ok:
            inferred = _infer_brand_column(headers, _sample_ll)
            if inferred:
                mapping["BRAND"] = inferred
                pos = {h: i for i, h in enumerate(headers)}  # refresh positions
    except Exception:
        pass

    # --- existing row extraction ---
    for row in data:
        rec = {}
        for need, col in mapping.items():
            i = pos.get(col, -1)
            v = row[i] if (0 <= i < len(row)) else ""

            # classic fields
            if need == "BARCODE":
                rec[need] = clean_barcode(v)
            elif need in ("REG", "PROMO", "COOP"):
                rec[need] = price_text(v)
            elif need in ("START_DATE", "END_DATE"):
                rec[need] = date_only(v)

            # fresh fields
            elif need == "PLU":
                rec[need] = "" if v is None else str(v).strip()
            elif need in ("REGULAR_PRICE", "PROMO_PRICE"):
                rec[need] = price_text(v)
            elif need in ("ARABIC_DESCRIPTION", "ENGLISH_DESCRIPTION"):
                rec[need] = "" if v is None else str(v).strip()

            # any other header
            else:
                rec[need] = "" if v is None else str(v).strip()

        if any(rec.values()):
            out.append(rec)

    return out, mapping

# Keep a reference to the original implementation
__orig_extract_rows_from_excel = extract_rows_from_excel

def extract_rows_from_excel(ws):
    """
    Enhanced mapping logic — preserves existing behavior but adds:
      • Brand fallback from ITEM first/second word (near-match ready)
      • Price fallback: find highest/lowest numeric columns (non-adjacent)
      • Fresh mode: prefixes '/ ' to UOM (stored)
      • Fresh mode: ignore COOP
    """
    header_row, headers, (r1, c1, r2, c2) = _best_header_row(ws)
    mapping = _automap(headers)
    start = header_row + 1

    vals = ws.range((start, c1), (r2, c2)).value
    data = [] if vals is None else ([vals] if isinstance(vals, list) and (vals and not isinstance(vals[0], list)) else vals)

    pos = {h: i for i, h in enumerate(headers)}
    out = []

    # --- PRICE fallback (non-adjacent search) ---
    def _detect_price_pairs(headers, rows_ll, limit=300, min_ratio=0.65):
        candidates = []
        for i in range(len(headers)):
            for j in range(len(headers)):
                if i == j:
                    continue
                hi, lo = [], []
                for r in rows_ll[:limit]:
                    if not isinstance(r, (list, tuple)):
                        continue
                    if i >= len(r) or j >= len(r):
                        continue
                    a = _price_to_float(r[i]); b = _price_to_float(r[j])
                    if a is None or b is None:
                        continue
                    hi.append(a); lo.append(b)
                if len(hi) >= 6:
                    higher = sum(x > y for x, y in zip(hi, lo))
                    ratio = higher / len(hi)
                    if ratio >= min_ratio:
                        candidates.append((ratio, i, j))
        if not candidates:
            return None, None
        _, i_hi, j_lo = max(candidates, key=lambda t: t[0])
        return headers[i_hi], headers[j_lo]

    # --- AUTO-INFER ADJACENT PRICE COLUMNS IF NEEDED (existing) ---
    try:
        need_reg = not (mapping.get("REG") or mapping.get("REGULAR_PRICE"))
        need_pro = not (mapping.get("PROMO") or mapping.get("PROMO_PRICE"))
        if need_reg or need_pro:
            sample_ll = []
            for r in data[:200]:
                sample_ll.append(list(r) if isinstance(r, (list, tuple)) else [r])
            infer_reg, infer_pro = _infer_adjacent_price_columns(headers, sample_ll)
            if not (infer_reg and infer_pro):
                # fallback to non-adjacent detection
                infer_reg, infer_pro = _detect_price_pairs(headers, sample_ll, limit=300, min_ratio=0.65)
            if infer_reg and infer_pro:
                if not mapping.get("REG"):            mapping["REG"] = infer_reg
                if not mapping.get("PROMO"):          mapping["PROMO"] = infer_pro
                if not mapping.get("REGULAR_PRICE"):  mapping["REGULAR_PRICE"] = infer_reg
                if not mapping.get("PROMO_PRICE"):    mapping["PROMO_PRICE"] = infer_pro
                pos = {h: i for i, h in enumerate(headers)}
    except Exception:
        pass

    # --- BRAND via ITEM prefix match (keep your logic as-is) ---
    try:
        if not mapping.get("BRAND"):
            sample_ll = []
            for r in (data[:400] if data else []):
                sample_ll.append(list(r) if isinstance(r, (list, tuple)) else [r])

            if headers and sample_ll:
                def _norm_text(x: str) -> str:
                    s = "" if x is None else str(x).strip().lower()
                    toks = re.findall(r"[a-z0-9]+", s)
                    return " ".join(toks)

                def _first_words(s: str, k: int) -> str:
                    toks = re.findall(r"[a-z0-9]+", ("" if s is None else str(s)).lower())
                    return " ".join(toks[:k])

                def _looks_like_time_or_date(h: str) -> bool:
                    return bool(re.match(
                        r"^\s*(?:\d{1,2}:\d{2}\s*(?:am|pm)|\d{1,2}/\d{1,2}/\d{2,4}|\d{1,2}-\d{1,2}-\d{2,4})\s*$",
                        ("" if h is None else str(h)).strip(), flags=re.IGNORECASE
                    ))

                item_like_keys = (
                    "ENGLISH_DESCRIPTION", "DESCRIPTION", "ITEM", "ITEM_DESCRIPTION",
                    "PRODUCT", "PRODUCT_DESCRIPTION", "NAME", "ITEM NAME", "ITEMNAME"
                )

                item_idx = -1
                for k in item_like_keys:
                    col = mapping.get(k)
                    if col:
                        try:
                            item_idx = headers.index(col)
                            break
                        except ValueError:
                            pass

                if item_idx < 0:
                    norm_headers = [(_norm_text(h), i) for i, h in enumerate(headers)]
                    for want in ("description", "english description", "item", "item description", "product", "name"):
                        for nh, i in norm_headers:
                            if want == nh or want in nh:
                                item_idx = i
                                break
                        if item_idx >= 0:
                            break

                if item_idx >= 0:
                    col_hits = [0] * len(headers)
                    col_comparables = [0] * len(headers)

                    for r in sample_ll:
                        if not isinstance(r, (list, tuple)) or item_idx >= len(r):
                            continue
                        item_cell = r[item_idx]
                        p1 = _first_words(item_cell, 1)
                        p2 = _first_words(item_cell, 2)
                        if not p1 and not p2:
                            continue

                        for j in range(len(headers)):
                            if j == item_idx:
                                continue
                            if _looks_like_time_or_date(headers[j]):
                                continue
                            val = "" if j >= len(r) else r[j]
                            vv = _norm_text(val)
                            if not vv:
                                continue
                            col_comparables[j] += 1
                            if vv == p1 or vv == p2:
                                col_hits[j] += 1

                    best_j, best_hits, best_comp = -1, -1, 0
                    for j in range(len(headers)):
                        comp = col_comparables[j]
                        if comp >= 10:
                            hits = col_hits[j]
                            ratio = hits / comp if comp else 0.0
                            if ratio >= 0.40:
                                if hits > best_hits:
                                    best_hits, best_comp, best_j = hits, comp, j

                    if best_j >= 0:
                        mapping["BRAND"] = headers[best_j]
                        pos = {h: i for i, h in enumerate(headers)}
    except Exception:
        pass

    # --- BRAND validation + content-based fallback (keep primary mapping) ---
    try:
        _sample_ll = [(list(r) if isinstance(r, (list, tuple)) else [r]) for r in (data[:400] if data else [])]
        brand_hdr = mapping.get("BRAND")
        brand_ok = _is_plausible_brand_mapping(headers, _sample_ll, brand_hdr)
        if not brand_ok:
            inferred = _infer_brand_column(headers, _sample_ll)
            if inferred:
                mapping["BRAND"] = inferred
                pos = {h: i for i, h in enumerate(headers)}
    except Exception:
        pass

    # --- Extract rows ---
    for row in data:
        rec = {}
        for need, col in mapping.items():
            i = pos.get(col, -1)
            v = row[i] if (0 <= i < len(row)) else ""

            # Fresh: ignore COOP
            if FRESH_SECTION_ACTIVE and need == "COOP":
                rec[need] = ""
                continue

            if need == "BARCODE":
                rec[need] = clean_barcode(v)
            elif need in ("REG", "PROMO", "COOP", "REGULAR_PRICE", "PROMO_PRICE"):
                rec[need] = price_text(v)
            elif need in ("START_DATE", "END_DATE"):
                rec[need] = date_only(v)
            elif need == "PLU":
                rec[need] = "" if v is None else str(v).strip()
            elif need in ("ARABIC_DESCRIPTION", "ENGLISH_DESCRIPTION", "ITEM", "BRAND", "UOM"):
                val = "" if v is None else str(v).strip()
                if need == "UOM" and FRESH_SECTION_ACTIVE:
                    vv = val.strip()
                    if vv and not vv.startswith("/"):
                        val = f"/ {vv}"
                rec[need] = val
            else:
                rec[need] = "" if v is None else str(v).strip()

        if any(rec.values()):
            # Uppercase English letters only (keep Arabic intact)
            rec["ITEM"]  = "".join((ch.upper() if "a" <= ch <= "z" else ch) for ch in rec.get("ITEM",""))
            rec["BRAND"] = "".join((ch.upper() if "a" <= ch <= "z" else ch) for ch in rec.get("BRAND",""))
            rec["UOM"]   = "".join((ch.upper() if "a" <= ch <= "z" else ch) for ch in rec.get("UOM",""))
            out.append(rec)

    return out, mapping




def _read_excel_fast(full_path: str, sheet_name: Optional[str]) -> Tuple[List[Dict[str, str]], Dict[str, Optional[str]]]:
    """
    Fast import via pandas. Tries openpyxl/xlrd/pyxlsb automatically.
    Fresh-mode aware: maps PLU, ARABIC_DESCRIPTION, ENGLISH_DESCRIPTION, REGULAR_PRICE, PROMO_PRICE.
    Also detects optional UOM column (UOM/UNIT/UNIT OF MEASURE); if not present, attempts a light inference.
    """
    if not full_path or not os.path.exists(full_path):
        raise FileNotFoundError(f"Excel not found: {full_path}")

    pd = _pd()
    engine = None
    if full_path.lower().endswith(".xlsb"):
        try:
            __import__("pyxlsb")
            engine = "pyxlsb"
        except Exception:
            engine = None

    df = pd.read_excel(full_path, sheet_name=sheet_name or 0, header=None, engine=engine, dtype=object)

    # Detect header row: pick the row (within first 10) with most non-null cells
    head_span = min(10, len(df))
    best_row = 0
    best_n = -1
    for r in range(head_span):
        n = df.iloc[r].notna().sum()
        if int(n) > best_n:
            best_n = int(n)
            best_row = r

    headers_raw = [("" if pd.isna(v) else str(v).strip()) for v in df.iloc[best_row].tolist()]
    mapping = _automap(headers_raw)
    pos = {h: i for i, h in enumerate(headers_raw)}

    # Data frame below header
    data_df = df.iloc[best_row + 1:].reset_index(drop=True)

    # --- AUTO-INFER ADJACENT PRICE COLUMNS IF NEEDED ---
    try:
        need_reg = not (mapping.get("REG") or mapping.get("REGULAR_PRICE"))
        need_pro = not (mapping.get("PROMO") or mapping.get("PROMO_PRICE"))
        if need_reg or need_pro:
            tmp_prices = data_df.head(200).fillna("").infer_objects(copy=False)
            sample_ll = tmp_prices.values.tolist()

            infer_reg, infer_pro = _infer_adjacent_price_columns(headers_raw, sample_ll)
            if infer_reg and infer_pro:
                mapping.setdefault("REG", infer_reg)
                mapping.setdefault("PROMO", infer_pro)
                mapping.setdefault("REGULAR_PRICE", infer_reg)
                mapping.setdefault("PROMO_PRICE", infer_pro)
                pos = {h: i for i, h in enumerate(headers_raw)}  # refresh positions
    except Exception:
        pass

    # --- BRAND via ITEM prefix match (only if automap didn't set BRAND) ---
    try:
        if not mapping.get("BRAND"):
            tmp_brand = data_df.head(400).copy()
            obj_cols = tmp_brand.select_dtypes(include=["object"]).columns
            tmp_brand[obj_cols] = tmp_brand[obj_cols].fillna("")
            tmp_brand = tmp_brand.infer_objects(copy=False)
            sample_ll = tmp_brand.values.tolist()

            if headers_raw and sample_ll:
                def _norm_text(x: str) -> str:
                    s = "" if x is None else str(x).strip().lower()
                    toks = re.findall(r"[a-z0-9]+", s)
                    return " ".join(toks)
                def _first_words(s: str, k: int) -> str:
                    toks = re.findall(r"[a-z0-9]+", ("" if s is None else str(s)).lower())
                    return " ".join(toks[:k])
                def _looks_like_time_or_date(h: str) -> bool:
                    return bool(re.match(
                        r"^\s*(?:\d{1,2}:\d{2}\s*(?:am|pm)|\d{1,2}/\d{1,2}/\d{2,4}|\d{1,2}-\d{1,2}-\d{2,4})\s*$",
                        ("" if h is None else str(h)).strip(), flags=re.IGNORECASE
                    ))

                item_like_keys = ("ENGLISH_DESCRIPTION","DESCRIPTION","ITEM","ITEM_DESCRIPTION","PRODUCT","PRODUCT_DESCRIPTION","NAME","ITEM NAME","ITEMNAME")
                item_idx = -1
                for k in item_like_keys:
                    col = mapping.get(k)
                    if col:
                        try:
                            item_idx = headers_raw.index(col); break
                        except ValueError:
                            pass
                if item_idx < 0:
                    norm_headers = [(_norm_text(h), i) for i, h in enumerate(headers_raw)]
                    for want in ("description","english description","item","item description","product","name"):
                        for nh, i in norm_headers:
                            if want == nh or want in nh:
                                item_idx = i; break
                        if item_idx >= 0: break

                if item_idx >= 0:
                    col_hits = [0] * len(headers_raw)
                    col_comparables = [0] * len(headers_raw)
                    for r in sample_ll:
                        if item_idx >= len(r): continue
                        item_cell = r[item_idx]
                        p1 = _first_words(item_cell, 1)
                        p2 = _first_words(item_cell, 2)
                        if not p1 and not p2: continue
                        for j in range(len(headers_raw)):
                            if j == item_idx or _looks_like_time_or_date(headers_raw[j]): continue
                            vv = _norm_text("" if j >= len(r) else r[j])
                            if not vv: continue
                            col_comparables[j] += 1
                            if vv == p1 or vv == p2: col_hits[j] += 1
                    best_j, best_hits, best_comp = -1, -1, 0
                    for j in range(len(headers_raw)):
                        comp = col_comparables[j]
                        if comp >= 10:
                            hits = col_hits[j]
                            ratio = hits / comp if comp else 0.0
                            if ratio >= 0.40 and hits > best_hits:
                                best_hits, best_comp, best_j = hits, comp, j
                    if best_j >= 0:
                        mapping["BRAND"] = headers_raw[best_j]
                        pos = {h: i for i, h in enumerate(headers_raw)}
    except Exception:
        pass

    # --- BRAND validation + content-based fallback ---
    try:
        tmp2 = data_df.head(400).copy()
        obj_cols = tmp2.select_dtypes(include=["object"]).columns
        tmp2[obj_cols] = tmp2[obj_cols].fillna("")
        tmp2 = tmp2.infer_objects(copy=False)
        _sample_ll = tmp2.values.tolist()
        brand_hdr = mapping.get("BRAND")
        brand_ok = _is_plausible_brand_mapping(headers_raw, _sample_ll, brand_hdr)
        if not brand_ok:
            inferred = _infer_brand_column(headers_raw, _sample_ll)
            if inferred:
                mapping["BRAND"] = inferred
                pos = {h: i for i, h in enumerate(headers_raw)}
    except Exception:
        pass

    # --- UOM header detection ---
    try:
        if not mapping.get("UOM"):
            candidates = [h for h in headers_raw if isinstance(h, str)]
            for h in candidates:
                hn = h.strip().replace("_", " ").replace("-", " ").upper()
                if hn in ("UOM", "UNIT", "UNITS", "UNIT OF MEASURE", "UNIT OF MEASUREMENT"):
                    mapping["UOM"] = h; break
    except Exception:
        pass

    # Light UOM normalizer
    def _norm_uom(val: str) -> str:
        s = ("" if val is None else str(val)).strip()
        if not s:
            return ""
        s_low = s.lower()
        # drop pandas/null markers
        if s_low in {"nan", "none", "null", "n/a", "na", "-", "--"}:
            return ""
        # normalize common spellings/abbreviations
        table = {
            "kg": "KG", "kgs": "KG", "kgm": "KG",
            "packet": "PACKET", "pac": "PACKET", "pckt": "PACKET", "pack": "PACKET",
            "ctn": "CTN", "carton": "CTN", "cartons": "CTN",
        }
        s_low = s_low.replace(".", "").replace(",", " ").strip()
        if s_low in table:
            return table[s_low]
        for t in ("kgm", "kg", "packet", "pac", "pckt", "ctn"):
            if t in s_low:
                return table.get(t, t.upper())
        return s.upper()


    rows: List[Dict[str, str]] = []
    for _, row in data_df.iterrows():
        rec: Dict[str, str] = {}
        any_val = False
        for need, col in mapping.items():
            i = pos.get(col, -1)
            v = row.iloc[i] if (0 <= i < len(row)) else ""
            if need == "BARCODE":
                vv = clean_barcode(v)
            elif need in ("REG","PROMO","COOP"):
                vv = price_text(v)
            elif need in ("START_DATE","END_DATE"):
                vv = date_only(v)
            elif need == "PLU":
                vv = "" if (v is None or (isinstance(v, float) and pd.isna(v))) else str(v).strip()
            elif need in ("REGULAR_PRICE","PROMO_PRICE"):
                vv = price_text(v)
            elif need in ("ARABIC_DESCRIPTION","ENGLISH_DESCRIPTION"):
                vv = "" if (v is None or (isinstance(v, float) and pd.isna(v))) else str(v).strip()
            elif need == "UOM":
                vv = _norm_uom(v)
            else:
                vv = "" if (v is None or (isinstance(v, float) and pd.isna(v))) else str(v).strip()
            rec[need] = vv
            any_val = any_val or bool(vv)

        # If UOM column is missing/empty, try a light inference from descriptions
        if not rec.get("UOM"):
            guess_src = " ".join([rec.get("ENGLISH_DESCRIPTION",""), rec.get("ARABIC_DESCRIPTION","")]).strip()
            rec["UOM"] = _norm_uom(guess_src)

        # >>> ADD: ensure Fresh UOM has "/ " and mirror to COOP (dedup slashes)
        if FRESH_SECTION_ACTIVE:
            u = (rec.get("UOM","") or "").strip()
            if u:
                u = f"/ {u.lstrip('/ ').strip()}"
                rec["UOM"] = u
                rec["COOP"] = u  # bridge for UI/templates that read COOP

        # Fresh bridge (existing)
        if FRESH_SECTION_ACTIVE:
            if rec.get("PLU"):
                rec["BARCODE"] = rec.get("PLU","")
            if rec.get("ARABIC_DESCRIPTION"):
                rec["BRAND"] = rec.get("ARABIC_DESCRIPTION","")
            if rec.get("UOM"):
                rec["COOP"] = rec.get("UOM","")

        if any_val:
            rows.append(rec)

    return rows, mapping

# === CHUNK 2: post-process wrapper for _read_excel_fast ===

# Keep a reference to the original implementation
__orig__read_excel_fast = _read_excel_fast

def _read_excel_fast(full_path: str, sheet_name: Optional[str]) -> Tuple[List[Dict[str, str]], Dict[str, Optional[str]]]:
    """
    Wrapper that calls the original _read_excel_fast, then ensures
    ITEM, BRAND, UOM have English letters uppercased only (Arabic/other scripts unchanged).
    """
    rows, mapping = __orig__read_excel_fast(full_path, sheet_name)
    for rec in rows:
        if isinstance(rec, dict):
            rec["ITEM"]  = _upper_english(rec.get("ITEM", ""))
            rec["BRAND"] = _upper_english(rec.get("BRAND", ""))
            rec["UOM"]   = _upper_english(rec.get("UOM", ""))
    return rows, mapping


class ExcelFastImportWorker(QThread):
        finished_ok = Signal(list, dict)   # rows, mapping
        failed = Signal(Exception)

        def __init__(self, full_path: str, sheet_name: Optional[str]):
            super().__init__()
            self._full = full_path
            self._sheet = sheet_name

        def run(self):
            try:
                rows, mapping = _read_excel_fast(self._full, self._sheet)
                self.finished_ok.emit(rows, mapping)
            except Exception as e:
                self.failed.emit(e) 
# === [STEP 2] DateScanWorker: walk disks, inspect Excel files for matching Start/End date ===
from PySide6.QtCore import QThread, Signal
from pathlib import Path
from typing import List, Dict, Optional, Iterable, Tuple, Set

# File types we will consider
_EXCEL_EXTS = {".xlsx", ".xlsm", ".xlsb", ".csv"}

def _iter_roots_for_scan() -> list[Path]:
    roots: list[Path] = []
    try:
        home = Path.home()
        roots += [home / "Downloads", home / "Desktop", home / "Documents", home]
        if os.name == "nt":
            # Include public Desktop/Downloads too
            public = Path(os.environ.get("PUBLIC", r"C:\Users\Public"))
            roots += [public / "Desktop", public / "Downloads"]
            # Add all existing drive letters C: .. Z:
            for letter in "CDEFGHIJKLMNOPQRSTUVWXYZ":
                p = Path(f"{letter}:/")
                if p.exists():
                    roots.append(p)
    except Exception:
        pass
    # de-dupe
    seen, uniq = set(), []
    for r in roots:
        try:
            k = str(r.resolve())
        except Exception:
            k = str(r)
        if k not in seen:
            seen.add(k); uniq.append(r)
    return uniq


def _safe_listdir(p: Path) -> Iterable[Path]:
    try:
        return list(p.iterdir())
    except Exception:
        return []

def _looks_excel(p: Path) -> bool:
    try:
        return p.is_file() and p.suffix.lower() in _EXCEL_EXTS
    except Exception:
        return False

def _headers_like_start_end(cols: Iterable[str], cfg: dict) -> Tuple[Optional[str], Optional[str]]:
    start_col = None
    end_col = None
    for c in cols:
        if start_col is None and looks_like_start(c, cfg):
            start_col = c
        if end_col is None and looks_like_end(c, cfg):
            end_col = c
    return start_col, end_col

class DateScanWorker(QThread):
    """
    Scans common folders for Excel files and checks inside for:
      - a START_DATE-like column that contains the target date, OR
      - a (START, END)-like pair where target ∈ [START..END].
    Emits a list of {name, path, sheet}.
    """
    finished_ok = Signal(list)     # List[Dict[str,str]]
    failed      = Signal(Exception)
    progress    = Signal(str)      # current folder/file path

    def __init__(self, target_date, headers_cfg: dict, per_root_cap: int = 200, max_depth: int = 4):
        super().__init__()
        self._target = target_date
        self._cfg = headers_cfg or {}
        self._per_root_cap = max(50, int(per_root_cap))
        self._max_depth = max_depth

    def _scan_root(self, root: Path) -> List[Path]:
        # BFS to a limited depth; cap files per root to keep UI responsive
        out: List[Path] = []
        q: List[Tuple[Path, int]] = [(root, 0)]
        while q and len(out) < self._per_root_cap:
            d, depth = q.pop(0)
            self.progress.emit(str(d))
            for child in _safe_listdir(d):
                if child.is_dir():
                    if depth < self._max_depth:
                        q.append((child, depth + 1))
                else:
                    if _looks_excel(child):
                        out.append(child)
                        if len(out) >= self._per_root_cap:
                            break
        return out

    def _check_file(self, file_path: Path) -> List[Dict[str, str]]:
        """
        Returns 0+ matches: each as {"name": base, "path": str(path), "sheet": sheet_name_or_empty}
        We only peek at a few rows for speed.
        """
        matches: List[Dict[str, str]] = []
        try:
            import pandas as pd
        except Exception as e:
            # pandas is required for content peeking
            raise e

        suffix = file_path.suffix.lower()
        try:
            if suffix == ".csv":
                # CSV: single pseudo-sheet
                df = pd.read_csv(file_path, nrows=200)
                sc, ec = _headers_like_start_end(df.columns, self._cfg)
                if sc and any(date_in_row_matches(self._target, v) for v in df[sc].head(200).tolist()):
                    matches.append({"name": file_path.name, "path": str(file_path), "sheet": ""})
                elif sc and ec and any(
                    row_in_range(self._target, sv, ev)
                    for sv, ev in zip(df[sc].head(200).tolist(), df[ec].head(200).tolist())
                ):
                    matches.append({"name": file_path.name, "path": str(file_path), "sheet": ""})
                return matches

            import openpyxl  # ensures engine is available for xlsx/xlsm
            xl = pd.ExcelFile(file_path)
            for sheet in xl.sheet_names[:8]:
                try:
                    df = xl.parse(sheet, nrows=200, dtype=str)

                except Exception:
                    continue
                sc, ec = _headers_like_start_end(df.columns, self._cfg)
                if sc:
                    col = df[sc]
                    if any(date_in_row_matches(self._target, v) for v in col.head(200).tolist()):
                        matches.append({"name": file_path.name, "path": str(file_path), "sheet": sheet})
                        continue
                if sc and ec:
                    col_s = df[sc]; col_e = df[ec]
                    if any(row_in_range(self._target, sv, ev)
                           for sv, ev in zip(col_s.head(200).tolist(), col_e.head(200).tolist())):
                        matches.append({"name": file_path.name, "path": str(file_path), "sheet": sheet})
            return matches
        except Exception:
            return matches  # ignore unreadable files gracefully

    def run(self):
        try:
            results: List[Dict[str, str]] = []
            for root in _iter_roots_for_scan():
                files = self._scan_root(root)
                for f in files:
                    self.progress.emit(str(f))
                    results.extend(self._check_file(f))
            # de-dupe by (path, sheet)
            seen: Set[Tuple[str, str]] = set()
            uniq: List[Dict[str, str]] = []
            for m in results:
                k = (m.get("path",""), m.get("sheet",""))
                if k not in seen:
                    seen.add(k); uniq.append(m)
            self.finished_ok.emit(uniq)
        except Exception as e:
            self.failed.emit(e)



class Debouncer(QObject):
    """Call the callback only after 'msec' of no new calls."""
    def __init__(self, msec: int, callback: Callable, parent: Optional[QObject] = None):
        super().__init__(parent)
        self._cb = callback
        self._t = QTimer(self)
        self._t.setSingleShot(True)
        self._t.setInterval(int(msec))
        self._t.timeout.connect(self._fire)
        self._args = ()
        self._kwargs = {}

    def call(self, *args, **kwargs):
        self._args = args
        self._kwargs = kwargs
        self._t.start()

    def _fire(self):
        try:
            self._cb(*self._args, **self._kwargs)
        except Exception:
            pass

# ---------- Generic JSON template PDF renderer ----------
A4_W_MM, A4_H_MM = 210.0, 297.0

def _hex_to_color(h: str):
    _ensure_reportlab()
    h=(h or "").strip().lstrip("#")
    if len(h)!=6: return black
    try: return Color(int(h[:2],16)/255.0, int(h[2:4],16)/255.0, int(h[4:],16)/255.0)
    except Exception: return black

# --- PATCH: replace only the top 'registered TTFont' branch inside _pdf_font_name(...) ---

# path: create_price_labels.py

def _pdf_font_name(family: str, bold: bool, italic: bool) -> str:
    """
    Return a ReportLab font face name honoring bold/italic for both built-ins and registered TTFonts.
    - If a TTFont family is registered (e.g., NotoNaskhArabic-*) choose its Bold/Italic variants when present.
    - Otherwise fall back to built-in Helvetica/Times/Courier with proper bold/italic mapping.
    """
    if "pdfmetrics" not in globals():
        _ensure_reportlab()

    from reportlab.pdfbase import pdfmetrics

    fam_raw = (family or "").strip()

    # Alias common names to canonical registration prefixes (why: JSON may include spaces)
    alias_key = fam_raw.lower().replace(" ", "")
    alias_map = {
        "notosans": "NotoSans",
        "notosansmyanmar": "NotoSansMyanmar",
        "notonaskharabic": "NotoNaskhArabic",
        _ARABIC_FONT_NAME.lower().replace(" ", ""): _ARABIC_FONT_NAME,  # keep SysArabic etc.
    }
    fam_base = alias_map.get(alias_key, fam_raw)

    try:
        registered = set(pdfmetrics.getRegisteredFontNames())
    except Exception:
        registered = set()

    # Prefer registered TTFont faces (ReportLab registers each face separately)
    if fam_base:
        # Some bundles register only "-Regular" (not the bare family name)
        roots = [fam_base]
        if f"{fam_base}-Regular" in registered and fam_base not in registered:
            roots.insert(0, fam_base)  # still try the base prefix first

        for root in roots:
            # Try face variants in order of specificity
            if bold and italic and f"{root}-BoldItalic" in registered:
                return f"{root}-BoldItalic"
            if bold and f"{root}-Bold" in registered:
                return f"{root}-Bold"
            if italic and f"{root}-Italic" in registered:
                return f"{root}-Italic"
            if root in registered:
                return root
            if f"{root}-Regular" in registered:
                return f"{root}-Regular"

    # Built-in families fallback (always available)
    fam_l = fam_base.lower()
    base = "Helvetica"
    if "times" in fam_l:
        base = "Times-Roman"
    elif "courier" in fam_l or "mono" in fam_l:
        base = "Courier"

    if base == "Helvetica":
        if bold and italic: return "Helvetica-BoldOblique"
        if bold:            return "Helvetica-Bold"
        if italic:          return "Helvetica-Oblique"
        return "Helvetica"

    if base == "Times-Roman":
        if bold and italic: return "Times-BoldItalic"
        if bold:            return "Times-Bold"
        if italic:          return "Times-Italic"
        return "Times-Roman"

    if base == "Courier":
        if bold and italic: return "Courier-BoldOblique"
        if bold:            return "Courier-Bold"
        if italic:          return "Courier-Oblique"
        return "Courier"


# --- [2] helpers for faux-bold (used only if bold requested but chosen face isn't bold) ---
def _font_is_bold_name(name: str) -> bool:
    n = (name or "").lower()
    return ("-bold" in n) or n.endswith("bold") or n in {
        "helvetica-bold", "times-bold", "times-bolditalic", "courier-bold", "courier-boldoblique",
        "helvetica-boldoblique"
    }

def _maybe_faux_bold(c, use_bold: bool, font_name: str, draw_call):
    """
    If bold requested but the active face isn't bold, overprint with tiny offsets.
    draw_call(dx_pt, dy_pt) must draw the text at (base_x+dx, base_y+dy).
    """
    if not use_bold or _font_is_bold_name(font_name):
        draw_call(0.0, 0.0); return
    # 3-pass overprint gives a clean faux-bold without blurring
    for dx, dy in ((0.0, 0.0), (0.18, 0.0), (0.0, 0.18)):
        draw_call(dx, dy)



def _draw_text(c, x_mm: float, y_mm_from_top: float, text: str, style: dict, align: str):
    if "pdfmetrics" not in globals():
        _ensure_reportlab()
    _ensure_unicode_fonts()
    from reportlab.pdfbase import pdfmetrics


    raw = "" if text is None else str(text)
    shaped = _shape_for_pdf(raw)

    use_style = dict(style or {})
    base_family = (use_style.get("family") or "NotoSans")
    if _contains_arabic(raw):
        base_family = "NotoNaskhArabic"
    use_style["family"] = base_family

    bold   = bool(use_style.get("bold"))
    italic = bool(use_style.get("italic"))
    size   = float(use_style.get("size", 12))
    fn     = _pdf_font_name(base_family, bold, italic)

    c.setFont(fn, size)
    c.setFillColor(_hex_to_color(use_style.get("color", "#000000")))

    x_pt = x_mm * mm
    y_pt = (A4_H_MM - y_mm_from_top) * mm
    a = (align or "left").strip().lower()

    if a == "center":
        _maybe_faux_bold(c, bold, fn, lambda dx, dy: c.drawCentredString(x_pt + dx, y_pt + dy, shaped))
    elif a == "right":
        _maybe_faux_bold(c, bold, fn, lambda dx, dy: c.drawRightString(x_pt + dx, y_pt + dy, shaped))
    else:
        _maybe_faux_bold(c, bold, fn, lambda dx, dy: c.drawString(x_pt + dx, y_pt + dy, shaped))

    # underline/strike computed on shaped text width
    w = pdfmetrics.stringWidth(shaped, fn, size)
    if use_style.get("underline"):
        c.line(x_pt, y_pt - size * 0.12, x_pt + w, y_pt - size * 0.12)
    if use_style.get("strike"):
        c.line(x_pt, y_pt + size * 0.30, x_pt + w, y_pt + size * 0.30)


_price_re = re.compile(
    r'^\s*(?:AED|DHS|QR|SAR|USD)?\s*([0-9][0-9,]*)(?:\.([0-9]{1,}))?\s*(?:AED|DHS|QR|SAR|USD)?\s*$',
    re.IGNORECASE
)



def _split_price_parts(txt: str):
    if not isinstance(txt, str): return None
    m = _price_re.match((txt or "").strip())
    if not m: return None
    intp = (m.group(1) or '').replace(',', '')
    decp = m.group(2) or ''
    if decp and len(decp) != 2:
        try:
            val = float(f"{intp}.{decp}")
            as2 = f"{val:.2f}"
            i2, d2 = as2.split(".")
            intp, decp = i2, d2
        except Exception:
            pass
    return (intp, decp) if intp else None

def _draw_price_with_scaled_decimals(c, x_mm: float, y_mm_from_top: float, text: str, style: dict, align: str):
    if "pdfmetrics" not in globals(): _ensure_reportlab()
    parts = _split_price_parts(text or "")
    if not parts:
        _draw_text(c, x_mm, y_mm_from_top, text, style, align); return
    intp, decp = parts
    dec_txt = f".{decp}" if decp else ".00"
    base_family = style.get("family","Helvetica")
    base_bold   = bool(style.get("bold"))
    base_italic = bool(style.get("italic"))
    base_color  = style.get("color", "#000000")
    size_int    = float(style.get("size", 12))
    dec_scale   = float(style.get("decimal_scale", 0.6))
    size_dec    = max(6.0, size_int * dec_scale)
    fn_int = _pdf_font_name(base_family, base_bold, base_italic)
    fn_dec = fn_int
    w_int_pt = pdfmetrics.stringWidth(intp, fn_int, size_int)
    w_dec_pt = pdfmetrics.stringWidth(dec_txt, fn_dec, size_dec) if dec_txt else 0.0
    total_w_pt = w_int_pt + w_dec_pt
    x_pt = x_mm * mm
    y_pt = (A4_H_MM - y_mm_from_top) * mm
    a = (align or "left").strip().lower()

    if a == "center":
        start_x = x_pt - total_w_pt / 2.0
    elif a == "right":
        start_x = x_pt - total_w_pt
    else:
        start_x = x_pt

    c.setFillColor(_hex_to_color(base_color))
    c.setFont(fn_int, size_int)
    _maybe_faux_bold(c, base_bold, fn_int, lambda dx, dy: c.drawString(start_x + dx, y_pt + dy, intp))
    if dec_txt:
        c.setFont(fn_dec, size_dec)
        _maybe_faux_bold(c, base_bold, fn_dec, lambda dx, dy: c.drawString(start_x + w_int_pt + dx, y_pt + dy, dec_txt))

    if style.get("underline"):
        x1 = start_x; x2 = start_x + total_w_pt
        c.setLineWidth(0.7); c.line(x1, y_pt - 1.5, x2, y_pt - 1.5)
    if style.get("strike"):
        x1 = start_x; x2 = start_x + total_w_pt
        strike_y = y_pt + size_int * float(style.get("strike_offset", 0.35))
        c.setLineWidth(0.7); c.line(x1, strike_y, x2, strike_y)

def draw_field(field, text, style, x_mm, y_mm_from_top, align="left"):
    is_price = field in {"PROMO", "REG"} and style.get("decimal_scale") is not None
    if is_price:
        _draw_price_with_scaled_decimals(canvas, x_mm, y_mm_from_top, text, style, align)
        return
    _draw_text(canvas, x_mm, y_mm_from_top, text, style, align)


# ---- FITTING HELPERS (wrap/shrink so nothing crosses the midline or margins) ----
def _side_bounds_mm(side: int, margin_mm: float = 2.0) -> tuple[float, float, float]:
    half = A4_W_MM / 2.0
    if side == 0:
        left, right = margin_mm, half - margin_mm
    else:
        left, right = half + margin_mm, A4_W_MM - margin_mm
    return left, right, (right - left)

def _avail_width_pt_for_anchor(side: int, x_mm: float, align: str, max_w_mm: float, margin_mm: float) -> float:
    if side in (0, 1):
        left_mm, right_mm, _ = _side_bounds_mm(side, margin_mm=margin_mm)
    else:
        left_mm, right_mm = margin_mm, A4_W_MM - margin_mm

    cap_pt = (max_w_mm * mm) if (max_w_mm and max_w_mm > 0) else float("inf")
    a = (align or "left").strip().lower()
    if a == "left":
        avail_mm = max(0.0, right_mm - x_mm)
    elif a == "right":
        avail_mm = max(0.0, x_mm - left_mm)
    else:
        avail_mm = max(0.0, 2.0 * min(right_mm - x_mm, x_mm - left_mm))

    return min(avail_mm * mm, cap_pt)


def _wrap_lines_to_width(text: str, font_name: str, size: float, max_w_pt: float, *, max_lines: int = 2):
    """
    Greedy wrap for up to `max_lines` lines.
    - Line 1: fill normally.
    - Last line: keep adding words while they fit; only then trim tail and add "…".
    - If the first word itself doesn't fit, trim that single word + "…".
    """
    from reportlab.pdfbase.pdfmetrics import stringWidth

    def fits(s: str) -> bool:
        return stringWidth(s, font_name, size) <= max_w_pt

    words = (text or "").split()
    if not words:
        return [""]

    # If even the first word doesn't fit, trim that token to fit + ellipsis
    if not fits(words[0]) and max_lines >= 1:
        w = words[0]
        lo, hi = 0, len(w)
        best = "…"
        while lo <= hi:
            mid = (lo + hi) // 2
            cand = (w[:mid] + "…") if mid > 0 else "…"
            if fits(cand):
                best = cand
                lo = mid + 1
            else:
                hi = mid - 1
        return [best]

    lines = []
    cur = words[0]
    i = 1

    # Fill all but the last line
    while i < len(words) and len(lines) + 1 < max_lines:
        nxt = words[i]
        if fits(cur + " " + nxt):
            cur = cur + " " + nxt
            i += 1
        else:
            lines.append(cur)
            cur = nxt
            i += 1

    # Build the last line: add as many words as fit, then ellipsize only if leftover remains
    last = cur
    while i < len(words) and fits(last + " " + words[i]):
        last = last + " " + words[i]
        i += 1

    # If there are leftover words, append ellipsis by trimming tail to fit
    if i < len(words):
        # ensure we end with an ellipsis that fits
        if fits(last + "…"):
            last = last + "…"
        else:
            # trim characters from the end until "…" fits
            base = last.rstrip()
            while base and not fits(base + "…"):
                base = base[:-1]
            last = (base + "…") if base else "…"

    lines.append(last)
    return lines[:max_lines]

def _wrap_lines_strict_no_ellipsis(text: str, font_name: str, size: float, max_w_pt: float, *, max_lines: int = 2):
    """
    Greedy word-wrap with a hard cap on max_lines.
    - Never adds ellipsis.
    - Returns a list of lines if the whole text fits within max_lines, else None.
    """
    from reportlab.pdfbase.pdfmetrics import stringWidth

    def fits(s: str) -> bool:
        return stringWidth(s, font_name, size) <= max_w_pt

    words = (text or "").split()
    if not words:
        return [""]

    # If even the first word doesn't fit at this size, caller must shrink.
    if not fits(words[0]):
        return None

    lines = []
    cur = words[0]

    for w in words[1:]:
        candidate = f"{cur} {w}"
        if fits(candidate):
            cur = candidate
        else:
            lines.append(cur)
            if len(lines) >= max_lines:
                return None  # would exceed line cap
            cur = w

    lines.append(cur)
    if len(lines) > max_lines:
        return None
    return lines


def _draw_text_fitting(c, side: int, x_mm: float, y_mm_from_top: float, text: str, style: dict, align: str, *,
                       max_w_mm: float = 0.0, margin_mm: float = 2.0):
    if "pdfmetrics" not in globals():
        _ensure_reportlab()
    _ensure_unicode_fonts()
    from reportlab.pdfbase import pdfmetrics

    # Price-aware override (same logic as _draw_text)
    try:
        if style and style.get("decimal_scale") is not None:
            raw_txt = "" if text is None else str(text)
            if _split_price_parts(raw_txt):
                _draw_price_with_scaled_decimals(c, x_mm, y_mm_from_top, raw_txt, style, align)
                return
    except Exception:
        pass

    raw = "" if text is None else str(text)
    raw = _sanitize_text(raw)
    shaped = _shape_for_pdf(raw)

    base_family = (style.get("family") or "NotoSans")
    if _contains_arabic(raw):
        _ensure_arabic_font()
        base_family = _ARABIC_FONT_NAME
    bold    = bool(style.get("bold"))
    italic  = bool(style.get("italic"))
    color   = style.get("color", "#000000")
    size    = float(style.get("size", 12))
    leading = float(style.get("leading", 1.0))

    fn = _pdf_font_name(base_family, bold, italic)
    c.setFillColor(_hex_to_color(color))

    x_base = x_mm * mm
    y_base = (A4_H_MM - y_mm_from_top) * mm

    # available width
    if max_w_mm and max_w_mm > 0:
        avail_pt = max(0.0, (max_w_mm - 2.0 * margin_mm) * mm)
    else:
        avail_pt = _avail_width_pt_for_anchor(side, x_mm, align, 0.0, margin_mm)

    a = (align or "left").strip().lower()

    def _wrap_lines(s: str, font_name: str, font_size: float):
        if avail_pt <= 0:
            return [s]
        lines, words = [], s.split()
        if not words:
            return [""]
        cur = words[0]
        for w in words[1:]:
            cand = cur + " " + w
            if pdfmetrics.stringWidth(cand, font_name, font_size) <= avail_pt:
                cur = cand
            else:
                lines.append(cur); cur = w
        if cur:
            lines.append(cur)
        return lines

    cur_size = size
    lines = _wrap_lines(shaped, fn, cur_size)
    while avail_pt > 0 and any(pdfmetrics.stringWidth(L, fn, cur_size) > avail_pt for L in lines) and cur_size > 6.0:
        cur_size -= 0.5
        lines = _wrap_lines(shaped, fn, cur_size)

    c.setFont(fn, cur_size)

    # draw text
    if len(lines) > 1:
        for i, L in enumerate(lines[:2]):
            w_i = pdfmetrics.stringWidth(L, fn, cur_size)
            if a == "center":
                x_line = x_base - (w_i / 2.0)
            elif a == "right":
                x_line = x_base - w_i
            else:
                x_line = (x_base - (avail_pt / 2.0) + margin_mm * mm) if avail_pt > 0 else x_base
            y_line = y_base - i * (cur_size * float(leading))
            c.drawString(x_line, y_line, L)
        w = max(pdfmetrics.stringWidth(L, fn, cur_size) for L in lines[:2])
        x_for_decoration = x_base - (w / 2.0) if a == "center" else (x_base - w if a == "right" else (x_base - (avail_pt / 2.0) + margin_mm * mm if avail_pt > 0 else x_base))
        x_pt = x_for_decoration
        y_pt = y_base
    else:
        w = pdfmetrics.stringWidth(shaped, fn, cur_size)
        if a == "center":
            x_pt = x_base - w / 2.0
        elif a == "right":
            x_pt = x_base - w
        else:
            x_pt = x_base
        y_pt = y_base
        c.drawString(x_pt, y_pt, shaped)

    # decoration
    if style.get("underline"):
        c.setLineWidth(0.7); c.line(x_pt, y_pt - cur_size * 0.12, x_pt + w, y_pt - cur_size * 0.12)
    if style.get("strike"):
        strike_y = y_pt + cur_size * float(style.get("strike_offset", 0.30))
        c.setLineWidth(0.7); c.line(x_pt, strike_y, x_pt + w, strike_y)

    # --- PLU barcode (Fresh): draw bars 6mm under small, pure-digit values ---
    try:
        if FRESH_SECTION_ACTIVE:
            digits_only = raw.isdigit()
            is_plu_len  = 3 <= len(raw) <= 6
            is_small    = cur_size <= 16.0  # avoids prices/headlines
            if digits_only and is_plu_len and is_small:
                from reportlab.graphics.barcode import code128
                bc = code128.Code128(raw, barHeight=6 * mm, barWidth=0.22 * mm)
                bc.drawOn(c, x_pt, y_pt - 4.0 * mm)
    except Exception:
        pass


def _draw_text_2line_shrink_left(c, x_mm: float, y_mm_from_top: float, text: str, style: dict, *,
                                 max_w_mm: float, margin_mm: float = 2.0):
    """
    Draw up to 2 lines, shrinking font as needed to fit inside a width box.
    IMPORTANT: x_mm is treated as the CENTER of the box (not the left edge).
    """

    # Ensure fonts + shapers are ready
    if "pdfmetrics" not in globals():
        _ensure_reportlab()
    _ensure_unicode_fonts()
    from reportlab.pdfbase import pdfmetrics

    # --- sanitize & shape (keeps Arabic correct) ---
    raw    = "" if text is None else str(text)
    raw    = _sanitize_text(raw)
    shaped = _shape_for_pdf(raw)

    # --- style ---
    base_family = style.get("family", "Helvetica")
    if _contains_arabic(raw):
        _ensure_arabic_font()
        base_family = _ARABIC_FONT_NAME
    base_bold   = bool(style.get("bold"))
    base_italic = bool(style.get("italic"))
    color       = style.get("color", "#000000")
    size        = float(style.get("size", 12))
    min_size    = float(style.get("min_size", max(6.0, size * 0.65)))
    leading     = float(style.get("leading", 1.12))

    font_name = _pdf_font_name(base_family, base_bold, base_italic)

    # Faux-bold detector (safe even if you don't have a real Bold face for Arabic)
    emulate_bold = False
    if base_bold:
        try:
            can_bold = font_name in pdfmetrics.getRegisteredFontNames() and (
                "Bold" in font_name or "BoldItalic" in font_name
            )
            if not can_bold and _contains_arabic(raw):
                emulate_bold = True
        except Exception:
            emulate_bold = True

    # --- geometry: treat x as CENTER of the box ---
    x_center = (x_mm * mm)
    if max_w_mm and max_w_mm > 0:
        avail_pt = max(1.0, (max_w_mm * mm) - 2.0 * margin_mm * mm)
    else:
        # fallback width (40% of page), centered on x
        avail_pt = max(1.0, 0.40 * (A4_W_MM * mm) - 2.0 * margin_mm * mm)

    y_pt = (A4_H_MM - y_mm_from_top) * mm

    # --- 2-line greedy wrap with shrink ---
    def wrap2(s: str, fname: str, fsize: float):
        words = (s or "").split()
        if not words:
            return [""]
        lines, cur = [], words[0]
        for w in words[1:]:
            cand = cur + " " + w
            if pdfmetrics.stringWidth(cand, fname, fsize) <= avail_pt:
                cur = cand
            else:
                lines.append(cur)
                cur = w
                if len(lines) >= 2:
                    return lines[:2]
        if cur:
            lines.append(cur)
        return lines[:2]

    cur = size
    lines = wrap2(shaped, font_name, cur)
    while any(pdfmetrics.stringWidth(L, font_name, cur) > avail_pt for L in lines) and cur > min_size:
        cur -= 0.5
        lines = wrap2(shaped, font_name, cur)

    # --- draw centered (with optional faux-bold) ---
    def draw_centered_line(y_line: float, s: str):
        w = pdfmetrics.stringWidth(s, font_name, cur)
        x_left = x_center - (w / 2.0)
        c.setFillColor(_hex_to_color(color))
        c.setFont(font_name, cur)
        c.drawString(x_left, y_line, s)
        if emulate_bold:
            # tiny offset pass to thicken glyphs
            c.drawString(x_left + 0.25, y_line, s)

    for i, line in enumerate(lines):
        s = (line or "").strip()
        y_line = y_pt - (i * cur * leading)
        draw_centered_line(y_line, s)


def _draw_price_fitting(c, side: int, x_mm: float, y_mm_from_top: float, text: str, style: dict, align: str, *,
                        max_w_mm: float = 0.0, margin_mm: float = 2.0):
    if "pdfmetrics" not in globals(): _ensure_reportlab()
    parts = _split_price_parts(text or "")
    if not parts:
        return _draw_text_fitting(c, side, x_mm, y_mm_from_top, text, style, align,
                                  max_w_mm=max_w_mm, margin_mm=margin_mm)

    avail_pt  = _avail_width_pt_for_anchor(side, x_mm, align, max_w_mm, margin_mm)
    base_family = style.get("family", "Helvetica")
    base_bold   = bool(style.get("bold"))
    base_italic = bool(style.get("italic"))
    color       = style.get("color", "#000000")
    size_int    = float(style.get("size", 12))
    min_size    = float(style.get("min_size", max(6.0, size_int * 0.65)))
    dec_scale   = float(style.get("decimal_scale", 0.6))

    fn = _pdf_font_name(base_family, base_bold, base_italic)

    def total_width(sz):
        intp, decp = parts
        dec_txt = f".{decp}" if decp else ".00"
        w_int = pdfmetrics.stringWidth(intp, fn, sz)
        w_dec = pdfmetrics.stringWidth(dec_txt, fn, max(6.0, sz * dec_scale))
        return w_int + w_dec

    cur = size_int
    while cur > min_size and total_width(cur) > avail_pt:
        cur -= 0.5

    st2 = dict(style)
    st2["size"] = max(min_size, cur)
    _draw_price_with_scaled_decimals(c, x_mm, y_mm_from_top, text, st2, align)

# ---- Robust clipping per page half to prevent cross-midline bleed ----
def _clip_to_side(canvas_obj, side: int, margin_mm: float = 2.0):
    """Clip drawing to left(0)/right(1) half of A4 to avoid spillover between slots."""
    _ensure_reportlab()
    width_pt  = A4_W_MM * mm
    height_pt = A4_H_MM * mm
    half_pt   = width_pt / 2.0
    left = (margin_mm * mm) if side == 0 else (half_pt + margin_mm * mm)
    rect_w = (half_pt - 2 * margin_mm * mm)
    path = canvas_obj.beginPath()
    path.rect(left, 0, rect_w, height_pt)
    canvas_obj.clipPath(path, stroke=0, fill=0)

def _coerce_int_keys(d: dict) -> dict:
    out={}
    for k,v in (d or {}).items():
        try: ik=int(k)
        except Exception: ik=k
        out[ik]=_coerce_int_keys(v) if isinstance(v, dict) else v
    return out

def _grid_order_positions(pos: dict) -> List[Tuple[int,int,dict]]:
    rows = sorted([k for k in pos.keys() if isinstance(k,int)])
    result=[]
    for r in rows:
        sides=pos.get(r,{})
        for s in (0,1):
            if isinstance(sides.get(s), dict): result.append((r, s, sides[s]))
    return result

def _infer_layout_mode(positions: dict) -> str:
    """
    Return 'two_col' only if any row has BOTH sides (0 and 1) — i.e. a 2-column (6-slot) grid.
    Otherwise return 'full_width' (e.g., 4-strip or any single-side layout).
    """
    pos = _coerce_int_keys(positions or {})
    for r, sides in pos.items():
        if isinstance(r, int) and isinstance(sides, dict) and (0 in sides) and (1 in sides):
            return "two_col"
    return "full_width"

def _safe_clip_to_layout(canvas_obj, side: int, layout: str, margin_mm: float = 2.0):
    """Only clip for true two-column layouts; no clip for full-width/strip layouts."""
    if layout == "two_col":
        _clip_to_side(canvas_obj, side, margin_mm=margin_mm)


def _value_for_header(rec: Dict[str,str], header: str) -> str:
    h = (header or "").upper()

    # >>> ADD THIS BLOCK <<<
    if FRESH_SECTION_ACTIVE and h in ("REG", "PROMO", "REGULAR_PRICE", "PROMO_PRICE"):
        val = rec.get(h, "")  # numeric text (e.g., "12.50")
        uom = (rec.get("UOM", "") or "").strip()  # already like "/ KG"
        return f"{val} {uom}".strip() if val else ""

    # Legacy
    if h == "BARCODE":     return rec.get("BARCODE","")
    if h == "START_DATE":  return rec.get("START_DATE","")
    if h == "END_DATE":    return rec.get("END_DATE","")
    if h == "BRAND":       return rec.get("BRAND","")
    if h == "ITEM":        return rec.get("ITEM","")
    if h == "REG":         return rec.get("REG","")
    if h == "PROMO":       return rec.get("PROMO","")
    if h == "COOP":        return rec.get("COOP","")
    if h == "SECTION":     return rec.get("SECTION","")
    if h == "BLANK":       return ""

    # Fresh
    if h == "PLU":                 return rec.get("PLU","")
    if h == "ARABIC_DESCRIPTION":  return rec.get("ARABIC_DESCRIPTION","")
    if h == "ENGLISH_DESCRIPTION": return rec.get("ENGLISH_DESCRIPTION","")
    if h == "REGULAR_PRICE":       return rec.get("REGULAR_PRICE","")
    if h == "PROMO_PRICE":         return rec.get("PROMO_PRICE","")
    if h == "UOM":                 return rec.get("UOM","")

    # Fallback
    return rec.get(h, "")


# REPLACE your in-function imports with these two lines
from reportlab.pdfgen.canvas import Canvas as _PDFCanvas
from reportlab.lib.units import mm


def render_page_JSON(out_path: str, tpl: dict, rows: List[Dict[str, str]]) -> str:
    """
    Render labels using the JSON template.

    Strict row filter (inside this function only):
      - Keep a row ONLY if it has a real item name (>=3 chars) AND a numeric price > 0.
      - Header matching is case/space/underscore-insensitive:
          Names: ITEM, ITEM DESCRIPTION, ENGLISH_DESCRIPTION, ARABIC DESCRIPTION, DESCRIPTION, DESC, BRAND
          Prices: PROMO, REG, PROMO PRICE, REG PRICE, PROMOTION PRICE, REGULAR PRICE, PRICE
    """
    # Local imports to avoid global NameError
    from reportlab.pdfgen.canvas import Canvas as _PDFCanvas
    from reportlab.lib.units import mm

    _ensure_reportlab()

    # ---------- helpers (local only) ----------
    import re
    def _is_nonempty_str(v) -> bool:
        try:
            return v is not None and str(v).strip() != ""
        except Exception:
            return False

    def _norm_key(k: str) -> str:
        # Uppercase and remove spaces/underscores/other non-alnums
        return re.sub(r"[^A-Z0-9]", "", str(k).upper())

    # Normalized key variants
    NAME_KEYS_N = [
        "ITEM", "ITEMDESCRIPTION", "ENGLISHDESCRIPTION", "ARABICDESCRIPTION",
        "DESCRIPTION", "DESC", "BRAND"
    ]
    PRICE_KEYS_N = [
        "PROMO", "REG", "PROMOPRICE", "REGPRICE", "PROMOTIONPRICE",
        "REGULARPRICE", "PRICE"
    ]

    def _norm_record(rec: dict) -> dict:
        # Build normalized-key view (does not mutate original)
        out = {}
        for k, v in (rec or {}).items():
            out[_norm_key(k)] = v
        return out

    def _first_nonempty_norm(nrec: dict, keys_norm: list[str]) -> str:
        for nk in keys_norm:
            v = nrec.get(nk)
            if _is_nonempty_str(v):
                return str(v).strip()
        return ""

    def _parse_price(s: str) -> float:
        try:
            t = str(s).strip().replace(",", "")
            m = re.findall(r"-?\d+(?:\.\d+)?", t)
            return float(m[-1]) if m else 0.0
        except Exception:
            return 0.0

    def _row_is_meaningful(rec: dict) -> bool:
        nrec = _norm_record(rec)
        name = _first_nonempty_norm(nrec, NAME_KEYS_N)
        if len(name) < 3:
            return False
        price_txt = _first_nonempty_norm(nrec, PRICE_KEYS_N)
        return _parse_price(price_txt) > 0.0
    # -----------------------------------------

    # Template bits
    positions   = _coerce_int_keys(tpl.get("positions", {}))
    styles      = tpl.get("styles", {}) or {}
    active      = tpl.get("active_headers", {}) or {}
    slots       = _grid_order_positions(positions)
    per_page    = len(slots) if slots else 6
    layout_mode = _infer_layout_mode(positions)

    # Filter rows strictly: no empty, no price-only, no 1–2 char noise
    try:
        data = [r for r in (rows or []) if isinstance(r, dict) and _row_is_meaningful(r)]
    except Exception:
        data = []

    def _page_size_mm() -> Tuple[float, float]:
        ps = (tpl.get("page_size") or "A4").upper()
        if ps == "A5":         return (A5_W_MM, A5_H_MM)
        if ps == "LETTER":     return (LETTER_W_MM, LETTER_H_MM)
        if ps == "LEGAL":      return (LEGAL_W_MM, LEGAL_H_MM)
        return (A4_W_MM, A4_H_MM)

    W_MM, H_MM = _page_size_mm()

    # Slot content check
    def _slot_has_content(rec: dict, headers: Dict[str, dict]) -> bool:
        for hname, pos in headers.items():
            if active and (hname in active) and (not active[hname]):
                continue
            if not pos.get("visible", True):
                continue
            if _is_nonempty_str(_value_for_header(rec, hname)):
                return True
        return False

    pages_to_draw = []
    if per_page > 0 and data:
        for page_start in range(0, len(data), per_page):
            batch = data[page_start:page_start + per_page]
            draw_this_page = False
            for slot_idx, (_row, _side, hdrs) in enumerate(slots):
                if slot_idx >= len(batch):
                    continue
                if _slot_has_content(batch[slot_idx], hdrs):
                    draw_this_page = True
                    break
            if draw_this_page:
                pages_to_draw.append((page_start, batch))

    # No pages left to draw → return without leaving an empty file
    if not pages_to_draw:
        try:
            import os
            if os.path.exists(out_path):
                os.remove(out_path)
        except Exception:
            pass
        return out_path

    c = _PDFCanvas(out_path, pagesize=(W_MM * mm, H_MM * mm))
    c.setAuthor("Price Label Generator")
    c.setTitle(tpl.get("title") or "Labels")


    for page_start, batch in pages_to_draw:
        c.saveState()
        drew_anything = False

        for slot_idx, (_row, _side, headers) in enumerate(slots):
            if slot_idx >= len(batch):
                continue
            rec = batch[slot_idx]
            if not isinstance(rec, dict):
                continue

            # --- STACKED LAYOUT: BRAND -> ITEM/ENGLISH_DESCRIPTION -> PRICE ---
            has_brand = "BRAND" in headers
            has_item  = ("ITEM" in headers) or ("ENGLISH_DESCRIPTION" in headers)

            # Choose a price header present in this slot
            price_key = None
            for pk in ("PROMO", "REG", "PROMO_PRICE", "REGULAR_PRICE", "PRICE"):
                if pk in headers:
                    price_key = pk
                    break
            if price_key is None:
                for hname, pos in headers.items():
                    if _is_price_header(hname, pos):
                        price_key = hname
                        break

            # --- measure helpers (unchanged math) ---
            def _measure_generic_height(pos: dict, text: str, st: dict, align: str) -> Tuple[float, float]:
                from reportlab.pdfbase import pdfmetrics
                raw = _sanitize_text("" if text is None else str(text))
                shaped = _shape_for_pdf(raw)
                base_family = (st.get("family") or "NotoSans")
                if _contains_arabic(raw):
                    _ensure_arabic_font(); base_family = _ARABIC_FONT_NAME
                bold = bool(st.get("bold")); italic = bool(st.get("italic"))
                size = float(st.get("size", 12)); leading = float(st.get("leading", 1.0))
                fn = _pdf_font_name(base_family, bold, italic)
                max_w_mm = float(pos.get("max_w_mm", 0) or 0); margin_mm = float(pos.get("margin_mm", 2.0))
                avail_pt = _avail_width_pt_for_anchor(_side if layout_mode=="two_col" else -1,
                                                      float(pos.get("x", 0.0)), (pos.get("align","left") or "left"),
                                                      max_w_mm, margin_mm)
                def wrap_all(s: str, fsize: float):
                    if avail_pt <= 0: return [s]
                    words = s.split()
                    if not words: return [""]
                    from reportlab.pdfbase import pdfmetrics as pm
                    lines, cur = [], words[0]
                    for w in words[1:]:
                        cand = f"{cur} {w}"
                        if pm.stringWidth(cand, fn, fsize) <= avail_pt: cur = cand
                        else: lines.append(cur); cur = w
                    lines.append(cur); return lines
                cur = size; lines = wrap_all(shaped, cur)
                from reportlab.pdfbase import pdfmetrics as pm
                while any(pm.stringWidth(L, fn, cur) > avail_pt for L in lines) and cur > 6.0:
                    cur -= 0.5; lines = wrap_all(shaped, cur)
                total_pt = max(0.0, len(lines)) * (cur * float(leading))
                return (total_pt / mm, cur)

            def _measure_item_two_line(pos: dict, text: str, st: dict) -> Tuple[float, float]:
                from reportlab.pdfbase import pdfmetrics
                from reportlab.pdfbase.pdfmetrics import stringWidth
                raw = _sanitize_text("" if text is None else str(text))
                shaped = _shape_for_pdf(raw)
                base_family = st.get("family", "Helvetica")
                if _contains_arabic(raw):
                    _ensure_arabic_font(); base_family = _ARABIC_FONT_NAME
                bold = bool(st.get("bold")); italic = bool(st.get("italic"))
                size = float(st.get("size", 12))
                min_size = float(st.get("min_size", max(6.0, size * 0.65)))
                leading  = float(st.get("leading", 1.0))
                font_name = _pdf_font_name(base_family, bold, italic)
                max_w_mm = float(pos.get("max_w_mm", 0) or 0); margin_mm = float(pos.get("margin_mm", 2.0))
                avail_pt = _avail_width_pt_for_anchor(_side if layout_mode=="two_col" else -1,
                                                      float(pos.get("x", 0.0)), (pos.get("align","left") or "left"),
                                                      max_w_mm, margin_mm)
                def wrap2(s: str, fsize: float):
                    words = (s or "").split()
                    if not words: return [""]
                    lines, cur = [], words[0]
                    for w in words[1:]:
                        cand = f"{cur} {w}"
                        if stringWidth(cand, font_name, fsize) <= avail_pt: cur = cand
                        else: lines.append(cur); cur = w
                        if len(lines) >= 2: return lines[:2]
                    lines.append(cur); return lines[:2]
                cur = size; lines = wrap2(shaped, cur)
                while any(pdfmetrics.stringWidth(L, font_name, cur) > avail_pt for L in lines) and cur > min_size:
                    cur -= 0.5; lines = wrap2(shaped, cur)
                total_pt = max(0.0, len(lines)) * (cur * float(leading))
                return (total_pt / mm, cur)

            def _measure_price_height(pos: dict, text: str, st: dict) -> Tuple[float, float]:
                from reportlab.pdfbase import pdfmetrics
                parts = _split_price_parts(text or "")
                base_family = st.get("family", "Helvetica")
                bold = bool(st.get("bold")); italic = bool(st.get("italic"))
                size_int = float(st.get("size", 12))
                min_size = float(st.get("min_size", max(6.0, size_int * 0.65)))
                dec_scale = float(st.get("decimal_scale", 0.6))
                fn = _pdf_font_name(base_family, bold, italic)
                max_w_mm = float(pos.get("max_w_mm", 0) or 0); margin_mm = float(pos.get("margin_mm", 2.0))
                avail_pt = _avail_width_pt_for_anchor(_side if layout_mode=="two_col" else -1,
                                                      float(pos.get("x", 0.0)), (pos.get("align","left") or "left"),
                                                      max_w_mm, margin_mm)
                def total_width(sz):
                    if not parts:
                        return pdfmetrics.stringWidth(str(text or ""), fn, sz)
                    intp, decp = parts
                    dec_txt = f".{decp}" if decp else ".00"
                    w_int = pdfmetrics.stringWidth(intp, fn, sz)
                    w_dec = pdfmetrics.stringWidth(dec_txt, fn, max(6.0, sz * dec_scale))
                    return w_int + w_dec
                cur = size_int
                while cur > min_size and total_width(cur) > avail_pt:
                    cur -= 0.5
                return (cur / mm, cur)
            # --- end measure helpers ---

            if has_brand and has_item and price_key:
                item_key = "ITEM" if "ITEM" in headers else "ENGLISH_DESCRIPTION"

                brand_text = _value_for_header(rec, "BRAND")
                item_text  = _value_for_header(rec, item_key)
                price_val  = _value_for_header(rec, price_key)

                brand_pos = headers["BRAND"]; item_pos = headers[item_key]; price_pos = headers[price_key]
                brand_st  = styles.get("BRAND", {"family":"Helvetica","size":12,"bold":False,"italic":False,"color":"#000000"})
                item_st   = styles.get(item_key, {"family":"Helvetica","size":12,"bold":False,"italic":False,"color":"#000000"})
                price_st  = styles.get(price_key, {"family":"Helvetica","size":12,"bold":False,"italic":False,"color":"#000000"})

                brand_align = (brand_pos.get("align", "left") or "left").strip().lower()
                item_align  = (item_pos.get("align",  "left") or "left").strip().lower()
                price_align = (price_pos.get("align", "left") or "left").strip().lower()

                h_brand_mm, _ = _measure_generic_height(brand_pos, brand_text, brand_st, brand_align)
                h_item_mm,  _ = _measure_item_two_line(item_pos, item_text, item_st)
                h_price_mm, _ = _measure_price_height(price_pos, price_val, price_st)

                gap_brand_item  = max(1.0, float(brand_pos.get("gap_after_mm", 1.2)))
                gap_item_price  = max(1.0, float(item_pos.get("gap_after_mm",   1.6)))

                y_brand = float(brand_pos.get("y", 0.0))
                y_item  = float(item_pos.get("y",  0.0))
                y_price = float(price_pos.get("y", 0.0))

                y_item  = max(y_item,  y_brand + h_brand_mm + gap_brand_item)
                y_price = max(y_price, y_item  + h_item_mm  + gap_item_price)

                _draw_text_fitting(
                    c, _side if layout_mode=="two_col" else -1,
                    float(brand_pos.get("x", 0.0)), y_brand, brand_text, brand_st, brand_align,
                    max_w_mm=float(brand_pos.get("max_w_mm", 0) or 0), margin_mm=float(brand_pos.get("margin_mm", 2.0))
                )
                if not item_pos.get("nowrap", False):
                    _draw_text_2line_shrink_left(
                        c, float(item_pos.get("x", 0.0)), y_item, item_text, item_st,
                        max_w_mm=float(item_pos.get("max_w_mm", 0) or 0), margin_mm=float(item_pos.get("margin_mm", 2.0))
                    )
                else:
                    _draw_text_fitting(
                        c, _side if layout_mode=="two_col" else -1,
                        float(item_pos.get("x", 0.0)), y_item, item_text, item_st, item_align,
                        max_w_mm=float(item_pos.get("max_w_mm", 0) or 0), margin_mm=float(item_pos.get("margin_mm", 2.0))
                    )

                _draw_price_fitting(
                    c, _side if layout_mode=="two_col" else -1,
                    float(price_pos.get("x", 0.0)), y_price, price_val, price_st, price_align,
                    max_w_mm=float(price_pos.get("max_w_mm", 0) or 0), margin_mm=float(price_pos.get("margin_mm", 2.0))
                )

                

                if _is_nonempty_str(brand_text) or _is_nonempty_str(item_text) or _is_nonempty_str(price_val):
                    drew_anything = True

                # Render other visible headers (with PLU barcode beside text)
                for hname, pos in headers.items():
                    if hname in ("BRAND", item_key, price_key):
                        continue
                    if active and (hname in active) and (not active[hname]):
                        continue
                    if not pos.get("visible", True):
                        continue

                    text = _value_for_header(rec, hname)
                    if not _is_nonempty_str(text):
                        continue

                    st = styles.get(hname, {"family":"Helvetica","size":12,"bold":False,"italic":False,"color":"#000000"})
                    align = (pos.get("align", "left") or "left").strip().lower()

                    _draw_text_fitting(
                        c, _side if layout_mode=="two_col" else -1,
                        float(pos.get("x", 0.0)), float(pos.get("y", 0.0)),
                        text, st, align,
                        max_w_mm=float(pos.get("max_w_mm", 0) or 0),
                        margin_mm=float(pos.get("margin_mm", 2.0))
                    )

                    # === PLU barcode (built-in, guaranteed to render) ===
                    # === PLU barcode (built-in; force visible) ===
                    if FRESH_SECTION_ACTIVE and (hname or "").strip().upper() == "PLU":
                        try:
                            from reportlab.graphics.barcode import code128
                            import re

                            raw = "" if text is None else str(text)
                            digits = "".join(re.findall(r"\d", raw))  # keep only 0-9
                            if digits:
                                x = float(pos.get("x", 0.0))
                                y = float(pos.get("y", 0.0))
                                size = float(st.get("size", 10))

                                # draw PLU text
                                c.setFont(_pdf_font_name(st.get("family") or "Helvetica",
                                                         bool(st.get("bold")), bool(st.get("italic"))),
                                          size)
                                c.setFillColor(black)
                                c.drawString(x * mm, y * mm, raw)

                                # draw a tiny debug line (1.5 mm) where barcode baseline will be
                                c.saveState()
                                c.setLineWidth(0.3)
                                c.line(x * mm, (y - 6.5) * mm, (x + 1.5) * mm, (y - 6.5) * mm)

                                # draw Code128 bars directly
                                bc = code128.Code128(digits, barHeight=12 * mm, barWidth=0.40 * mm)
                                bc.drawOn(c, x * mm, (y - 12.0) * mm)   # 12 mm below text baseline
                                c.restoreState()
                        except Exception as e:
                            print("[PDF] PLU barcode render failed:", e)




                    drew_anything = True




            else:
                # Fallback: default rendering path
                for hname, pos in headers.items():
                    if active and (hname in active) and (not active[hname]):
                        continue
                    if not pos.get("visible", True):
                        continue

                    text = _value_for_header(rec, hname)
                    st = styles.get(hname, {"family":"Helvetica","size":12,"bold":False,"italic":False,"color":"#000000"})
                    align = (pos.get("align", "left") or "left").strip().lower()

                    if _is_price_header(hname, pos):
                        _draw_price_fitting(
                            c, _side if layout_mode=="two_col" else -1,
                            float(pos.get("x", 0.0)), float(pos.get("y", 0.0)),
                            text, st, align,
                            max_w_mm=float(pos.get("max_w_mm", 0) or 0),
                            margin_mm=float(pos.get("margin_mm", 2.0))
                        )
                    else:
                        _draw_text_fitting(
                            c, _side if layout_mode=="two_col" else -1,
                            float(pos.get("x", 0.0)), float(pos.get("y", 0.0)),
                            text, st, align,
                            max_w_mm=float(pos.get("max_w_mm", 0) or 0),
                            margin_mm=float(pos.get("margin_mm", 2.0))
                        )

                    if _is_nonempty_str(text):
                        drew_anything = True



        c.restoreState()
        if drew_anything:
            c.showPage()

    c.save()
    return out_path






def open_file(p:str)->None:
    try:
        if sys.platform.startswith("darwin"): subprocess.Popen(["open",p])
        elif os.name=="nt": os.startfile(p)
        else: subprocess.Popen(["xdg-open",p])
    except: pass

# ---------- Hover preview widgets ----------
class HoverImagePopup(QWidget):
    def __init__(self, pix: QPixmap, parent=None):
        super().__init__(parent, Qt.Tool | Qt.FramelessWindowHint | Qt.NoDropShadowWindowHint)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setWindowFlag(Qt.WindowStaysOnTopHint, True)
        self.lbl = QLabel(self)
        self.lbl.setAlignment(Qt.AlignCenter)
        self.lbl.setPixmap(pix)
        self.resize(pix.width(), pix.height())
        self.setStyleSheet("background: rgba(30,30,30,220); border: 1px solid #333;")

    def set_pixmap(self, pix: QPixmap):
        self.lbl.setPixmap(pix)
        self.resize(pix.width(), pix.height())

class HoverPreviewButton(QToolButton):
    def __init__(self, small_pix: QPixmap, large_pix: QPixmap, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._popup: Optional[HoverImagePopup] = None
        self._large_pix = large_pix
        self.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        self.setIcon(QIcon(small_pix))
        self.setIconSize(QSize(small_pix.width(), small_pix.height()))
        self.setMouseTracking(True)
        # --- important: prevent layout over-expansion on Windows ---
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.setMinimumSize(small_pix.size())
        self.setMaximumSize(QSize(small_pix.width()+8, small_pix.height()+36))  # icon + text space   

    def enterEvent(self, event):
        global_pos = self.mapToGlobal(QPoint(self.width(), 0))
        w, h = self._large_pix.width(), self._large_pix.height()
        x, y = _clamp_to_screen(global_pos.x(), global_pos.y() - h - 8, w, h)
        self._popup = HoverImagePopup(self._large_pix)
        self._popup.move(x, y)
        self._popup.show()
        super().enterEvent(event)

    def leaveEvent(self, event):
        if self._popup:
            self._popup.close()
            self._popup = None
        super().leaveEvent(event)

class FlowLayout(QLayout):
    """Like CSS flex-wrap: items go left-to-right and wrap to the next line."""
    def __init__(self, parent=None, margin=0, hspacing=8, vspacing=8):
        super().__init__(parent)
        self._items = []
        self._h = hspacing
        self._v = vspacing
        self.setContentsMargins(margin, margin, margin, margin)

    def addItem(self, item): self._items.append(item)
    def count(self): return len(self._items)
    def itemAt(self, i): return self._items[i] if 0 <= i < len(self._items) else None
    def takeAt(self, i): return self._items.pop(i) if 0 <= i < len(self._items) else None
    def expandingDirections(self): return Qt.Orientations(Qt.Orientation(0))
    def hasHeightForWidth(self): return True
    def heightForWidth(self, w): return self._doLayout(QRect(0, 0, w, 0), test_only=True)
    def setGeometry(self, rect): super().setGeometry(rect); self._doLayout(rect, test_only=False)
    def sizeHint(self): return self.minimumSize()

    def minimumSize(self):
        size = QSize(0, 0)
        for it in self._items:
            size = size.expandedTo(it.minimumSize())
        l, t, r, b = self.getContentsMargins()
        size += QSize(l + r, t + b)
        return size

    def _hspace(self):
        if self._h >= 0: return self._h
        return self.smartSpacing(QStyle.PM_LayoutHorizontalSpacing)

    def _vspace(self):
        if self._v >= 0: return self._v
        return self.smartSpacing(QStyle.PM_LayoutVerticalSpacing)

    def smartSpacing(self, pm):
        parent = self.parent()
        if parent is None:
            return self.spacing()
        if isinstance(parent, QWidget):
            return parent.style().pixelMetric(pm, None, parent)
        return self.spacing()

    def _doLayout(self, rect, *, test_only: bool):
        l, t, r, b = self.getContentsMargins()
        x = rect.x() + l
        y = rect.y() + t
        line_height = 0
        right = rect.right() - r
        hspace = self._hspace()
        vspace = self._vspace()

        for it in self._items:
            hint = it.sizeHint()
            next_x = x + hint.width() + hspace
            if (x > rect.x() + l) and (next_x - hspace > right):
                # wrap
                x = rect.x() + l
                y += line_height + vspace
                next_x = x + hint.width() + hspace
                line_height = 0

            if not test_only:
                it.setGeometry(QRect(QPoint(x, y), hint))

            x = next_x
            line_height = max(line_height, hint.height())

        return y + line_height + b - rect.y()


# ---------- Click-away helper ----------
class ClickAwayCloser(QObject):
    """Close the filter popup when clicking anywhere outside it or pressing Esc."""
    def __init__(self, popup):
        super().__init__(popup)
        self._popup = popup

    def eventFilter(self, obj, event):
        et = event.type()
        if et in (QEvent.MouseButtonPress, QEvent.MouseButtonDblClick):
            # why: Qt6 deprecated globalPos(); prefer globalPosition(), but keep Qt5 fallback
            try:
                gp = event.globalPosition().toPoint()  # Qt6
            except Exception:
                try:
                    gp = event.globalPos()  # Qt5 fallback
                except Exception:
                    return False

            try:
                # compare in widget-local coords to avoid geometry() vs global mismatch
                local_pt = self._popup.mapFromGlobal(gp)
                if self._popup.isVisible() and not self._popup.rect().contains(local_pt):
                    self._popup.close()
                    return False
            except Exception:
                pass

        elif et == QEvent.KeyPress and getattr(event, "key", lambda: None)() == Qt.Key_Escape:
            self._popup.close()
            return True

        return False

class FilterableTable(QTableWidget):
    """
    Table with Excel-like header filters on selected columns.
    Keeps external Search box behavior unchanged (owner renders rows).
    """
    def __init__(self, columns, enable_filters: bool = True, parent=None, *, filterable_columns=None):
        super().__init__(0, len(columns), parent)
        self._all = []                                  # internal rows (optional)
        self._filters = {}                              # {col_name: set(str) | None}
        self._columns = tuple(columns)
        self._filters_enabled = bool(enable_filters)
        self._filterable_cols = set(filterable_columns or self._columns)

        # External rendering hooks (so Search All stays exactly as-is)
        self.get_all_rows = None                        # () -> list[list[object]]
        self.external_refresh = False                   # if True, don't paint rows here
        self.on_filters_changed = None                  # callback to parent

        self.setHorizontalHeaderLabels(self._columns)
        header = self.horizontalHeader()
        header.setSectionsClickable(True)
        header.setSectionResizeMode(QHeaderView.Stretch)
        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.setAlternatingRowColors(True)
        header.sectionClicked.connect(self._open_filter)
        header.sectionDoubleClicked.connect(self._on_header_double_clicked)



        if self._filters_enabled:
            self.setContextMenuPolicy(Qt.CustomContextMenu)
            self.customContextMenuRequested.connect(self._show_context_menu)

        self.on_header_toggle_all = None                # optional hook for CHK column
        header.installEventFilter(self)
        self._update_header_labels()

    # Public API ---------------------------------------------------------------

    def attach_rows(self, rows):
        self._all = [list(r) for r in (rows or [])]
        self._refresh()

    def clear_filters(self):
        self._filters.clear()
        self._update_header_labels()
        if self.external_refresh and self.on_filters_changed:
            self.on_filters_changed()
        else:
            self._refresh()

    # Internal -----------------------------------------------------------------

    def _update_header_labels(self):
        """Append a small arrow to filterable headers; mark if filtered."""
        if not self._filters_enabled:
            return
        for idx, col in enumerate(self._columns):
            if col not in self._filterable_cols:
                continue
            item = self.horizontalHeaderItem(idx)
            if item is None:
                item = QTableWidgetItem(col)
                self.setHorizontalHeaderItem(idx, item)
            active = (col in self._filters) and bool(self._filters[col])
            arrow = " ▼"  # Windows-safe; '▾' can miss in some fonts
            suffix = f"{arrow}*" if active else arrow
            item.setText(f"{col}{suffix}")

    def _refresh(self):
        if self.external_refresh:
            # External mode: owner renders rows; just notify.
            if self.on_filters_changed:
                self.on_filters_changed()
            return

        self.setRowCount(0)

        def visible(row):
            if not self._filters_enabled:
                return True
            for col, allowed in self._filters.items():
                if not allowed:
                    continue
                try:
                    idx = self._columns.index(col)
                except ValueError:
                    continue
                if str(row[idx]) not in allowed:
                    return False
            return True

        for r in self._all:
            if visible(r):
                row_idx = self.rowCount()
                self.insertRow(row_idx)
                for c, v in enumerate(r):
                    item = QTableWidgetItem(str(v))
                    item.setTextAlignment(Qt.AlignCenter)
                    self.setItem(row_idx, c, item)

    def _show_context_menu(self, pos):
        if not self._filters_enabled:
            return
        menu = QMenu(self)
        clear_act = menu.addAction("Clear All Filters")
        clear_act.triggered.connect(self.clear_filters)
        menu.exec(self.viewport().mapToGlobal(pos))

    def _open_filter(self, logical_index):
        if not self._filters_enabled:
            return
        try:
            col = self._columns[logical_index]
        except Exception:
            return

        # Let the app handle CHK header toggle
        # Non-filterable headers do nothing on single-click
        if col not in self._filterable_cols:
            return


        # Build the value universe from rows that pass ALL OTHER active header filters
        rows_source = self._all
        if (not rows_source) and self.get_all_rows:
            try:
                rows_source = self.get_all_rows() or []
            except Exception:
                rows_source = []

        idx_map = {name: i for i, name in enumerate(self._columns)}




        def _row_passes_other_filters(row) -> bool:
            # why: brand list should reflect currently filtered section (and others), not the whole dataset
            for f_col, allowed in (self._filters or {}).items():
                if not allowed:
                    continue
                if f_col == col:
                    continue
                i = idx_map.get(f_col, -1)
                if i < 0:
                    continue
                if str(row[i]) not in allowed:
                    return False
            return True

        uniq = set()
        for r in rows_source:
            try:
                if _row_passes_other_filters(r):
                    uniq.add(str(r[logical_index]))
            except Exception:
                continue
        values_all = sorted(uniq, key=lambda s: (s is None, str(s).lower()))
        selected = set(self._filters.get(col, set())) or set(values_all)

        # Popup dialog (click-away closes); add explicit Apply
        dlg = QDialog(self, Qt.Popup | Qt.FramelessWindowHint)
        dlg.setModal(False)
        root = QVBoxLayout(dlg)
        root.setContentsMargins(10, 8, 10, 10)
        root.setSpacing(6)

        # Top bar
        top = QHBoxLayout()
        top.setContentsMargins(0, 0, 0, 0)
        top.setSpacing(6)
        title = QLabel(f"Filter: {col}")
        title.setObjectName("Small")
        close_btn = QPushButton("✕")
        close_btn.setFixedWidth(26)
        close_btn.clicked.connect(dlg.close)
        top.addWidget(title)
        top.addStretch()
        top.addWidget(close_btn)
        root.addLayout(top)

        # Search
        search = QLineEdit()
        root.addWidget(search)

        # (Select All)
        master_var = QCheckBox("(Select All)")
        root.addWidget(master_var)

        # Values list
        list_frame = QListWidget()
        list_frame.setAlternatingRowColors(True)
        root.addWidget(list_frame)

        # Buttons row: Apply (explicit), keeps instant-apply too
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        apply_btn = QPushButton("Apply")
        btn_row.addWidget(apply_btn)
        root.addLayout(btn_row)

        _populating = {"on": False}

        def _visible_values():
            q = (search.text() or "").lower().strip()
            return [v for v in values_all if not q or q in str(v).lower()]

        def populate():
            _populating["on"] = True
            list_frame.clear()
            for v in _visible_values():
                it = QListWidgetItem(str(v))
                it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
                it.setCheckState(Qt.Checked if str(v) in selected else Qt.Unchecked)
                list_frame.addItem(it)
            master_var.setChecked(
                list_frame.count() > 0 and all(list_frame.item(j).checkState() == Qt.Checked for j in range(list_frame.count()))
            )
            _populating["on"] = False

        def apply_now():
            # Update only visible slice; keep hidden selections intact
            vis = set(_visible_values())
            checked_now = {
                list_frame.item(j).text()
                for j in range(list_frame.count())
                if list_frame.item(j).checkState() == Qt.Checked
            }
            merged = (selected - vis) | checked_now
            if merged == set(values_all):
                self._filters.pop(col, None)
            else:
                self._filters[col] = merged
            self._update_header_labels()
            if self.external_refresh and self.on_filters_changed:
                self.on_filters_changed()
            else:
                self._refresh()

        def on_item_changed(_item: QListWidgetItem):
            if _populating["on"]:
                return
            apply_now()  # instant apply

        def on_master(checked: bool):
            if _populating["on"]:
                return
            _populating["on"] = True
            for j in range(list_frame.count()):
                list_frame.item(j).setCheckState(Qt.Checked if checked else Qt.Unchecked)
            _populating["on"] = False
            apply_now()

        populate()
        search.textChanged.connect(lambda _: populate())
        list_frame.itemChanged.connect(on_item_changed)
        master_var.toggled.connect(on_master)
        apply_btn.clicked.connect(lambda: (apply_now(), dlg.close()))

        # Positioning: prefer ABOVE; fallback to BELOW if not enough space
        try:
            dlg.resize(300, min(380, max(240, list_frame.sizeHintForRow(0) * min(8, max(1, list_frame.count())) + 150)))
        except Exception:
            dlg.resize(300, 320)

        section_x = self.horizontalHeader().sectionViewportPosition(logical_index)
        header_top_global = self.mapToGlobal(QPoint(section_x, 0))
        below_pt = QPoint(header_top_global.x(), header_top_global.y() + self.horizontalHeader().height() + 4)
        above_pt = QPoint(header_top_global.x(), header_top_global.y() - dlg.height() - 6)

        try:
            w, h = dlg.width(), dlg.height()
            # Clamp both candidates to screen bounds
            x_b, y_b = _clamp_to_screen(below_pt.x(), below_pt.y(), w, h)
            x_a, y_a = _clamp_to_screen(above_pt.x(), above_pt.y(), w, h)

            ag = QGuiApplication.primaryScreen().availableGeometry()
            # Enough space above?
            has_room_above = (header_top_global.y() - h - 6) >= (ag.top() + 4)

            if has_room_above:
                dlg.move(x_a, y_a)  # prefer above
            else:
                dlg.move(x_b, y_b)  # fallback below
        except Exception:
            # Safe fallback: above
            dlg.move(self.mapToGlobal(QPoint(section_x, -dlg.height() - 6)))

        # Install click-away closer (global event filter) and auto-remove on close
        try:
            _clickaway = ClickAwayCloser(dlg)
            app = QApplication.instance()
            if app is not None:
                app.installEventFilter(_clickaway)
                try:
                    dlg.finished.connect(lambda _=None, a=app, f=_clickaway: a.removeEventFilter(f))
                except Exception:
                    pass
        except Exception:
            pass

        dlg.exec()

    def _on_header_double_clicked(self, logical_index: int):
        """Double-click on CHK header toggles all visible rows; others keep normal behavior."""
        try:
            col = self._columns[logical_index]
        except Exception:
            return
        if col == "CHK" and self.on_header_toggle_all:
            # why: only double-click on CHK selects/deselects all visible
            self.on_header_toggle_all()
        else:
            # Optional: double-click on filterable columns still opens the filter popup
            if self._filters_enabled and col in self._filterable_cols:
                self._open_filter(logical_index)
                




# ---------- Manual Form ----------
class ManualForm(QWidget):
    def __init__(self, parent, on_add, on_proceed, on_clear_form, on_clear_table):
        super().__init__(parent)
        self.on_add = on_add
        self.on_proceed = on_proceed
        self.on_clear_form = on_clear_form
        self.on_clear_table = on_clear_table

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(24, 8, 24, 8)
        main_layout.setSpacing(10)

        def mk_stack(title: str, min_w: Optional[int] = None, with_clear: bool = False):
            wrap = QWidget()
            v = QVBoxLayout(wrap); v.setContentsMargins(0,0,0,0); v.setSpacing(4)
            lbl = QLabel(title); lbl.setObjectName("FormLabel")
            v.addWidget(lbl)
            if with_clear:
                line_wrap = QWidget()
                h = QHBoxLayout(line_wrap); h.setContentsMargins(0,0,0,0); h.setSpacing(6)
                edit = QLineEdit()
                btn = QPushButton("❌"); btn.setFixedWidth(28); btn.setToolTip("Clear all manual fields")
                btn.clicked.connect(self.on_clear_form)
                h.addWidget(edit, 1); h.addWidget(btn, 0, Qt.AlignRight)
                v.addWidget(line_wrap)
            else:
                edit = QLineEdit()
                v.addWidget(edit)
            edit.setFont(QFont("Arial", 18))
            edit.setMinimumHeight(34)
            if min_w: edit.setMinimumWidth(min_w)
            edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            return wrap, edit
    
        # Dynamic labels for Fresh ON/OFF
        bar_label   = "Barcode / PLU" if FRESH_SECTION_ACTIVE else "Barcode"
        brand_label = "Brand / Arabic Description" if FRESH_SECTION_ACTIVE else "Brand"
        coop_label  = "CO-OP / UOM" if FRESH_SECTION_ACTIVE else "CO-OP Price"

        row1 = QWidget(); row1_h = QHBoxLayout(row1); row1_h.setContentsMargins(0,0,0,0); row1_h.setSpacing(20)
        bar_stack, self.e_bar   = mk_stack(bar_label, min_w=220)
        sd_stack,  self.e_sd    = mk_stack("Start Date (dd.mm.yyyy)", min_w=160)
        ed_stack,  self.e_ed    = mk_stack("End Date (dd.mm.yyyy)",  min_w=160, with_clear=True)
        row1_h.addWidget(bar_stack, 1)
        row1_h.addWidget(sd_stack, 1)
        row1_h.addWidget(ed_stack, 1)
        main_layout.addWidget(row1)

        row2 = QWidget(); row2_h = QHBoxLayout(row2); row2_h.setContentsMargins(0,0,0,0); row2_h.setSpacing(20)
        brand_stack, self.e_brand = mk_stack(brand_label, min_w=240)
        item_stack,  self.e_item  = mk_stack("Item (Description)", min_w=320)
        row2_h.addWidget(brand_stack, 1)
        row2_h.addWidget(item_stack, 2)
        main_layout.addWidget(row2)

        row3 = QWidget(); row3_h = QHBoxLayout(row3); row3_h.setContentsMargins(0,0,0,0); row3_h.setSpacing(20)
        reg_stack,   self.e_reg   = mk_stack("Regular Price",   min_w=150)
        promo_stack, self.e_promo = mk_stack("Promotion Price", min_w=150)
        coop_stack,  self.e_coop  = mk_stack(coop_label,        min_w=150)
        row3_h.addWidget(reg_stack, 1)
        row3_h.addWidget(promo_stack, 1)
        row3_h.addWidget(coop_stack, 1)
        main_layout.addWidget(row3)

        btnrow = QWidget()
        btnrow_layout = QHBoxLayout(btnrow)
        btnrow_layout.setContentsMargins(0, 4, 0, 0)
        clear_table_btn = QPushButton("🧹 Clear Table")
        clear_table_btn.setObjectName("Primary")
        clear_table_btn.clicked.connect(self.on_clear_table)
        add_btn = QPushButton("Add"); add_btn.setObjectName("Primary")
        add_btn.clicked.connect(lambda: self.on_add(self.values()))
        proceed_btn = QPushButton("Proceed →"); proceed_btn.setObjectName("Primary")
        proceed_btn.clicked.connect(self.on_proceed)
        btnrow_layout.addWidget(clear_table_btn)
        btnrow_layout.addWidget(add_btn)
        btnrow_layout.addWidget(proceed_btn)
        main_layout.addWidget(btnrow)

        for e in (self.e_bar, self.e_brand, self.e_item, self.e_sd, self.e_ed, self.e_reg, self.e_promo, self.e_coop):
            e.returnPressed.connect(lambda: self.on_add(self.values()))

        def _setup_uppercase(line_edit: QLineEdit):
            def to_upper(txt: str):
                up = txt.upper()
                if up != txt:
                    pos = line_edit.cursorPosition()
                    line_edit.blockSignals(True)
                    line_edit.setText(up)
                    line_edit.blockSignals(False)
                    line_edit.setCursorPosition(pos)
            line_edit.textEdited.connect(to_upper)

        _setup_uppercase(self.e_bar)
        _setup_uppercase(self.e_brand)
        _setup_uppercase(self.e_item)

    def values(self)->Dict[str,str]:
        g=lambda w:w.text().strip()
        return dict(BARCODE=g(self.e_bar), BRAND=g(self.e_brand), ITEM=g(self.e_item),
                    REG=g(self.e_reg), PROMO=g(self.e_promo),COOP=g(self.e_coop),
                    START_DATE=g(self.e_sd), END_DATE=g(self.e_ed))

    def clear(self):
        for e in (self.e_bar,self.e_brand,self.e_item,self.e_sd,self.e_ed,self.e_reg,self.e_promo,self.e_coop):
            e.setText("")

    def fill(self,r:Dict[str,str]):
        def put(e,v): e.setText(v or "")
        put(self.e_bar, r.get("BARCODE","")); put(self.e_brand, r.get("BRAND","")); put(self.e_item, r.get("ITEM",""))
        put(self.e_reg, r.get("REG","")); put(self.e_promo, r.get("PROMO",""));put(self.e_coop,  r.get("COOP",""))
        put(self.e_sd,  r.get("START_DATE","")); put(self.e_ed,  r.get("END_DATE",""))

# ---------- Header Manager & Template dialogs ----------
class HeaderManagerDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("Headers Manager")
        self.setMinimumSize(500, 300)
        self.cfg = load_headers_cfg()
        self.cur_key = None
        layout = QHBoxLayout(self)
        left = QVBoxLayout()
        layout.addLayout(left)
        left.addWidget(QLabel("Headers", objectName="FormLabel"))
        fbar = QHBoxLayout()
        left.addLayout(fbar)
        add_btn = QPushButton("Add +"); add_btn.setFixedWidth(60); fbar.addWidget(add_btn)
        del_btn = QPushButton("Delete −"); del_btn.setFixedWidth(80); fbar.addWidget(del_btn)
        self.listbox = QListWidget(); left.addWidget(self.listbox)
        self._reload_lb(); self.listbox.itemSelectionChanged.connect(self._on_pick)

        right = QGridLayout()
        layout.addLayout(right)
        right.addWidget(QLabel("Name:"), 0, 0)
        self.name_var = QLineEdit(); right.addWidget(self.name_var, 0, 1)
        self.vis_var = QCheckBox("Visible (show in tables)"); right.addWidget(self.vis_var, 1, 1)
        self.sea_var = QCheckBox("Searchable (used in search)"); right.addWidget(self.sea_var, 2, 1)
        right.addWidget(QLabel("Validation Regex (optional):"), 3, 0)
        self.regex_txt = QTextEdit(); self.regex_txt.setFixedHeight(40); right.addWidget(self.regex_txt, 3, 1)
        right.addWidget(QLabel("Synonyms (one per line):"), 4, 0)
        self.syn_txt = QTextEdit(); self.syn_txt.setFixedHeight(160); right.addWidget(self.syn_txt, 4, 1)
        btns = QHBoxLayout(); right.addLayout(btns, 5, 1)
        save_btn = QPushButton("Save"); btns.addWidget(save_btn)
        add_btn.clicked.connect(self._add_header)
        del_btn.clicked.connect(self._del_header)
        save_btn.clicked.connect(self._save_current)
        if self.listbox.count() > 0: self.listbox.setCurrentRow(0); self._on_pick()

    def _reload_lb(self):
        self.listbox.clear()
        for k in all_headers():
            mark = " (core)" if k in CORE_HEADERS else ""
            self.listbox.addItem(f"{k}{mark}")

    def _on_pick(self):
        items = self.listbox.selectedItems()
        if not items: return
        raw = items[0].text(); key = raw.split(" (core)")[0]
        self.cur_key = key
        item = self.cfg.get(key, {"visible":True,"searchable":True,"synonyms":[]})
        self.name_var.setText(key)
        self.vis_var.setChecked(bool(item.get("visible",True))); self.sea_var.setChecked(bool(item.get("searchable",True)))
        self.regex_txt.setPlainText(item.get("regex",""))
        self.syn_txt.setPlainText("\n".join(item.get("synonyms",[])))
        self.name_var.setReadOnly(key in CORE_HEADERS)

    def _save_current(self):
        if not self.cur_key: return
        new_name = self.name_var.text().strip().upper()
        if self.cur_key in CORE_HEADERS and new_name != self.cur_key:
            QMessageBox.information(self, "Headers", "Core headers cannot be renamed."); return
        if not new_name:
            QMessageBox.information(self, "Headers", "Enter a header name."); return
        syn = [ln.strip() for ln in self.syn_txt.toPlainText().splitlines() if ln.strip()]
        item = {"visible": self.vis_var.isChecked(), "searchable": self.sea_var.isChecked(),
                "regex": self.regex_txt.toPlainText().strip(), "synonyms": syn}
        if new_name != self.cur_key:
            self.cfg[new_name] = self.cfg.pop(self.cur_key, {"visible":True,"searchable":True,"synonyms":[]})
            self.cur_key = new_name
        self.cfg[self.cur_key] = item; save_headers_cfg(self.cfg); self._reload_lb()
        QMessageBox.information(self, "Headers", "Saved.")

    def _add_header(self):
        name, ok = QInputDialog.getText(self, "Add Header", "Enter new header key (e.g., MATERIAL_ID):")
        if not ok or not name: return
        key = name.strip().upper()
        if key in self.cfg:
            QMessageBox.information(self, "Headers","Already exists."); return
        self.cfg[key] = {"visible": False, "searchable": True, "synonyms": []}
        save_headers_cfg(self.cfg); self._reload_lb()

    def _del_header(self):
        items = self.listbox.selectedItems()
        if not items: return
        raw = items[0].text(); key = raw.split(" (core)")[0]
        if key in CORE_HEADERS:
            QMessageBox.information(self, "Headers","Core headers cannot be deleted."); return
        if QMessageBox.question(self, "Delete Header", f"Delete '{key}'?") != QMessageBox.Yes: return
        self.cfg.pop(key, None); save_headers_cfg(self.cfg); self._reload_lb()

class AddTemplateDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("Add Template from JSON")
        self.setMinimumSize(600, 400)
        self.result = None
        layout = QGridLayout(self)
        layout.addWidget(QLabel("Template name:"), 0, 0)
        self.name_var = QLineEdit(); layout.addWidget(self.name_var, 0, 1)
        layout.addWidget(QLabel("Paste JSON:"), 1, 0)
        self.txt = QTextEdit(); layout.addWidget(self.txt, 1, 1)
        btns = QHBoxLayout(); layout.addLayout(btns, 2, 0, 1, 2)
        cancel_btn = QPushButton("Cancel"); btns.addWidget(cancel_btn)
        save_btn = QPushButton("Save"); btns.addWidget(save_btn)
        save_btn.clicked.connect(self._ok)
        cancel_btn.clicked.connect(self._cancel)

    def _ok(self):
        name = self.name_var.text().strip()
        if not name: QMessageBox.information(self, "Add Template", "Please enter a template name."); return
        raw = self.txt.toPlainText().strip()
        if not raw: QMessageBox.information(self, "Add Template", "Please paste JSON first."); return
        try: data = json.loads(raw)
        except Exception as e: QMessageBox.critical(self, "Invalid JSON", str(e)); return
        self.result = (name, data); self.accept()

    def _cancel(self):
        self.result = None; self.reject()

class DeleteTemplateDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("Delete Template")
        self.setMinimumSize(400, 300)
        self.result = None
        layout = QVBoxLayout(self)
        self.lbl = QLabel("Select a template to delete:")
        layout.addWidget(self.lbl)
        self.listbox = QListWidget(); layout.addWidget(self.listbox)
        self._items = _list_template_files()
        for name, _ in self._items: self.listbox.addItem(name)
        btns = QHBoxLayout(); layout.addLayout(btns)
        cancel_btn = QPushButton("Cancel"); btns.addWidget(cancel_btn)
        delete_btn = QPushButton("Delete"); btns.addWidget(delete_btn)
        delete_btn.clicked.connect(self._ok)
        cancel_btn.clicked.connect(self.reject)

    def _ok(self):
        sel = self.listbox.selectedItems()
        if not sel: QMessageBox.information(self, "Delete", "Please select a template."); return
        idx = self.listbox.row(sel[0]); self.result = self._items[idx]; self.accept()

# ---------- App ----------
class App(QMainWindow):
    def _defer_startup_tasks(self):
        """Run non-UI-critical work after the window is visible."""
        try:
            _first_run_seed()
        except Exception:
            pass
        try:
            _ = load_headers_cfg()
            _ = load_excel_sources()
        except Exception:
            pass
    def __init__(self):
        super().__init__()
        self._centered_once = False 
        self.setWindowTitle(APP_TITLE)
        self.central = QWidget()
        self.setCentralWidget(self.central)
        self.main_layout = QVBoxLayout(self.central)
        self._preferred_size = "small"; self._auto_shrunk = False
        self._user_overrides_size = False; self._custom_geom = None
        self._resizing_programmatically = False
        self.connected = None
        self.preview_rows = []; self.preview_qty = []
        self.last_mapping = {}
        self.selected_template = None
        self.selected_template_name = None
        self.staged_rows = []; self.staged_qty = []
        self._search_enter_armed = False; self._manual_enter_armed = False
        self._multi_mode_active = False
        self._multi_found_queue = []
        self._multi_found_qty = []
        self._multi_unfound_tokens = []
        self._multi_index = 0
        self._last_autofill_key = None
        self._just_pasted_search = False
        self._paste_items = []
        self._paste_panel = None
        self._editing_from_stage = False
        self.settings = load_settings()
        self._ui_state = load_ui_state()
        self._strict_manual_on = True 
        self._excel_lookup_var = ""
        self._excel_lookup_popup = None
        self._excel_suggest_btn = None
        self._excel_lookup_items = []
        self._excel_lookup_sel = 0
        self._manual_search_var = ""
        self._current_gen_source = None
        self._excel_checked_keys = set()
        self._excel_row_keys = []
        self._excel_iid_to_index = {}
        self._excel_search_var = ""
        self._excel_header_check_state = Qt.Unchecked
        self._build_home()

        # Wire any matching button to commit_manual_form (tries common names)
        try:
            btn = _safe_widget(self, "btnManualSave") or _safe_widget(self, "manual_save_btn") \
                  or _safe_widget(self, "create_manual_btn") or _safe_widget(self, "btnCreateManual")
            if btn:
                try:
                    btn.clicked.disconnect()
                except Exception:
                    pass
                btn.clicked.connect(self.commit_manual_form)
        except Exception:
            pass

        # Keyboard shortcut: Ctrl+Shift+D to open "pick by date"
        try:
            sc = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Shift+D"), self)
            sc.activated.connect(self.pick_excel_by_date_and_open)
        except Exception:
            pass

       
        except Exception:
            pass

        # Shortcut: Ctrl+Shift+U to run UOM migration
        try:
            sc_uom = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Shift+U"), self)
            sc_uom.activated.connect(self.run_uom_slash_migration)
        except Exception:
            pass

        # Try to guard delete buttons if present in UI
        try:
            self._guard_header_delete_button()
        except Exception:
            pass
        try:
            self._guard_template_delete_button()
        except Exception:
            pass

    def showEvent(self, event):  # center only on first show
            super().showEvent(event)
            if self._centered_once:
                return
            self._centered_once = True
            try:
                qr = self.frameGeometry()
                cp = QGuiApplication.primaryScreen().availableGeometry().center()
                qr.moveCenter(cp)
                self.move(qr.topLeft())
            except Exception:
                pass        



    def commit_manual_form(self):
        """
        Safe manual-save handler.
        Reads common widgets if they exist, builds a normalized record, and upserts.
        Uses _safe_widget(...) so missing fields won't crash.
        """
        def _t(name: str) -> str:
            w = _safe_widget(self, name)
            try:
                return w.text().strip() if w else ""
            except Exception:
                return ""

        record = build_manual_record(
            barcode=_t("barcode_edit"),
            brand=_t("brand_edit"),
            item=_t("item_edit"),
            reg=_t("reg_edit"),
            promo=_t("promo_edit"),
            start_date=_t("start_date_edit"),
            end_date=_t("end_date_edit"),
            section=_t("section_edit"),
            coop=_t("coop_edit"),
            plu=_t("plu_edit"),
            arabic_description=_t("arabic_desc_edit"),
            english_description=_t("english_desc_edit"),
            regular_price=_t("regular_price_edit"),
            promo_price=_t("promo_price_edit"),
            uom=_t("uom_edit"),
            source_file="",
            source_sheet=""
        )
        upsert_db_rows([record])
        try:
            QMessageBox.information(self, "Saved", "Manual price saved.")
        except Exception:
            pass
    

    def resizeEvent(self, event: QResizeEvent):
        if self._resizing_programmatically:
            return
        geom = self.geometry()
        w, h = geom.width(), geom.height()
        SMALL_W, SMALL_H = SMALL_GEOM[0], SMALL_GEOM[1]
        LARGE_W, LARGE_H = LARGE_GEOM[0], LARGE_GEOM[1]
        if (w, h) != (SMALL_W, SMALL_H) and (w, h) != (LARGE_W, LARGE_H):
            self._user_overrides_size = True
            self._custom_geom = f"{w}x{h}+{geom.x()}+{geom.y()}"
            self._preferred_size = "custom"           # ← add
        super().resizeEvent(event)

    # inside class App
    def changeEvent(self, event):
        if event.type() == QEvent.ActivationChange:
            if self.isActiveWindow():
                self._apply_preferred_size()
                self._auto_shrunk = False
            else:
                self.set_small()
        super().changeEvent(event)
    



    def moveEvent(self, event):
        if not self._resizing_programmatically:
            self._user_overrides_size = True
            self._custom_geom = f"{self.width()}x{self.height()}+{self.x()}+{self.y()}"
            self._preferred_size = "custom"
        super().moveEvent(event)

    def _on_app_state(self, state):
        if state == Qt.ApplicationInactive:
            self.set_small()
        elif state == Qt.ApplicationActive:
            self._apply_preferred_size()
            self._auto_shrunk = False


    def _resize_keep_center(self, width: int, height: int):
        g = self.geometry()
        cx = g.x() + g.width() // 2
        cy = g.y() + g.height() // 2
        nx = int(cx - width / 2)
        ny = int(cy - height / 2)
        nx, ny = _clamp_to_screen(nx, ny, width, height)
        self._resizing_programmatically = True
        self.setGeometry(nx, ny, width, height)
        QTimer.singleShot(50, lambda: setattr(self, "_resizing_programmatically", False))
        



    

    def set_small(self):
        self._resize_keep_center(SMALL_GEOM[0], SMALL_GEOM[1])

    def set_large(self):
        self._resize_keep_center(LARGE_GEOM[0], LARGE_GEOM[1])



    def _apply_preferred_size(self):
        if self._preferred_size == "large": self.set_large()
        elif self._preferred_size == "small": self.set_small()
        elif self._preferred_size == "custom" and self._custom_geom:
            self._resizing_programmatically = True
            parts = self._custom_geom.split("+")
            wh = parts[0].split("x")
            self.resize(int(wh[0]), int(wh[1]))
            self.move(int(parts[1]), int(parts[2]))
            QTimer.singleShot(50, lambda: setattr(self, "_resizing_programmatically", False))

    def clear(self):
        # Hide floating Excel suggestion before tearing down
        self._hide_excel_popup()

        while self.main_layout.count():
            item = self.main_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        # Proactively drop refs to widgets that are now deleted
        for attr in (
            "templates_frame", "templates_scroll", "tree", "stage", "hits",
            "mform", "_excel_lookup_edit", "_excel_suggest_btn",
        ):
            if hasattr(self, attr):
                setattr(self, attr, None)

        # Stop deferred batch timer if active (prevents late ticks touching dead tables)
        if hasattr(self, "_excel_batch_timer") and alive(getattr(self, "_excel_batch_timer")):
            try:
                self._excel_batch_timer.stop()
            except Exception:
                pass



    def _ask_password(self, prompt="Enter password"):
        return QInputDialog.getText(self, "Security", prompt, QLineEdit.Password)[0]

    def _ensure_password_set(self)->bool:
        if self.settings.get("password_hash"): return True
        p1 = self._ask_password("Create a password (used to unlock settings)")
        if not p1: return False
        p2 = self._ask_password("Re-enter password to confirm")
        if not p2 or p2 != p1:
            QMessageBox.critical(self, "Security","Passwords do not match."); return False
        self.settings["password_hash"] = sha(p1); self.settings["locked"]=True; self.settings["failed_attempts"]=0
        save_settings(self.settings); return True

    def _check_privileged(self, action_name: str)->bool:
        if self.settings.get("failed_attempts",0) >= 3 and self.settings.get("locked",True):
            p = self._ask_password("Locked (3 failed attempts). Enter MASTER password to enable.")
            if p and p == MASTER_PASSWORD:
                self.settings["failed_attempts"]=0; save_settings(self.settings)
                QMessageBox.information(self, "Security","Unlocked by master. Please set/enter your password next time.")
            else:
                QMessageBox.critical(self, "Security","Still locked."); return False
        if not self.settings.get("locked",True): return True
        if not self._ensure_password_set(): return False
        p = self._ask_password(f"{action_name} is locked. Enter password.")
        if not p: return False
        ok = (sha(p)==self.settings.get("password_hash")) or (p==MASTER_PASSWORD)
        if ok:
            self.settings["locked"]=False; self.settings["failed_attempts"]=0; save_settings(self.settings)
            QMessageBox.information(self, "Security","Unlocked."); return True
        else:
            self.settings["failed_attempts"]=self.settings.get("failed_attempts",0)+1; save_settings(self.settings)
            left = max(0, 3 - self.settings["failed_attempts"])
            QMessageBox.critical(self, "Security", f"Wrong password. Attempts left: {left}"); return False

    def _toggle_lock(self):
        self.settings["locked"]=not self.settings.get("locked",True); save_settings(self.settings); self._update_lock_btn()

    def _update_lock_btn(self):
        if hasattr(self, "lock_btn"):
            self.lock_btn.setText("🔒 Locked" if self.settings.get("locked",True) else "🔓 Unlocked")

    # === [CHUNK 8 — A] Paste INSIDE class App(QMainWindow):
    def _is_builtin_header_ui(self, name: str) -> bool:
        try:
            return (name or "").strip() in _BUILTIN_HEADERS
        except Exception:
            return False

    def delete_header_ui(self, header_name_getter: Callable[[], str]):
        """
        UI-safe delete for headers:
          - Built-in: ask master; otherwise allowed.
        """
        name = (header_name_getter() or "").strip()
        if not name:
            QMessageBox.information(self, "Delete header", "Select a header first.")
            return
        if self._is_builtin_header_ui(name):
            pw, ok = QInputDialog.getText(self, "Master required",
                                          f"'{name}' is protected. Enter master password:",
                                          QLineEdit.Password)
            if not ok or not pw:
                return
            if delete_header(name, master_password=pw):
                QMessageBox.information(self, "Deleted", f"Header '{name}' deleted.")
            else:
                QMessageBox.warning(self, "Blocked", "Delete refused (wrong password or protected).")
            return
        # user-added header
        if delete_header(name, master_password=None):
            QMessageBox.information(self, "Deleted", f"Header '{name}' deleted.")
        else:
            QMessageBox.warning(self, "Failed", "Could not delete header.")

    def delete_template_ui(self, template_name_getter: Callable[[], str]):
        """
        UI-safe delete for templates:
          - Bundled templates: ask master; user-added allowed.
        """
        nm = (template_name_getter() or "").strip()
        if not nm:
            QMessageBox.information(self, "Delete template", "Select a template first.")
            return
        if is_bundled_template(nm):
            pw, ok = QInputDialog.getText(self, "Master required",
                                          f"'{nm}' is bundled. Enter master password:",
                                          QLineEdit.Password)
            if not ok or not pw:
                return
            okdel = delete_template_file(nm, master_password=pw)
        else:
            okdel = delete_template_file(nm, master_password=None)
        if okdel:
            QMessageBox.information(self, "Deleted", f"Template '{nm}' deleted.")
        else:
            QMessageBox.warning(self, "Blocked", "Delete refused (wrong password or protected).")

    def _guard_header_delete_button(self):
        """
        If a 'delete header' button exists, disable it when a built-in header is selected.
        Expects a list widget/tree with headers; tries common ids.
        """
        try_list = (_safe_widget(self, "headers_list") or
                    _safe_widget(self, "headers_tree") or
                    _safe_widget(self, "headersView"))
        btn = _safe_widget(self, "btnDeleteHeader") or _safe_widget(self, "delete_header_btn")
        if not (try_list and btn):
            return
        def _sel_name():
            try:
                it = try_list.currentItem() if hasattr(try_list, "currentItem") else None
                if it:
                    return it.text().split("\t")[0].strip()
                # QTreeWidget?
                if hasattr(try_list, "currentItem"):
                    it = try_list.currentItem()
                    if it:
                        return it.text(0).strip()
            except Exception:
                return ""
            return ""
        # toggle enabled
        try:
            name = _sel_name()
            btn.setEnabled(not self._is_builtin_header_ui(name))
        except Exception:
            pass
        # connect delete action to guarded version
        try:
            btn.clicked.disconnect()
        except Exception:
            pass
        btn.clicked.connect(lambda: self.delete_header_ui(_sel_name))
        # live update on selection change
        try:
            if hasattr(try_list, "itemSelectionChanged"):
                try_list.itemSelectionChanged.connect(self._guard_header_delete_button)
        except Exception:
            pass

    def _guard_template_delete_button(self):
        """
        If a 'delete template' button exists, ensure delete is guarded and state reflected.
        """
        tpl_list = _safe_widget(self, "templates_list") or _safe_widget(self, "templatesView")
        btn = _safe_widget(self, "btnDeleteTemplate") or _safe_widget(self, "delete_template_btn")
        if not (tpl_list and btn):
            return
        def _sel_tpl():
            try:
                it = tpl_list.currentItem() if hasattr(tpl_list, "currentItem") else None
                if it:
                    # If your list stores name separately, adjust here
                    return it.text().split("\t")[0].strip()
            except Exception:
                return ""
            return ""
        # set enabled based on bundled status
        try:
            nm = _sel_tpl()
            btn.setEnabled(not is_bundled_template(nm))
        except Exception:
            pass
        # connect guarded delete
        try:
            btn.clicked.disconnect()
        except Exception:
            pass
        btn.clicked.connect(lambda: self.delete_template_ui(_sel_tpl))
        # live update
        try:
            if hasattr(tpl_list, "itemSelectionChanged"):
                tpl_list.itemSelectionChanged.connect(self._guard_template_delete_button)
        except Exception:
            pass
        

    def run_uom_slash_migration(self):
        """
        One-click migration: add '/ ' to UOM for existing Fresh rows.
        Shows result to the user.
        """
        try:
            count = migrate_add_slash_to_uom_for_fresh_rows(dry_run=False)
            QMessageBox.information(self, "Migration complete",
                                    f"Updated UOM with '/ ' for {count} Fresh row(s).")
        except Exception as e:
            QMessageBox.warning(self, "Migration failed", f"Error: {e}")
        
        




    # --- replace your existing _build_home with this version (same UI, adds one line at end) ---
    def _build_home(self):
        # --- session resets when returning to MAIN screen ---
        # Reset Excel quick-import selections so old checkboxes don't "stick"
        try:
            self._excel_checked_keys = set()
        except Exception:
            pass
        try:
            self._excel_row_keys = []
        except Exception:
            pass
        try:
            self._excel_header_check_state = Qt.Unchecked
        except Exception:
            pass

        # Reset Manual staged data ONLY when coming back to MAIN
        try:
            self.staged_rows = []
            self.staged_qty = []
        except Exception:
            pass
        # --- end resets ---

        self.clear(); self._preferred_size = "small"; self.set_small()

        header = QFrame(objectName="Card"); header_layout = QVBoxLayout(header); self.main_layout.addWidget(header)
        title = QLabel("Create Price Labels"); title.setObjectName("Title"); header_layout.addWidget(title)

        card = QFrame(objectName="Card"); card_layout = QVBoxLayout(card); self.main_layout.addWidget(card)
        card_layout.setAlignment(Qt.AlignCenter)

        center_wrap = QVBoxLayout()
        center_wrap.setAlignment(Qt.AlignHCenter)
        card_layout.addStretch(); card_layout.addLayout(center_wrap); card_layout.addStretch()

        search_label = QLabel("Quick Import (type a previously connected Excel name):")
        search_label.setObjectName("Small"); search_label.setAlignment(Qt.AlignHCenter)
        center_wrap.addWidget(search_label, alignment=Qt.AlignHCenter)

        # Row: Fresh Section toggle (left) + Excel search box (right)
        search_row = QWidget()
        search_row_h = QHBoxLayout(search_row)
        search_row_h.setContentsMargins(0, 0, 0, 0)
        search_row_h.setSpacing(10)

        # Fresh toggle
        self.fresh_btn = QPushButton()
        self.fresh_btn.setObjectName("FreshToggle")
        is_on = FRESH_SECTION_ACTIVE
        self.fresh_btn.setText("Fresh Section: ON" if is_on else "Fresh Section: OFF")
        self.fresh_btn.setProperty("active", is_on)
        self.fresh_btn.setFixedHeight(40)
        self.fresh_btn.setMinimumWidth(150)
        self.fresh_btn.clicked.connect(self._toggle_fresh_section)
        self.fresh_btn.style().unpolish(self.fresh_btn); self.fresh_btn.style().polish(self.fresh_btn)
        search_row_h.addWidget(self.fresh_btn, 0, Qt.AlignLeft)
        


        # Excel search box
        self._excel_lookup_edit = QLineEdit()
        self._excel_lookup_edit.setFont(QFont("Arial", 15))
        self._excel_lookup_edit.setMinimumHeight(40)
        self._excel_lookup_edit.setFixedWidth(300)
        self._excel_lookup_edit.installEventFilter(self)
        search_row_h.addWidget(self._excel_lookup_edit, 0, Qt.AlignLeft)

        # Preserve the original Enter handler so PLUS can fall back to it
        if not hasattr(App, "__orig__on_excel_lookup_enter"):
            App.__orig__on_excel_lookup_enter = App._on_excel_lookup_enter

        # PATCH: (re)wire signals once per *widget instance*
        # The flag is reset in clear(), so this will run on each fresh build.
        if not getattr(self, "_excel_lookup_wired", False):
            try:
                self._excel_lookup_edit.textChanged.connect(self._on_excel_lookup_typing)
            except (TypeError, RuntimeError):
                pass
            try:
                self._excel_lookup_edit.returnPressed.connect(self._on_excel_date_enter)
            except (TypeError, RuntimeError):
                pass
            self._excel_lookup_wired = False


        # ▼ Recent files button
        self._excel_recent_btn = QToolButton()
        self._excel_recent_btn.setText("▼")              # simple/portable glyph
        self._excel_recent_btn.setFixedWidth(28)
        self._excel_recent_btn.setToolTip("Recent files (anywhere)")
        self._excel_recent_btn.clicked.connect(self._show_recent_downloads_menu)
        search_row_h.addWidget(self._excel_recent_btn, 0, Qt.AlignLeft)


        # Put the row in the center column
        center_wrap.addWidget(search_row, alignment=Qt.AlignHCenter)

        # Buttons: Connect + Manual (unchanged)
        btn_col = QVBoxLayout()
        btn_col.setAlignment(Qt.AlignHCenter)

        connect_btn = QPushButton("Connect to Excel File")
        connect_btn.setObjectName("Primary")
        manual_btn  = QPushButton("Create Price (Manual)")
        manual_btn.setObjectName("Primary")

        connect_btn.setFixedSize(220, 44)
        manual_btn.setFixedSize(220, 44)

        btn_col.addWidget(connect_btn, alignment=Qt.AlignHCenter)
        btn_col.addWidget(manual_btn,  alignment=Qt.AlignHCenter)

        center_wrap.addLayout(btn_col)

        toolbar = QFrame(objectName="Card"); toolbar_layout = QHBoxLayout(toolbar); self.main_layout.addWidget(toolbar)
        add_tpl_btn = QPushButton("+"); add_tpl_btn.setFixedWidth(40); add_tpl_btn.clicked.connect(self._priv_add_template); toolbar_layout.addWidget(add_tpl_btn)
        del_tpl_btn = QPushButton("−"); del_tpl_btn.setFixedWidth(40); del_tpl_btn.clicked.connect(self._priv_delete_template); toolbar_layout.addWidget(del_tpl_btn)
        center = QFrame(objectName="Card"); center_layout = QHBoxLayout(center); toolbar_layout.addWidget(center)
        headers_btn = QPushButton("Headers…"); headers_btn.setFixedWidth(100); headers_btn.clicked.connect(self._priv_headers_mgr); center_layout.addWidget(headers_btn)
        self.lock_btn = QPushButton(""); self.lock_btn.setFixedWidth(120); self.lock_btn.clicked.connect(self._toggle_lock); center_layout.addWidget(self.lock_btn)
        self._update_lock_btn()

        self.status = QLabel("Click your Excel then press Connect, or type a saved file name above.")
        self.status.setObjectName("Small"); self.main_layout.addWidget(self.status)

        # Wire as before
        connect_btn.clicked.connect(self._on_connect)
        manual_btn.clicked.connect(self._build_manual)

        # NEW: defer any heavy startup work until after the window is visible
        QTimer.singleShot(0, self._defer_startup_tasks)



    def _tune_excel_column_widths(self):
        """Make price/date/Q/COOP compact and let the description column breathe (Excel table)."""
        if not hasattr(self, "tree") or self.tree is None:
            return

        if getattr(self, "_excel_widths_tuned", False):
            try:
                self._excel_apply_qty_buttons()
            except Exception:
                pass
            return

        header = self.tree.horizontalHeader()
        cols = getattr(self.tree, "_columns", ()) or ()

        def idx(name: str) -> int:
            try:
                return cols.index(name)
            except ValueError:
                return -1

        header.setSectionResizeMode(QHeaderView.Stretch)

        for name, w in (
            ("CHK", 28),
            ("Q", 140),
            ("REG", 64), ("PROMO", 64),
            ("REGULAR_PRICE", 64), ("PROMO_PRICE", 64),
            ("COOP", 60),
            ("START", 78), ("END", 78),
            ("BARCODE", 110),
            ("BRAND", 110),
            ("SECTION", 100),
            ("PLU", 88),
            ("ARABIC_DESCRIPTION", 140),
        ):
            i = idx(name)
            if i >= 0:
                header.setSectionResizeMode(i, QHeaderView.Interactive)
                self.tree.setColumnWidth(i, w)

        for big in ("ITEM", "ENGLISH_DESCRIPTION"):
            i = idx(big)
            if i >= 0:
                header.setSectionResizeMode(i, QHeaderView.Stretch)
                break

        q_i = idx("Q")
        if q_i >= 0:
            try:
                from PyQt5.QtWidgets import QStyledItemDelegate
            except Exception:
                pass
            else:
                class _BigQtyDelegate(QStyledItemDelegate):
                    def createEditor(self, parent, option, index):
                        w = super().createEditor(parent, option, index)
                        try:
                            f = w.font()
                            f.setPointSize(18); f.setBold(True)
                            w.setFont(f)
                            if hasattr(w, "setAlignment"): w.setAlignment(Qt.AlignCenter)
                            if hasattr(w, "setFixedHeight"): w.setFixedHeight(36)
                        except Exception:
                            pass
                        return w
                self._excel_qty_delegate = _BigQtyDelegate(self.tree)
                self.tree.setItemDelegateForColumn(q_i, self._excel_qty_delegate)

        self._excel_widths_tuned = True

        try:
            self._excel_apply_qty_buttons()
            for r in range(self.tree.rowCount()):
                if self.tree.rowHeight(r) < 34:
                    self.tree.setRowHeight(r, 34)
        except Exception:
            pass


    def _excel_apply_qty_buttons(self):
        """
        Replace 'Q' (or 'QTY') cells in the Excel table with +/- widgets.
        Persist values in self._excel_qty_override keyed by BARCODE+ITEM.
        """
        if not hasattr(self, "tree") or self.tree is None:
            return

        cols = getattr(self.tree, "_columns", ()) or ()
        if not cols:
            return

        q_col = None
        for name in ("Q", "QTY"):
            if name in cols:
                q_col = cols.index(name); break
        if q_col is None:
            return

        bc_col = cols.index("BARCODE") if "BARCODE" in cols else None
        item_col = cols.index("ITEM") if "ITEM" in cols else None

        if not hasattr(self, "_excel_qty_override"):
            self._excel_qty_override = {}

        def _row_key(r: int) -> str:
            bc = self.tree.item(r, bc_col).text() if (bc_col is not None and self.tree.item(r, bc_col)) else ""
            it = self.tree.item(r, item_col).text() if (item_col is not None and self.tree.item(r, item_col)) else ""
            return f"{bc}||{it}"

        for r in range(self.tree.rowCount()):
            try:
                if self.tree.rowHeight(r) < 34:
                    self.tree.setRowHeight(r, 34)
            except Exception:
                pass

            key = _row_key(r)
            cell = self.tree.item(r, q_col)

            cur = None
            if key in self._excel_qty_override:
                cur = self._excel_qty_override[key]
            elif cell:
                try:
                    cur = int(cell.text())
                except Exception:
                    cur = None
            if cur is None:
                cur = 1

            def _on_change(new_v: int, rk=key, rr=r, cc=q_col):
                self._excel_qty_override[rk] = new_v
                base = self.tree.item(rr, cc)
                if base:
                    base.setText(str(new_v))
                try:
                    idx = self._excel_iid_to_index.get(rr)
                    if idx is not None and 0 <= idx < len(self.preview_qty):
                        self.preview_qty[idx] = new_v
                except Exception:
                    pass

            w = self._make_qty_widget(cur, _on_change)
            self.tree.setCellWidget(r, q_col, w)


    def _tune_stage_column_widths(self):
        """Compact qty/price/date/coop and let ITEM breathe on the Manual 'Your List' table."""
        if not hasattr(self, "stage") or self.stage is None:
            return
        if getattr(self, "_stage_widths_tuned", False):
            return

        header = self.stage.horizontalHeader()
        cols = tuple(getattr(self.stage, "_columns", ()) or ())

        def idx(name: str) -> int:
            try:
                return cols.index(name)
            except ValueError:
                return -1

        header.setSectionResizeMode(QHeaderView.Stretch)

        for name, w in (
            ("QTY", 140),
            ("REG", 70), ("PROMO", 70), ("COOP", 70),
            ("START", 86), ("END", 86),
            ("BARCODE", 120),
            ("BRAND", 120),
        ):
            i = idx(name)
            if i >= 0:
                header.setSectionResizeMode(i, QHeaderView.Interactive)
                self.stage.setColumnWidth(i, w)

        i_item = idx("ITEM")
        if i_item >= 0:
            header.setSectionResizeMode(i_item, QHeaderView.Stretch)

        try:
            for r in range(self.stage.rowCount()):
                if self.stage.rowHeight(r) < 34:
                    self.stage.setRowHeight(r, 34)
        except Exception:
            pass

        self._stage_widths_tuned = True

    
    

    def eventFilter(self, obj, event):
        excel_edit = _safe_widget(self, "_excel_lookup_edit")
        if obj is excel_edit:
            et = event.type()
            if et == QEvent.FocusIn:
                self._on_excel_lookup_typing(excel_edit.text() if excel_edit else "")
            elif et in (QEvent.FocusOut, QEvent.Hide, QEvent.WindowDeactivate):
                self._hide_excel_popup()
        return super().eventFilter(obj, event)    



    def _hide_excel_popup(self):
        if self._excel_lookup_popup:
            self._excel_lookup_popup.close()
        self._excel_lookup_popup = None
        self._excel_suggest_btn = None
        self._excel_lookup_items = []
        self._excel_lookup_sel = 0

    def _toggle_fresh_section(self):
        """Toggle Fresh mode + persist to ui_state.json."""
        global FRESH_SECTION_ACTIVE
        FRESH_SECTION_ACTIVE = not FRESH_SECTION_ACTIVE

        # update button ui
        is_on = FRESH_SECTION_ACTIVE
        if hasattr(self, "fresh_btn") and _is_alive(self.fresh_btn):
            self.fresh_btn.setText("Fresh Section: ON" if is_on else "Fresh Section: OFF")
            self.fresh_btn.setProperty("active", is_on)
            self.fresh_btn.style().unpolish(self.fresh_btn)
            self.fresh_btn.style().polish(self.fresh_btn)

        # status hint
        if hasattr(self, "status") and _is_alive(self.status):
            hint = " (Fresh ON)" if is_on else ""
            self.status.setText(
                f"Click your Excel then press Connect, or type a saved file name above.{hint}"
            )

        # persist
        try:
            self._ui_state = getattr(self, "_ui_state", load_ui_state())
            self._ui_state["fresh_on"] = bool(FRESH_SECTION_ACTIVE)
            save_ui_state(self._ui_state)
        except Exception:
            pass

    def _toggle_strict_manual(self):
        """Flip 'Strict Manual' mode and persist to ui_state.json."""
        self._strict_manual_on = not bool(getattr(self, "_strict_manual_on", True))
        # persist
        try:
            self._ui_state = getattr(self, "_ui_state", load_ui_state())
            self._ui_state["strict_manual_on"] = bool(self._strict_manual_on)
            save_ui_state(self._ui_state)
        except Exception:
            pass
        # reflect in UI if the button exists
        btn = getattr(self, "_strict_btn", None)
        if btn:
            is_on = bool(self._strict_manual_on)
            btn.setText("Strict Manual: ON" if is_on else "Strict Manual: OFF")
            btn.setProperty("active", is_on)
            btn.style().unpolish(btn); btn.style().polish(btn)
        

    # === Excel quick search popup (ONE suggestion, above input) ===
    def _ensure_excel_popup(self):
        if self._excel_lookup_popup:
            return

        top = QDialog(self)
        top.setWindowFlags(Qt.FramelessWindowHint | Qt.Tool | Qt.WindowStaysOnTopHint)
        top.setModal(False)
        # Don't steal focus from the line edit (lets Enter work reliably)
        top.setFocusPolicy(Qt.NoFocus)

        # One place to activate the import from chip (click or double-click)
        def _activate_from_chip():
            if self._excel_lookup_items:
                self._import_saved_excel(self._excel_lookup_items[0])

        # Also import on double-click anywhere on the chip surface
        top.mouseDoubleClickEvent = lambda e: _activate_from_chip()

        layout = QVBoxLayout(top)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        btn = QPushButton("")
        btn.setFlat(True)
        btn.setMinimumHeight(28)
        btn.clicked.connect(_activate_from_chip)
        layout.addWidget(btn)

        self._excel_lookup_popup = top
        self._excel_suggest_btn = btn

    def _place_excel_popup(self):
        edit = _safe_widget(self, "_excel_lookup_edit")
        popup = self._excel_lookup_popup
        btn   = self._excel_suggest_btn

        if not (edit and popup and _is_alive(btn)):
            return

        h = max(28, btn.sizeHint().height())
        top_left = edit.mapToGlobal(QPoint(0, 0))
        x = top_left.x()
        y = top_left.y() - h - 4
        w = edit.width()
        x, y = _clamp_to_screen(x, y, w, h)
        popup.resize(w, h)
        popup.move(x, y)

        if not popup.isVisible():
            popup.setAttribute(Qt.WA_ShowWithoutActivating, True)
            popup.setWindowFlag(Qt.WindowStaysOnTopHint, True)
            popup.show()

    # === [STEP 5] Helpers to update current selection and label ===
    def _excel_set_match_index(self, i: int):
        try:
            ms = getattr(self, "_excel_lookup_matches", [])
            if not ms:
                return
            i = max(0, min(i, len(ms)-1))
            self._excel_lookup_index = i
            btn = getattr(self, "_excel_suggest_btn", None)
            if btn and _is_alive(btn):
                base = ms[i]["name"]
                extra = len(ms) - 1
                btn.setText(f"{base}  (+{extra})" if extra > 0 else base)
                self._place_excel_popup()
        except Exception:
            pass

    def _excel_cycle_match(self, delta: int):
        try:
            i = int(getattr(self, "_excel_lookup_index", 0))
            self._excel_set_match_index(i + delta)
        except Exception:
            pass
        

    def _on_excel_date_enter(self):
        """When user presses Enter in the Excel search box with a date."""
        edit = getattr(self, "_excel_lookup_edit", None)
        if not edit:
            return

        # If date matches chip is visible, Enter imports the current one
        try:
            if getattr(self, "_excel_date_matches", None) and getattr(self, "_excel_lookup_popup", None):
                if self._excel_lookup_popup.isVisible() and self._excel_date_matches:
                    idx = int(getattr(self, "_excel_date_index", 0) or 0)
                    cur = self._excel_date_matches[idx]
                    self._open_recent_excel(cur["path"], cur.get("sheet") or None)
                    self._hide_excel_popup()
                    self._excel_date_matches = []
                    return
        except Exception:
            pass

        raw = (edit.text() or "").strip()
        target = parse_user_date(raw)
        if not target:
            # Not a date → fall back to original filename search
            try:
                return App.__orig__on_excel_lookup_enter(self)
            except Exception:
                return

        # Busy hint
        try:
            self.status.setText(f"Scanning files for {target.strftime('%d/%m/%Y')} …")
        except Exception:
            pass

        # Create worker
        cfg = load_headers_cfg() if "load_headers_cfg" in globals() else DEFAULT_HEADERS_CFG
        self._date_scan_worker = DateScanWorker(target, cfg, per_root_cap=300, max_depth=4)

        def _on_progress(path_str: str):
            try:
                base = os.path.basename(path_str)
                if hasattr(self, "status") and self.status:
                    self.status.setText(f"Scanning… {base}")
            except Exception:
                pass

        def _on_fail(err: Exception):
            try:
                self.status.setText("Scan failed.")
            except Exception:
                pass
            try:
                QMessageBox.warning(self, "Date Scan", f"Scan error:\n{err}")
            except Exception:
                pass
            try:
                _stop_thread(self._date_scan_worker)
            except Exception:
                pass
            self._date_scan_worker = None

        def _on_done(matches: list):
            try:
                self.status.setText(f"Found {len(matches)} file(s).")
            except Exception:
                pass

            if not matches:
                try:
                    QMessageBox.information(self, "No matches", f"No files found for Start Date = {raw}")
                except Exception:
                    pass
                try:
                    _stop_thread(self._date_scan_worker)
                except Exception:
                    pass
                self._date_scan_worker = None
                return

            if len(matches) == 1:
                choice = matches[0]
                self._open_recent_excel(choice["path"], choice.get("sheet") or None)
                try:
                    _stop_thread(self._date_scan_worker)
                except Exception:
                    pass
                self._date_scan_worker = None
                return

            # Multiple: show ALL matches in a popup menu above the search box
            edit = getattr(self, "_excel_lookup_edit", None)
            if not edit:
                # fallback: first
                self._open_recent_excel(matches[0]["path"], matches[0].get("sheet") or None)
                try: _stop_thread(self._date_scan_worker)
                except Exception: pass
                self._date_scan_worker = None
                return

            menu = QMenu(self)
            for idx, m in enumerate(matches):
                name   = m.get("name") or os.path.basename(m.get("path",""))
                folder = os.path.basename(os.path.dirname(m.get("path","")))
                sheet  = m.get("sheet") or ""
                label  = f"{idx+1}. {name}  •  {folder}" + (f" [{sheet}]" if sheet else "")
                act = menu.addAction(label)
                act.triggered.connect(lambda _=False, p=m["path"], s=m.get("sheet") or None: self._open_recent_excel(p, s))

            # position the menu directly ABOVE the line edit
            h = edit.height()
            pt = edit.mapToGlobal(edit.rect().topLeft())
            menu.exec(pt)

            # Also update the small chip to show the first entry + count (optional)
            self._ensure_excel_popup()
            btn = getattr(self, "_excel_suggest_btn", None)
            if btn:
                first = matches[0]
                folder = os.path.basename(os.path.dirname(first['path']))
                btn.setText(f"1/{len(matches)}  {first['name']}  •  {folder}" + (f" [{first.get('sheet')}]" if first.get('sheet') else ""))
                self._place_excel_popup()

            # store for Enter-to-import current
            self._excel_date_matches = matches
            self._excel_date_index   = 0


            def _render_chip():
                cur = self._excel_date_matches[self._excel_date_index]
                folder = os.path.basename(os.path.dirname(cur['path']))
                label = f"{self._excel_date_index+1}/{len(self._excel_date_matches)}  {cur['name']}  •  {folder}"
                sh = cur.get("sheet")
                if sh:
                    label += f" [{sh}]"
                btn.setText(label)

            def _cycle_next():
                if not getattr(self, "_excel_date_matches", None):
                    return
                self._excel_date_index = (self._excel_date_index + 1) % len(self._excel_date_matches)
                _render_chip()

            def _import_current():
                cur = self._excel_date_matches[self._excel_date_index]
                self._open_recent_excel(cur["path"], cur.get("sheet") or None)

            try:
                try:
                    btn.clicked.disconnect()
                except Exception:
                    pass
                btn.clicked.connect(_cycle_next)
                btn.mouseDoubleClickEvent = lambda e: _import_current()
            except Exception:
                pass

            self._place_excel_popup()

            def _context_menu(point):
                menu = QMenu(self)
                for idx, m in enumerate(self._excel_date_matches):
                    folder = os.path.basename(os.path.dirname(m['path']))
                    act = menu.addAction(f"{idx+1}. {m['name']}  •  {folder}" + (f" [{m.get('sheet')}]" if m.get('sheet') else ""))
                    act.triggered.connect(
                        lambda _=False, i=idx: (
                            setattr(self, "_excel_date_index", i),
                            _render_chip(),
                            self._open_recent_excel(self._excel_date_matches[i]['path'], self._excel_date_matches[i].get('sheet') or None)
                        )
                    )
                menu.exec(btn.mapToGlobal(point))

            try:
                btn.setContextMenuPolicy(Qt.CustomContextMenu)
                btn.customContextMenuRequested.connect(_context_menu)
            except Exception:
                pass

            _render_chip()

            try:
                _stop_thread(self._date_scan_worker)
            except Exception:
                pass
            self._date_scan_worker = None

        self._date_scan_worker.progress.connect(_on_progress)
        self._date_scan_worker.failed.connect(_on_fail)
        self._date_scan_worker.finished_ok.connect(_on_done)
        self._date_scan_worker.start()
        self._date_scan_worker.finished_ok.connect(_on_done)
        self._date_scan_worker.start()



    def _show_recent_downloads_menu(self):
        """Show a dropdown list. If we have date matches, show those; else show recent downloads."""
        btn = getattr(self, "_excel_recent_btn", None)
        if not _is_alive(btn):
            return

        menu = QMenu(self)

        # If we have date-matches from the scanner, prefer showing those
        matches = getattr(self, "_excel_date_matches", None)
        if matches:
            for idx, m in enumerate(matches):
                name = m.get("name") or os.path.basename(m.get("path",""))
                folder = os.path.basename(os.path.dirname(m.get("path","")))
                sheet = m.get("sheet") or ""
                label = f"{idx+1}. {name}  •  {folder}" + (f" [{sheet}]" if sheet else "")
                act = menu.addAction(label)
                act.triggered.connect(lambda _=False, p=m["path"], s=m.get("sheet") or None: self._open_recent_excel(p, s))
            menu.exec(btn.mapToGlobal(btn.rect().bottomLeft()))
            return

        # Otherwise: recent downloads fallback
        try:
            recents = _recent_excels_anywhere(limit=5, max_depth=3, per_root_cap=300)
        except Exception:
            recents = []

        if not recents:
            act = menu.addAction("No recent Excel files found")
            act.setEnabled(False)
        else:
            for p in recents:
                try:
                    ts = datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
                except Exception:
                    ts = ""
                folder = os.path.basename(os.path.dirname(str(p)))
                label = f"{p.name}  •  {folder}" + (f"  •  {ts}" if ts else "")
                act = menu.addAction(label)
                act.triggered.connect(lambda _=False, path=str(p): self._open_recent_excel(path))

        menu.exec(btn.mapToGlobal(btn.rect().bottomLeft()))



        

    def _open_recent_excel(self, full_path: str, sheet_name: Optional[str] = None):
        """
        Import a recently-downloaded Excel file directly (fast pandas path).
        Uses first sheet by default. Remembers for Quick Import.
        """
        try:
            if not full_path or not os.path.exists(full_path):
                QMessageBox.warning(self, "Recent Downloads", "File not found on disk.")
                return

            file_name = os.path.basename(full_path)
            sheet_name = ""  # unknown; use default first sheet

            # Busy overlay during background read
            self._show_busy(f"Reading {file_name}…")
            self._import_worker = ExcelFastImportWorker(full_path, None)

            def _ok(rows, mapping):
                try:
                    # app/wb/ws are None in this path
                    self._finalize_import(
                        rows, mapping, file_name, full_path, sheet_name,
                        None, None, None, used_fast=True
                    )
                    # Remember for Quick Import (sheet left blank)
                    try:
                        remember_excel_source(file_name, full_path, sheet_name)
                    except Exception:
                        pass
                finally:
                    self._hide_busy()
                    self._import_worker = None

            def _fail(err: Exception):
                try:
                    QMessageBox.critical(self, "Import", f"Could not read file:\n{err}")
                finally:
                    self._hide_busy()
                    self._import_worker = None

            self._import_worker.finished_ok.connect(_ok)
            self._import_worker.failed.connect(_fail)
            self._import_worker.start()

        except Exception as e:
            self._hide_busy()
            QMessageBox.critical(self, "Import", f"Unexpected error:\n{e}")
    
        


    def _on_excel_lookup_typing(self, q: str):
        edit = _safe_widget(self, "_excel_lookup_edit")
        if not edit:
            self._hide_excel_popup()
            return

        q = (q or "").strip()
        if not edit.hasFocus() or len(q) < EXCEL_LOOKUP_MIN_CHARS:
            self._hide_excel_popup()
            return

        items = search_excel_sources_by_name(q)
        self._excel_lookup_items = items

        if not items:
            self._hide_excel_popup()
            return

        self._ensure_excel_popup()
        name = items[0].get("name", "")
        sheet = items[0].get("sheet", "")
        if _is_alive(self._excel_suggest_btn):

            self._excel_suggest_btn.setText(f"{name} [{sheet}]".strip())
        self._place_excel_popup()

    


    def _on_excel_lookup_enter(self):
        if self._excel_lookup_items:
            self._import_saved_excel(self._excel_lookup_items[0])
            return
        q = (self._excel_lookup_edit.text() or "").strip()
        if len(q) < EXCEL_LOOKUP_MIN_CHARS:
            self._hide_excel_popup()
            return
        matches = search_excel_sources_by_name(q)
        if matches:
            self._import_saved_excel(matches[0])
        else:
            self._hide_excel_popup()


    def _import_saved_excel(self, src: dict):
        """Quick Import: prefer reading from the remembered file (shows even incomplete rows).
        DB is used as a fallback, and only complete rows are saved to the DB.
        """
        name  = (src.get("name")  or "").strip()     # workbook base name
        sheet = (src.get("sheet") or "").strip()
        path  = (src.get("path")  or "").strip()

        if not name:
            QMessageBox.critical(self, "Import", "No file name selected.")
            self._hide_excel_popup()
            return

        # --- Prefer the file on disk (shows rows even if BRAND is missing) ---
        if path and os.path.exists(path):
            try:
                fresh_rows, mapping = _read_excel_fast(path, sheet or None)
            except Exception as e:
                # If file read fails, fall back to DB
                fresh_rows, mapping = [], {h: h for h in all_headers()}
                QMessageBox.warning(self, "Import", f"Could not read file directly:\n{e}\n\nFalling back to DB…")

            if fresh_rows:
                # Enrich + persist (DB gate will skip incomplete rows by design)
                enriched = []
                for r in fresh_rows:
                    nr = {k: "" for k in all_headers()}
                    nr.update(r)
                    nr["SOURCE_FILE"]  = name
                    nr["SOURCE_SHEET"] = sheet
                    enriched.append(nr)
                try:
                    upsert_db_rows(enriched)
                except Exception as e:
                    QMessageBox.critical(self, "DB Save", f"Could not save rows from Quick Import:\n{e}")

                # UI should reflect what we actually parsed (even if not saved)
                saved_count = sum(1 for r in enriched if _is_complete_db_row(r))
                self.status.setText(
                    f"Loaded {len(fresh_rows)} rows from file: {name}"
                    + (f" [{sheet}]" if sheet else "")
                    + f"  •  saved {saved_count} complete rows to DB"
                )
                self.connected = (None, None, None)
                self.preview_rows = fresh_rows
                self.preview_qty  = [1] * len(fresh_rows)
                self.last_mapping = mapping
                self._hide_excel_popup()
                self._build_generate_from_excel()

                # Remember recency & prune
                try:
                    remember_excel_source(name, path, sheet)
                    _prune_db_to_recent_sources(limit=15)
                except Exception:
                    pass
                return

        # --- Fallback: use whatever is in the DB for this (file, sheet) ---
        allrows = load_db_rows()
        rows = [dict(r) for r in allrows
                if (r.get("SOURCE_FILE", "") == name) and (not sheet or r.get("SOURCE_SHEET", "") == sheet)]

        if not rows:
            # Nothing in DB and no readable file → explain next step
            msg = (f"No saved rows in DB for '{name}'"
                   f"{f' [{sheet}]' if sheet else ''}"
                   + ("" if path else "\n(There is no remembered file path.)")
                   + "\n\nOpen the workbook and press 'Connect to Excel File' once on this PC.")
            QMessageBox.information(self, "Import", msg)
            self._hide_excel_popup()
            return

        # Show what the DB has (will be only the complete rows)
        mapping = {h: h for h in all_headers()}
        self.status.setText(f"Loaded {len(rows)} rows from DB: {name}" + (f" [{sheet}]" if sheet else ""))
        self.connected = (None, None, None)
        self.preview_rows = rows
        self.preview_qty  = [1] * len(rows)
        self.last_mapping = mapping
        self._hide_excel_popup()
        self._build_generate_from_excel()

        try:
            remember_excel_source(name, path, sheet)
            _prune_db_to_recent_sources(limit=15)
        except Exception:
            pass



    def _priv_add_template(self):
        if not self._check_privileged("Add Template"):
            return
        dlg = AddTemplateDialog(self)
        dlg.exec()
        if not dlg.result:
            return

        name, data = dlg.result
        safe = re.sub(r"[^a-zA-Z0-9._ -]", "_", name).strip()
        if not safe:
            QMessageBox.critical(self, "Save Template", "Invalid template name.")
            return

        path = os.path.join(TEMPLATES_DIR, f"{safe}.json")
        if os.path.exists(path):
            if QMessageBox.question(self, "Overwrite?", f"'{safe}' exists. Overwrite?") != QMessageBox.Yes:
                return

        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            QMessageBox.critical(self, "Save Template", str(e))
            return

        QMessageBox.information(self, "Saved", f"Template saved: {safe}")

        frame = getattr(self, "templates_frame", None)
        if _is_alive(frame):
            QTimer.singleShot(0, self._reload_template_buttons)



    def _priv_delete_template(self):
        if not self._check_privileged("Delete Template"):
            return
        dlg = DeleteTemplateDialog(self)
        dlg.exec()
        if not dlg.result:
            return

        name, full = dlg.result
        if not os.path.exists(full):
            QMessageBox.critical(self, "Delete", "Not found. It may have been removed already.")
            return

        if QMessageBox.question(self, "Delete", f"Delete template '{name}'?") != QMessageBox.Yes:
            return

        try:
            os.remove(full)
        except Exception as e:
            QMessageBox.critical(self, "Delete error", str(e))
            return

        if self.selected_template_name == name:
            self.selected_template = None
            self.selected_template_name = None

        QMessageBox.information(self, "Deleted", f"Template removed: {name}")

        frame = getattr(self, "templates_frame", None)
        if _is_alive(frame):
            QTimer.singleShot(0, self._reload_template_buttons)



    def _priv_headers_mgr(self):
        if not self._check_privileged("Headers Manager"): return
        dlg = HeaderManagerDialog(self); dlg.exec()



    def _on_connect(self):
        """Fast connect: pandas on a worker thread when file is saved; COM fallback if needed + busy overlay."""
        self._hide_excel_popup()

        # Discover active Excel workbook/sheet (light COM only)
        try:
            import xlwings as xw
        except Exception:
            QMessageBox.critical(self, "Excel", "Install xlwings:  pip install xlwings")
            return
        try:
            app = xw.apps.active
            wb  = app.books.active
            ws  = wb.sheets.active
        except Exception as e:
            QMessageBox.critical(self, "Excel", f"Could not connect: {e}")
            return

        # Names/paths
        try:
            full = wb.fullname  # empty if unsaved
            file_name  = os.path.basename(full) if full else wb.name
            sheet_name = ws.name
        except Exception:
            full = ""
            file_name  = wb.name
            sheet_name = ws.name

        self.status.setText("Reading…")


        # Saved workbook → run fast pandas read in background
        if full and os.path.exists(full):
            self._import_worker = ExcelFastImportWorker(full, sheet_name or None)

            def _fast_ok(rows, mapping):
                try:
                    self._finalize_import(rows, mapping, file_name, full, sheet_name, app, wb, ws, used_fast=True)
                finally:
                    self._hide_busy()
                    self._import_worker = None

            def _fast_fail(err: Exception):
                # Fallback: current COM-based extractor (main thread)
                try:
                    rows, mapping = extract_rows_from_excel(ws)
                    self._finalize_import(rows, mapping, file_name, full, sheet_name, app, wb, ws, used_fast=False)
                except Exception as e:
                    QMessageBox.critical(
                        self, "Excel",
                        f"Fast read failed:\n{err}\n\nFallback read failed:\n{e}"
                    )
                finally:
                    self._hide_busy()
                    self._import_worker = None

            self._import_worker.finished_ok.connect(_fast_ok)
            self._import_worker.failed.connect(_fast_fail)
            self._import_worker.start()
            return

        # Unsaved workbook → fallback to current extractor
        try:
            rows, mapping = extract_rows_from_excel(ws)
            self._finalize_import(rows, mapping, file_name, full, sheet_name, app, wb, ws, used_fast=False)
        except Exception as e:
            QMessageBox.critical(self, "Excel", f"Read failed: {e}")
        finally:
            self._hide_busy()


    def _show_busy(self, message: str = "Working…"):
        """Lightweight translucent overlay that blocks clicks while work runs."""
        try:
            if getattr(self, "_busy", None):
                # Update message if already showing
                lbl = self._busy.findChild(QLabel, "busy_msg")
                if lbl:
                    lbl.setText(message)
                return
            dlg = QDialog(self, Qt.FramelessWindowHint | Qt.Dialog)
            dlg.setModal(True)  # ApplicationModal by default for QDialog children
            dlg.setAttribute(Qt.WA_TranslucentBackground, True)
            dlg.setObjectName("BusyOverlay")

            # Root layout (full-screen overlay)
            root = QVBoxLayout(dlg)
            root.setContentsMargins(0, 0, 0, 0)

            # Center card
            card = QFrame()
            card.setObjectName("BusyCard")
            card_layout = QVBoxLayout(card)
            card_layout.setContentsMargins(22, 18, 22, 18)
            card_layout.setSpacing(8)
            msg = QLabel(message)
            msg.setObjectName("busy_msg")
            msg.setAlignment(Qt.AlignCenter)
            sub = QLabel("Please wait…")
            sub.setAlignment(Qt.AlignCenter)
            card_layout.addWidget(msg)
            card_layout.addWidget(sub)

            root.addStretch(1)
            row = QHBoxLayout()
            row.addStretch(1)
            row.addWidget(card)
            row.addStretch(1)
            root.addLayout(row)
            root.addStretch(1)

            # Styling for overlay + card (no external assets)
            dlg.setStyleSheet("""
            QDialog#BusyOverlay { background: rgba(0,0,0,110); }
            QFrame#BusyCard {
                background: #FFFFFF; border: 1px solid #D8D2CA; border-radius: 12px;
            }
            QLabel#busy_msg { font-weight: 600; font-size: 14px; }
            """)

            # Size/position to cover window
            g = self.geometry()
            dlg.setGeometry(g)
            dlg.show()
            self._busy = dlg
        except Exception:
            self._busy = None  # never block if something goes wrong


    def _hide_busy(self):
        """Close the busy overlay if shown."""
        try:
            if getattr(self, "_busy", None):
                self._busy.close()
        finally:
            self._busy = None



    def _finalize_import(self, rows, mapping, file_name, full, sheet_name, app, wb, ws, used_fast: bool):
        """Common post-read path: normalize, upsert, remember, and navigate."""
        # Normalize
        for r in rows:
            r["BARCODE"] = clean_barcode(r.get("BARCODE", ""))
            for p in ("REG", "PROMO", "COOP", "REGULAR_PRICE", "PROMO_PRICE"):
                r[p] = price_text(r.get(p, ""))
            for d in ("START_DATE", "END_DATE"):
                r[d] = date_only(r.get(d, ""))

        # Persist to DB
        try:
            enriched = []
            for r in rows:
                nr = dict({k: "" for k in all_headers()})
                nr.update(r)
                nr["SOURCE_FILE"]  = file_name
                nr["SOURCE_SHEET"] = sheet_name
                enriched.append(nr)
            upsert_db_rows(enriched)
        except Exception as e:
            QMessageBox.critical(self, "DB Save", f"Could not save rows: {e}")

        # Remember source for Quick Import
        try:
            remember_excel_source(file_name, full, sheet_name)
        except Exception:
            pass

        try:
            _prune_db_to_recent_sources(limit=15)
        except Exception:
            pass    

        # Save handles
        self.connected = (app, wb, ws)

        # ✅ UI should reflect what was actually read, even if some rows weren’t saved
        #    (DB gate blocks rows missing BARCODE/BRAND/ITEM, but we still want to show them here)
        self.preview_rows = rows
        self.preview_qty  = [1] * len(self.preview_rows)
        self.last_mapping = mapping

        how = "fast reader" if used_fast else "Excel"
        self.status.setText(
            f"Connected: {file_name} → {sheet_name} (imported {len(rows)} rows via {how}; "
            f"saved {sum(1 for r in rows if _is_complete_db_row(r))} complete rows to DB)"
        )
        self._build_mapping()



    def _build_mapping(self):
        self._hide_excel_popup()
        e = getattr(self, "_excel_lookup_edit", None)
        if _is_alive(e):
            try:
                e.blockSignals(True)
                e.clear()
                e.blockSignals(False)
            except RuntimeError:
                pass
        self._excel_lookup_edit = None

        self.clear(); self._preferred_size = "small"; self.set_small()
        self.clear(); self._preferred_size = "small"; self.set_small()
        top = QFrame(objectName="Card"); top_layout = QHBoxLayout(top); self.main_layout.addWidget(top)
        back = QLabel("‹ Back"); back.setObjectName("Back"); back.setCursor(QCursor(Qt.PointingHandCursor)); top_layout.addWidget(back)
        back.mousePressEvent = lambda e: self._build_home()
        title = QLabel("Mapped Headers"); title.setObjectName("Title"); top_layout.addWidget(title)

        body = QFrame(objectName="Card"); body_layout = QVBoxLayout(body); self.main_layout.addWidget(body)
        body_layout.addWidget(QLabel("Auto-mapped using your Header Manager synonyms.", objectName="Small"))
        list_widget = QListWidget(); body_layout.addWidget(list_widget)
        mapping = self.last_mapping or {}
        for need in all_headers():
            src = mapping.get(need) or "—"
            list_widget.addItem(f"{need:>12}  ⇢  {src}")

        btns = QFrame(objectName="Card"); btns_layout = QHBoxLayout(btns); self.main_layout.addWidget(btns)
        next_btn = QPushButton("Next"); next_btn.setObjectName("Primary"); next_btn.clicked.connect(self._build_generate_from_excel); btns_layout.addWidget(next_btn)

    def _build_generate_from_excel(self):
        self._build_generate(source="excel")

    def _build_generate(self, source="excel"):
        self.clear(); self._preferred_size = "large"; self.set_large()
        self._current_gen_source = source
        self._excel_coop_only = False

        top = QFrame(objectName="Card"); top_layout = QHBoxLayout(top); self.main_layout.addWidget(top)
        back = QLabel("‹ Back"); back.setObjectName("Back"); back.setCursor(QCursor(Qt.PointingHandCursor)); top_layout.addWidget(back)
        back.mousePressEvent = lambda e: (setattr(self, "_preferred_size", "small"), self.set_small(),
                                          self._build_mapping() if source == "excel" else self._build_manual())
        title = QLabel("Select Template & Generate"); title.setObjectName("Title"); top_layout.addWidget(title)

        self.templates_frame = QFrame(objectName="Card")
        # Flowing, wrapping layout (like flex-wrap)
        flow = FlowLayout(self.templates_frame, margin=8, hspacing=30, vspacing=30)
        self.templates_frame.setLayout(flow)

        self.templates_scroll = QScrollArea()
        self.templates_scroll.setWidget(self.templates_frame)
        self.templates_scroll.setWidgetResizable(True)
        self.templates_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.templates_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.templates_scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.templates_scroll.setFixedHeight(220)  # keep your gallery height
        self.templates_scroll.setContentsMargins(0, 0, 0, 0)  # keep scroll area tight
        self.templates_frame.setContentsMargins(16, 8, 16, 8) # extra padding inside the card

        self.main_layout.addWidget(self.templates_scroll)

        QTimer.singleShot(0, self._reload_template_buttons)

        gen_row = QFrame(objectName="Card"); gen_row_layout = QHBoxLayout(gen_row)
        gen_row_layout.setContentsMargins(8, 8, 8, 8); gen_row_layout.setSpacing(16)

        self.gen_btn = QPushButton("Generate"); self.gen_btn.setObjectName("GenerateOff")
        self.gen_btn.setFixedHeight(34); self.gen_btn.setMinimumWidth(160); self.gen_btn.setFont(QFont("Arial", 11))
        self.gen_btn.clicked.connect(lambda: self._on_generate(source))
        gen_row_layout.addStretch()
        gen_row_layout.addWidget(self.gen_btn, 0, Qt.AlignCenter)

        smart_ai = QPushButton("Smart AI ✨")
        smart_ai.setObjectName("SmartAI")
        smart_ai.setFixedHeight(36)
        smart_ai.setMinimumWidth(180)
        smart_ai.clicked.connect(self.smart_ai_render)
        smart_ai.setStyleSheet("""
        QPushButton#SmartAI {
            color: white; font-weight:600; font-size:12px; font-family: Arial;
            padding: 8px 18px; border: none; border-radius: 12px;
            background: qlineargradient(x1:0,y1:0, x2:0,y2:1, stop:0 #3B82F6, stop:1 #1D4ED8);
        }
        QPushButton#SmartAI:hover {
            background: qlineargradient(x1:0,y1:0, x2:0,y2:1, stop:0 #60A5FA, stop:1 #2563EB);
        }
        QPushButton#SmartAI:pressed {
            background: qlineargradient(x1:0,y1:0, x2:0,y2:1, stop:0 #1E40AF, stop:1 #1E3A8A);
        }""")
        glow = QGraphicsDropShadowEffect(smart_ai)  # why: draw attention to primary
        glow.setColor(QColor(59, 130, 246, 180))
        glow.setBlurRadius(28); glow.setOffset(0, 4)
        smart_ai.setGraphicsEffect(glow)
        self._smart_ai_glow_anim = QPropertyAnimation(glow, b"blurRadius", self)
        self._smart_ai_glow_anim.setStartValue(22); self._smart_ai_glow_anim.setEndValue(40)
        self._smart_ai_glow_anim.setDuration(1000); self._smart_ai_glow_anim.setEasingCurve(QEasingCurve.InOutQuad)
        self._smart_ai_glow_anim.setLoopCount(-1); self._smart_ai_glow_anim.start()

        gen_row_layout.addWidget(smart_ai, 0, Qt.AlignCenter)
        gen_row_layout.addStretch()
        self.main_layout.addWidget(gen_row)
        self._refresh_gen_btn()

        if source == "excel":
            wrap = QFrame(objectName="Card"); wrap_layout = QVBoxLayout(wrap); self.main_layout.addWidget(wrap)
            self.main_layout.setStretchFactor(self.templates_scroll, 0)
            self.main_layout.setStretchFactor(gen_row, 0)
            self.main_layout.setStretchFactor(wrap, 1)

            sb = QHBoxLayout(); wrap_layout.addLayout(sb)
            sb.addWidget(QLabel("Search All", objectName="Small"))
            self._excel_search_edit = QLineEdit()
            self._excel_search_edit.setFont(QFont("Arial", 14)); self._excel_search_edit.setFixedHeight(34); self._excel_search_edit.setMinimumWidth(520)
            sb.addWidget(self._excel_search_edit, 1, alignment=Qt.AlignLeft)
            clear_excel_search = QPushButton("✖"); clear_excel_search.setFixedWidth(20); sb.addWidget(clear_excel_search)

            # Clear all header filters (SECTION/BRAND)
            clear_hdr_filters = QPushButton("🧹")
            clear_hdr_filters.setFixedWidth(28)
            clear_hdr_filters.setToolTip("Clear header filters")
            sb.addWidget(clear_hdr_filters)
            clear_hdr_filters.clicked.connect(lambda: self.tree.clear_filters())

            self._excel_coop_btn = QPushButton("CO-OP only: OFF")
            self._excel_coop_btn.setFixedHeight(28); self._excel_coop_btn.setMinimumWidth(130)
            self._excel_coop_btn.setToolTip("Show only rows with COOP price > 0.00")
            sb.addWidget(self._excel_coop_btn)

            # --- Debounced search wiring (Excel screen) ---
            if not hasattr(self, "_debounce_excel_search") or self._debounce_excel_search is None:
                self._debounce_excel_search = Debouncer(150, lambda: self._excel_refresh_table(), self)
            self._excel_search_edit.textChanged.connect(lambda _t: self._debounce_excel_search.call())
            clear_excel_search.clicked.connect(lambda: (self._excel_search_edit.setText(""), self._debounce_excel_search.call()))

            def _flip_coop_only():
                self._excel_coop_only = not getattr(self, "_excel_coop_only", False)
                self._excel_coop_btn.setText("CO-OP only: ON" if self._excel_coop_only else "CO-OP only: OFF")
                self._excel_refresh_table()  # filtering change should apply immediately
            self._excel_coop_btn.clicked.connect(_flip_coop_only)

            wrap_layout.addWidget(QLabel("Imported Data (☑ = selected; double-click Q to edit)", objectName="Small"))

            # >>> CHANGE: Always use LEGACY columns for the Excel grid (even if Fresh is ON)
            excel_cols = ("CHK","Q","SECTION","BARCODE","BRAND","ITEM","REG","PROMO","START","END","COOP")

            # Enable Excel-like dropdown filters only for columns that exist
            filterable_cols = tuple(c for c in ("SECTION", "BRAND") if c in excel_cols)

            self.tree = FilterableTable(
                excel_cols,
                enable_filters=bool(filterable_cols),
                parent=self,
                filterable_columns=filterable_cols,
            )
            self.tree.external_refresh = True                          # we render rows ourselves
            self.tree.get_all_rows = self._excel_rows_for_filters      # supply values for dropdowns
            self.tree.on_filters_changed = self._excel_refresh_table   # re-render grid when filters change
            self.tree.setObjectName("ExcelTable")
            wrap_layout.addWidget(self.tree)

            # keep existing interactions intact
            self.tree.cellDoubleClicked.connect(self._edit_qty_excel)
            self.tree.cellClicked.connect(self._excel_on_click)
            
            self.tree.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
            self.tree.on_header_toggle_all = self._excel_toggle_all_visible

            self._excel_row_keys = [self._excel_row_key(r) for r in self.preview_rows]
            self._excel_refresh_table()
            return

        # Manual path (unchanged here)
        body = QFrame(objectName="Card"); body_layout = QVBoxLayout(body); self.main_layout.addWidget(body)
        self.main_layout.setStretchFactor(self.templates_scroll, 0)
        self.main_layout.setStretchFactor(gen_row, 0)
        self.main_layout.setStretchFactor(body, 1)

        body_layout.addWidget(QLabel("Your List", objectName="Small"))
        self.stage = FilterableTable(("QTY","BARCODE","BRAND","ITEM","REG","PROMO","START","END","COOP"), enable_filters=False)
        self.stage.setObjectName("StageTable")
        self.stage.setMinimumHeight(680)
        body_layout.addWidget(self.stage)
        self.stage.cellDoubleClicked.connect(self._on_stage_double_click)
        self._stage_widths_tuned = False 
        if not hasattr(self, "_manual_search_edit"):
            self._manual_search_edit = QLineEdit()
        self._manual_refresh_stage_table()
        return



        

    # === Template gallery: cached image thumbnails + hover zoom ===
    def _reload_template_buttons(self):
        frame = getattr(self, "templates_frame", None)
        if not _is_alive(frame):
            return

        layout = frame.layout()
        if layout is None:
            return

        # Remove any existing items safely
        while layout.count():
            it = layout.takeAt(0)
            w = it.widget()
            if w is not None:
                w.deleteLater()

        files = _list_template_files()
        if not files:
            label = QLabel("(No templates yet — add one from JSON on Home)")
            label.setObjectName("Small")
            layout.addWidget(label)
            self.selected_template = None
            self.selected_template_name = None
            return

        SMALL_SZ = QSize(140, 90)
        LARGE_SZ = QSize(440, 280)

        for name, full in files:
            data = _read_json(full) or {}

            pm_small, pm_large = _pixmaps_from_template_json(data, SMALL_SZ, LARGE_SZ)

            if pm_small and pm_large:
                tb = HoverPreviewButton(pm_small, pm_large)
                # lock tile size so wrapping is predictable
                tile_w = pm_small.width() + 18
                tile_h = pm_small.height() + 40
                tb.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                tb.setMinimumSize(tile_w, tile_h)
                tb.setMaximumSize(QSize(tile_w, tile_h))
                tb.setAutoRaise(True)
                tb.setToolTip(name)
                tb.setText(name)
                tb.clicked.connect(lambda _=False, nm=name, dt=data: self._pick_template(nm, dt))
                layout.addWidget(tb)
            else:
                btn = QPushButton(name)
                btn.setObjectName("Template")
                btn.setFixedSize(140, 30)  # fixed tile so FlowLayout can wrap neatly
                btn.setFont(QFont("Arial", 10, QFont.Medium))
                btn.clicked.connect(lambda _=False, nm=name, dt=data: self._pick_template(nm, dt))
                layout.addWidget(btn)


    def _pick_template(self, name: str, data: dict):
        """Select a template and persist last selection."""
        self.selected_template = data
        self.selected_template_name = name
        self._refresh_gen_btn()

        # persist last template name
        try:
            self._ui_state = getattr(self, "_ui_state", load_ui_state())
            self._ui_state["last_template_name"] = name or ""
            save_ui_state(self._ui_state)
            self._last_template_name_saved = name or ""
        except Exception:
            pass

    def _refresh_gen_btn(self):
        is_on = bool(self.selected_template)
        self.gen_btn.setObjectName("GenerateOn" if is_on else "GenerateOff")
        self.gen_btn.setEnabled(is_on)
        self.gen_btn.setText(f"Generate ({self.selected_template_name})" if is_on else "Generate")
        self.gen_btn.setStyle(self.gen_btn.style())

    def _manual_stage_match(self, rec: Dict[str,str], q: str) -> bool:
        if not q: return True
        ql = q.lower()
        for v in (
            rec.get("SECTION",""),
            rec.get("BARCODE",""), rec.get("PLU",""),
            rec.get("BRAND",""),   rec.get("ARABIC_DESCRIPTION",""),
            rec.get("ITEM",""),    rec.get("ENGLISH_DESCRIPTION",""),
            rec.get("REG",""), rec.get("PROMO",""),
            rec.get("START_DATE",""), rec.get("END_DATE",""),
            rec.get("COOP",""),
        ):
            if v and ql in str(v).lower():
                return True
        return False

    def _make_qty_widget(self, initial_value: int, on_change: Callable[[int], None]) -> QWidget:
        """
        Polished, compact qty control:
          • Light, borderless +/- with subtle circular hover
          • Centered number in a soft 'pill'
          • Tight spacing so the Q column stays airy
        Works for both Excel and Manual tables (same callback contract).
        """
        w = QWidget()
        h = QHBoxLayout(w)
        h.setContentsMargins(2, 0, 2, 0)
        h.setSpacing(6)

        # --- shared styles ---
        btn_css = """
            QToolButton {
                border: none;
                background: transparent;
                padding: 0px 4px;
                color: #353535;
            }
            QToolButton:hover {
                background: rgba(0,0,0,0.06);
                border-radius: 12px;          /* subtle circular hover chip */
                color: #111111;
            }
            QToolButton:pressed {
                background: rgba(0,0,0,0.10);
                border-radius: 12px;
                color: #666666;
            }
        """
        pill_css = """
            QLineEdit {
                background: #F2EEE8;
                border: 1px solid #E0D9CF;
                border-radius: 10px;
                padding: 0px 8px;
                color: #1E1E1E;
            }
        """

        def _mk_btn(symbol: str) -> QToolButton:
            b = QToolButton(w)
            b.setAutoRaise(True)     # flat, native
            b.setText(symbol)
            f = b.font(); f.setPointSize(14); f.setBold(True)
            b.setFont(f)
            b.setCursor(QCursor(Qt.PointingHandCursor))
            b.setStyleSheet(btn_css)
            b.setFocusPolicy(Qt.NoFocus)
            return b

        minus = _mk_btn("−")
        plus  = _mk_btn("+")

        # Number 'pill'
        disp = QLineEdit(w)
        disp.setReadOnly(True)
        disp.setAlignment(Qt.AlignCenter)
        fnum = disp.font()
        fnum.setPointSize(13)
        fnum.setBold(True)
        disp.setFont(fnum)
        disp.setFixedHeight(24)
        disp.setFixedWidth(44)       # compact but readable
        disp.setStyleSheet(pill_css)
        disp.setFocusPolicy(Qt.NoFocus)

        # --- behavior ---
        def set_val(v: int):
            v = max(0, int(v))
            disp.setText(str(v))
            try:
                on_change(v)
            except Exception:
                pass

        try:
            start = int(initial_value)
            if start < 0:
                start = 1
        except Exception:
            start = 1
        set_val(start)

        def _dec():
            try: v = int(disp.text() or "1")
            except Exception: v = 1
            set_val(max(0, v - 1))

        def _inc():
            try: v = int(disp.text() or "1")
            except Exception: v = 1
            set_val(v + 1)

        minus.clicked.connect(_dec)
        plus.clicked.connect(_inc)

        # Layout: keep it balanced and compact
        h.addWidget(minus)
        h.addWidget(disp)
        h.addWidget(plus)

        # Slightly larger minimum so it doesn't clip on some styles
        w.setMinimumHeight(26)
        return w


    def _manual_refresh_stage_table(self):
        q = self._manual_search_edit.text() if hasattr(self, "_manual_search_edit") else ""
        rows_for_table = []
        # Build rows (without the qty widget yet)
        for i, d in enumerate(self.staged_rows):
            if self._manual_stage_match(d, q):
                qty = self.staged_qty[i] if i < len(self.staged_qty) else 1
                rows_for_table.append((
                    qty,
                    d.get("BARCODE",""),
                    d.get("BRAND",""),
                    d.get("ITEM",""),
                    d.get("REG",""),
                    d.get("PROMO",""),
                    d.get("START_DATE",""),
                    d.get("END_DATE",""),
                    d.get("COOP",""),
                ))

        # Attach basic rows first
        self.stage.attach_rows(rows_for_table)

        # Find QTY column index (should be 0 by our header order, but resolve by name)
        cols = getattr(self.stage, "_columns", ())
        try:
            qty_col = cols.index("QTY")
        except Exception:
            qty_col = 0

        # Replace QTY cells with +/- widget, wiring back to self.staged_qty
        for r in range(self.stage.rowCount()):
            row_qty = 1
            it = self.stage.item(r, qty_col)
            if it:
                try:
                    row_qty = int(it.text())
                except Exception:
                    row_qty = 1

            # Map visible row r back to staged_rows index.
            # We use BARCODE+ITEM to find original index safely.
            bc = self.stage.item(r, cols.index("BARCODE")).text() if "BARCODE" in cols else ""
            itxt = self.stage.item(r, cols.index("ITEM")).text() if "ITEM" in cols else ""
            # fallback: position-based
            staged_idx = None
            for idx, d in enumerate(self.staged_rows):
                if (d.get("BARCODE","") == bc) and (d.get("ITEM","") == itxt):
                    staged_idx = idx
                    break
            if staged_idx is None:
                staged_idx = r if r < len(self.staged_rows) else len(self.staged_rows)-1

            # Ensure staged_qty capacity
            while len(self.staged_qty) <= staged_idx:
                self.staged_qty.append(1)

            def _on_change(new_v: int, i=staged_idx, cell_r=r, cell_c=qty_col):
                self.staged_qty[i] = new_v
                # keep the underlying item text in sync (useful for copy/export)
                it2 = self.stage.item(cell_r, cell_c)
                if it2:
                    it2.setText(str(new_v))

            w = self._make_qty_widget(row_qty, _on_change)
            self.stage.setCellWidget(r, qty_col, w)

        # Tune widths after widgets are in place
        self._tune_stage_column_widths()



    def _excel_row_key(self, rec: Dict[str, str]) -> str:
        # Prefer classic key if available (normalize barcode/brand/item)
        k1 = (
            clean_barcode(rec.get("BARCODE", "")),
            (rec.get("BRAND", "") or "").strip().upper(),
            (rec.get("ITEM", "") or "").strip().upper(),
        )
        if any(k1):
            return "|".join(k1)

        # Fresh-first fallback: prefer PLU, then descriptions + normalized prices
        k2 = (
            (rec.get("PLU", "") or "").strip(),
            (rec.get("ENGLISH_DESCRIPTION", "") or "").strip().upper(),
            (rec.get("ARABIC_DESCRIPTION", "") or "").strip().upper(),
            price_text(rec.get("REGULAR_PRICE", "")),
            price_text(rec.get("PROMO_PRICE", "")),
        )
        if any(k2):
            return "|".join(k2)

        # Last resort: stable hash of the entire record
        import hashlib, json
        return hashlib.sha256(json.dumps(rec, sort_keys=True).encode("utf-8")).hexdigest()[:16]



    def _excel_legacy_cell(self, rec: dict, col: str) -> str:
        """
        Value resolver for the Excel grid.
        When Fresh is ON, show Fresh fields in the legacy columns:
          BARCODE<=PLU, BRAND<=Arabic, ITEM<=English, REG<=Regular Price, PROMO<=Promo Price, COOP<=UOM.
        When Fresh is OFF, still make START/END/SECTION show their mapped values.
        """
        def _get_first(keys):
            for k in keys:
                if k in rec:
                    v = rec.get(k)
                    if v is not None and str(v).strip() != "":
                        return str(v).strip()
            return ""

        # --- Map dates/section for BOTH modes up-front ---
        if col in ("START", "END", "SECTION"):
            if col == "SECTION":
                return _get_first(("SECTION","Section","Dept","DEPT"))
            if col == "START":
                val = _get_first((
                    "START_DATE","Start Date","START","Start","FROM","From",
                    "OFFER START","Offer Start","PROMO START","Promo Start"
                ))
                return date_only(val)
            if col == "END":
                val = _get_first((
                    "END_DATE","End Date","END","End","TO","To",
                    "OFFER END","Offer End","PROMO END","Promo End"
                ))
                return date_only(val)

        # --- Legacy mode: simple passthrough for the rest ---
        if not FRESH_SECTION_ACTIVE:
            return str(rec.get(col, "") or "")

        # --- Fresh mode → legacy columns ---
        if col == "BARCODE":
            plu = _get_first(("PLU","Plu","PLU Number","PLU_NO","PLU No","PLU#"))
            if plu:
                return plu
            has_bar = _get_first(("Barcode","BARCODE","EAN","UPC","CODE"))
            return "" if has_bar != "" else ""

        if col == "BRAND":
            return _get_first((
                "Arabic Description","ARABIC DESCRIPTION","Arabic_Desc","Arabic","Brand Arabic",
                "ARABIC_DESCRIPTION"
            ))

        if col == "ITEM":
            return _get_first((
                "English Description","ENGLISH DESCRIPTION","Description English","English","Item English",
                "ENGLISH_DESCRIPTION"
            ))

        if col == "REG":
            return _get_first(("REGULAR PRICE","Regular Price","REGULAR_PRICE","REGULARPRICE","REGULAR"))

        if col == "PROMO":
            return _get_first(("PROMO PRICE","Promo Price","PROMO_PRICE","PROMOPRICE","PROMO"))

        if col == "COOP":
            return _get_first((
                "UOM","Unit of Measurement","UNIT OF MEASUREMENT","Unit of Measure","Unit Of Measure",
                "Unit of measurement","Unit","U/M","UnitOfMeasure","Measure","UNIT_OF_MEASUREMENT"
            ))

        # Fallback
        return str(rec.get(col, "") or "")



    def _excel_rows_for_filters(self) -> List[List[object]]:
        """
        Provide current Excel-grid rows shaped to the table's columns,
        so the header filter dropdown can list distinct values.
        """
        cols = getattr(self.tree, "_columns", ())
        rows: List[List[object]] = []
        for i, rec in enumerate(self.preview_rows):
            vals: List[object] = []
            for col in cols:
                if col == "CHK":
                    vals.append("☑" if self._excel_row_keys[i] in self._excel_checked_keys else "☐")
                elif col == "Q":
                    qty = self.preview_qty[i] if i < len(self.preview_qty) else 1
                    vals.append(qty)
                else:
                    vals.append(self._excel_legacy_cell(rec, col))
            rows.append(vals)
        return rows

    

    def _excel_match(self, rec: Dict[str, str], q: str) -> bool:
        """
        Case/spacing-insensitive "Search All".
        When Fresh is ON, include PLU + Fresh fields first, then legacy fields.
        """
        if not q:
            return True

        try:
            qn = norm(q)  # uses your existing normalizer (case/space insensitive)
        except Exception:
            qn = str(q).strip().lower()

        haystack: list[str] = []

        # Prioritize Fresh fields when active
        if FRESH_SECTION_ACTIVE:
            haystack.extend([
                rec.get("PLU", ""),
                rec.get("ARABIC_DESCRIPTION", ""),
                rec.get("ENGLISH_DESCRIPTION", ""),
                rec.get("REGULAR_PRICE", ""),
                rec.get("PROMO_PRICE", ""),
            ])

        # Always include legacy fields so mixed sheets still match
        haystack.extend([
            rec.get("SECTION", ""),
            rec.get("BARCODE", ""),
            rec.get("BRAND", ""),
            rec.get("ITEM", ""),
            rec.get("REG", ""),
            rec.get("PROMO", ""),
            rec.get("START_DATE", ""),
            rec.get("END_DATE", ""),
            rec.get("COOP", ""),
        ])

        for s in haystack:
            try:
                if qn in norm(str(s)):
                    return True
            except Exception:
                if qn in str(s).strip().lower():
                    return True
        return False

    def _has_positive_coop(self, rec: Dict[str, str]) -> bool:
        """Return True if COOP price is a valid number > 0.00."""
        raw = rec.get("COOP", "")
        s = price_text(raw)  # normalizes things like "AED 12" -> "12.00"
        try:
            return float(s) > 0.0
        except Exception:
            return False
    


    def _excel_visible_indices(self) -> List[int]:
        q = self._excel_search_edit.text() if hasattr(self, "_excel_search_edit") else ""
        coop_only = getattr(self, "_excel_coop_only", False)
        table_filters = getattr(self.tree, "_filters", {}) if hasattr(self, "tree") else {}

        def passes_filters(rec: Dict[str, str]) -> bool:
            if not table_filters:
                return True
            for col, allowed in table_filters.items():
                if not allowed:
                    continue
                field = {"START": "START_DATE", "END": "END_DATE"}.get(col, col)
                val = str(rec.get(field, ""))
                if val not in allowed:
                    return False
            return True

        vis: List[int] = []
        for i, rec in enumerate(self.preview_rows):
            if not self._excel_match(rec, q):           # Search All
                continue
            if coop_only and not self._has_positive_coop(rec):
                continue
            if not passes_filters(rec):                 # Header filters
                continue
            vis.append(i)
        return vis


    def _excel_update_header_checkbox(self) -> None:
        vis = self._excel_visible_indices()
        if not vis:
            self._excel_header_check_state = Qt.Unchecked
        else:
            checked = sum(1 for i in vis if self._excel_row_keys[i] in self._excel_checked_keys)
            if checked == 0:
                self._excel_header_check_state = Qt.Unchecked
            elif checked == len(vis):
                self._excel_header_check_state = Qt.Checked
            else:
                self._excel_header_check_state = Qt.PartiallyChecked
        header_item = self.tree.horizontalHeaderItem(0)
        if header_item is None:
            header_item = QTableWidgetItem("☐")
            self.tree.setHorizontalHeaderItem(0, header_item)
        header_item.setText("☑" if self._excel_header_check_state == Qt.Checked
                            else "◪" if self._excel_header_check_state == Qt.PartiallyChecked else "☐")

    def _excel_on_click(self, row: int, col: int) -> None:
        """Toggle the checkmark in the CHK column for the clicked row."""
        # Find the CHK column (fallback to 0 if not found)
        try:
            chk_col = self.tree._columns.index("CHK")
        except Exception:
            chk_col = 0

        if col != chk_col:
            return

        # Ensure selection set exists
        if not hasattr(self, "_excel_checked_keys"):
            self._excel_checked_keys = set()

        idx = self._excel_iid_to_index.get(row)
        if idx is None:
            return

        key = self._excel_row_keys[idx]
        is_checked = key in self._excel_checked_keys

        # Toggle
        if is_checked:
            self._excel_checked_keys.remove(key)
            new_text = "☐"
        else:
            self._excel_checked_keys.add(key)
            new_text = "☑"

        cell = self.tree.item(row, chk_col)
        if cell is None:
            cell = QTableWidgetItem(new_text)
            cell.setTextAlignment(Qt.AlignCenter)
            cell.setFlags(cell.flags() & ~Qt.ItemIsEditable)
            self.tree.setItem(row, chk_col, cell)
        else:
            cell.setText(new_text)

        self._excel_update_header_checkbox()
    

    def _excel_toggle_all_visible(self):
        vis = self._excel_visible_indices()
        if not vis:
            return
        all_visible_checked = all(self._excel_row_keys[i] in self._excel_checked_keys for i in vis)
        if all_visible_checked:
            for i in vis:
                self._excel_checked_keys.discard(self._excel_row_keys[i])
        else:
            for i in vis:
                self._excel_checked_keys.add(self._excel_row_keys[i])
        self._excel_refresh_table()


    def _excel_refresh_table(self):
        """Rebuild the Excel table view incrementally (prevents UI freeze on large data)."""
        if not hasattr(self, "tree") or not hasattr(self.tree, "_columns"):
            return

        self._excel_widths_tuned = False    

        # Reset
        self.tree.setRowCount(0)
        self._excel_iid_to_index = {}

        cols = tuple(getattr(self.tree, "_columns", ()))
        visible = self._excel_visible_indices()

        # Precompute values for each visible row
        vals_rows = []
        for i in visible:
            rec = self.preview_rows[i]
            key = self._excel_row_keys[i]
            checked = "☑" if key in self._excel_checked_keys else "☐"

            row_vals = []
            for col in cols:
                if col == "CHK":
                    row_vals.append(checked)
                elif col == "Q":
                    row_vals.append(self.preview_qty[i])
                else:
                    row_vals.append(self._excel_legacy_cell(rec, col))

            vals_rows.append((i, row_vals))  # keep original index for iid map

        # remove rows that are visually empty (ignore CHK and Q)
        def _is_visibly_empty(row_vals):
            def _clean(v):
                if v is None:
                    return ""
                s = str(v).strip()
                return "" if s.lower() in ("nan", "none", "null", "nat") else s
            cleaned = [_clean(v) for v in row_vals]
            data_cells = [cleaned[i] for i, name in enumerate(cols) if name not in ("CHK", "Q")]
            return (not data_cells) or all(x == "" for x in data_cells)

        vals_rows = [(i, v) for (i, v) in vals_rows if not _is_visibly_empty(v)]

        # Start batched fill
        self._excel_begin_table_build(vals_rows)
        self._tune_excel_column_widths()


    def _map_excel_row_to_legacy(self, row_dict, fresh_on=True):
        """
        Transform a raw Excel row (as a dict of {header: value}) into the *legacy* schema
        that the rest of the app expects: SECTION, BARCODE, BRAND, ITEM, REG, PROMO, START, END, COOP.

        If fresh_on=True, map Fresh columns:
          - BARCODE <= PLU  (leave empty if PLU missing but Barcode exists)
          - BRAND   <= Arabic Description
          - ITEM    <= English Description
          - COOP    <= UOM (auto-detected)
        Otherwise, pass through legacy columns as-is.

        row_dict: dict with keys matching the Excel headers exactly as your import produced.
        Returns: dict with the legacy keys.
        """

        # --- Normalization helpers ------------------------------------------------
        def _get_first_existing(keys):
            """Return (key, value) for the first key present in row_dict with a non-empty value."""
            for k in keys:
                if k in row_dict:
                    v = row_dict.get(k)
                    if v is not None and str(v).strip() != "":
                        return k, v
            return None, ""

        def _norm(v):
            return "" if v is None else str(v).strip()

        # --- When Fresh is OFF: just map passthrough using common legacy header variants
        if not fresh_on:
            section = _norm(_get_first_existing(("SECTION", "Section", "Dept", "DEPT"))[1] if _get_first_existing(("SECTION", "Section", "Dept", "DEPT")) else "")
            barcode = _norm(_get_first_existing(("BARCODE", "Barcode", "EAN", "UPC"))[1] if _get_first_existing(("BARCODE", "Barcode", "EAN", "UPC")) else "")
            brand   = _norm(_get_first_existing(("BRAND", "Brand"))[1] if _get_first_existing(("BRAND", "Brand")) else "")
            item    = _norm(_get_first_existing(("ITEM", "Item", "Description", "Desc"))[1] if _get_first_existing(("ITEM", "Item", "Description", "Desc")) else "")
            reg     = _norm(_get_first_existing(("REG", "Regular", "Price", "REGULAR"))[1] if _get_first_existing(("REG", "Regular", "Price", "REGULAR")) else "")
            promo   = _norm(_get_first_existing(("PROMO", "Promo", "Offer"))[1] if _get_first_existing(("PROMO", "Promo", "Offer")) else "")
            start   = _norm(_get_first_existing(("START", "Start", "From", "START_DATE"))[1] if _get_first_existing(("START", "Start", "From", "START_DATE")) else "")
            end     = _norm(_get_first_existing(("END", "End", "To", "END_DATE"))[1] if _get_first_existing(("END", "End", "To", "END_DATE")) else "")
            coop    = _norm(_get_first_existing(("COOP", "Coop", "UCM", "CONTRIB"))[1] if _get_first_existing(("COOP", "Coop", "UCM", "CONTRIB")) else "")
            return {
                "SECTION": section, "BARCODE": barcode, "BRAND": brand, "ITEM": item,
                "REG": reg, "PROMO": promo, "START": start, "END": end, "COOP": coop
            }

        # --- Fresh is ON: remap columns as requested ------------------------------
        # 1) BARCODE <= PLU  (if PLU missing but Barcode exists → BARCODE empty)
        # Common Fresh header variants to look for:
        plu_keys = ("PLU", "Plu", "PLU Number", "PLU_NO", "PLU No", "PLU#")
        barcode_keys = ("Barcode", "BARCODE", "EAN", "UPC", "CODE")
        arabic_keys = ("Arabic Description", "ARABIC DESCRIPTION", "Arabic_Desc", "Arabic", "Brand Arabic")
        english_keys = ("English Description", "ENGLISH DESCRIPTION", "Description English", "English", "Item English")
        # UOM detection: we *force* a column whose values look like kg/kgm/pack/packet/pck/cartoon/ctn/etc
        # try friendly headers first, then heuristics on the values
        uom_header_candidates = ("UOM", "Unit", "Unit of Measure", "UnitOfMeasure", "Measure", "U/M")

        # Grab PLU
        _, plu_value = _get_first_existing(plu_keys) if _get_first_existing(plu_keys) else (None, "")
        plu_value = _norm(plu_value)

        # If PLU missing but Barcode exists, we leave BARCODE empty string per your rule.
        has_barcode_any = _get_first_existing(barcode_keys) is not None

        # Grab Arabic / English
        _, ar_value = _get_first_existing(arabic_keys) if _get_first_existing(arabic_keys) else (None, "")
        ar_value = _norm(ar_value)
        _, en_value = _get_first_existing(english_keys) if _get_first_existing(english_keys) else (None, "")
        en_value = _norm(en_value)

        # Detect UOM column: prefer explicit header first
        uom_value = ""
        hit = _get_first_existing(uom_header_candidates)
        if hit:
            uom_value = _norm(hit[1])
        else:
            # Heuristic scan: pick the column whose *value* looks like a UOM keyword
            # We look across current row only (safe & cheap); your sheet usually uses a dedicated UOM column.
            tokens = ("kg", "kgm", "pck", "pack", "packet", "pkt", "ctn", "carton", "cartoon", "pcs", "piece")
            for k, v in row_dict.items():
                if v is None:
                    continue
                val = str(v).strip().lower()
                if any(t in val for t in tokens):
                    uom_value = str(v).strip()
                    break  # first reasonable match wins

        # Other legacy fields keep same behavior as before
        section = _norm(_get_first_existing(("SECTION", "Section", "Dept", "DEPT"))[1] if _get_first_existing(("SECTION", "Section", "Dept", "DEPT")) else "")
        reg     = _norm(_get_first_existing(("REG", "Regular", "Price", "REGULAR"))[1] if _get_first_existing(("REG", "Regular", "Price", "REGULAR")) else "")
        promo   = _norm(_get_first_existing(("PROMO", "Promo", "Offer"))[1] if _get_first_existing(("PROMO", "Promo", "Offer")) else "")
        start   = _norm(_get_first_existing(("START", "Start", "From", "START_DATE"))[1] if _get_first_existing(("START", "Start", "From", "START_DATE")) else "")
        end     = _norm(_get_first_existing(("END", "End", "To", "END_DATE"))[1] if _get_first_existing(("END", "End", "To", "END_DATE")) else "")

        legacy_barcode = plu_value if plu_value else ("" if has_barcode_any else "")  # explicit per rule
        legacy_brand   = ar_value
        legacy_item    = en_value
        legacy_coop    = uom_value

        return {
            "SECTION": section,
            "BARCODE": legacy_barcode,
            "BRAND":   legacy_brand,
            "ITEM":    legacy_item,
            "REG":     reg,
            "PROMO":   promo,
            "START":   start,
            "END":     end,
            "COOP":    legacy_coop,
        }
    



    # at top of file once if missing:
    # import time

    def _excel_begin_table_build(self, vals_rows):
        """Kick off batched insertion using a 0ms timer (adaptive batch size)."""
        self._excel_vals_rows = vals_rows or []
        self._excel_batch_pos = 0

        # Create timer once
        if not hasattr(self, "_excel_batch_timer"):
            self._excel_batch_timer = QTimer(self)
            self._excel_batch_timer.setInterval(0)
            self._excel_batch_timer.timeout.connect(self._excel_fill_batch)

        # Adaptive params (create once)
        if not hasattr(self, "_excel_batch_size") or not isinstance(self._excel_batch_size, int):
            self._excel_batch_size = 500                     # start point
        if not hasattr(self, "_excel_batch_target_ms"):
            self._excel_batch_target_ms = 14.0               # ~60fps budget
        if not hasattr(self, "_excel_batch_min"):
            self._excel_batch_min = 100
        if not hasattr(self, "_excel_batch_max"):
            self._excel_batch_max = 5000

        self._excel_batch_timer.start()  # header checkbox finalizes on completion


    def _excel_fill_batch(self):
        """Insert the next batch and auto-tune batch size based on render time.
        Also installs +/- qty widgets once the final batch is done.
        """
        total = len(getattr(self, "_excel_vals_rows", []))
        pos = getattr(self, "_excel_batch_pos", 0)

        # --- If all rows are already inserted, finalize and attach qty widgets ---
        if pos >= total:
            if hasattr(self, "_excel_batch_timer"):
                self._excel_batch_timer.stop()
            # Update header checkbox state and then decorate Q cells with +/-.
            self._excel_update_header_checkbox()
            try:
                self._excel_apply_qty_buttons()  # <-- ensures widgets appear after rows exist
            except Exception:
                pass
            return

        # take a batch
        size = int(getattr(self, "_excel_batch_size", 500))
        end = min(pos + size, total)
        chunk = self._excel_vals_rows[pos:end]

        # measure
        try:
            import time
            t0 = time.perf_counter()
        except Exception:
            t0 = None

        # find Q column index once for this batch
        cols = tuple(getattr(self.tree, "_columns", ()))
        q_col = cols.index("Q") if "Q" in cols else -1

        # ---- row insertion ----
        for idx, vals in chunk:
            # sanitize cells: turn NaN/None/null/NaT/"" into empty strings
            def _clean_cell(v):
                if v is None:
                    return ""
                s = str(v).strip()
                return "" if s.lower() in ("nan", "none", "null", "nat") else s

            cleaned = [_clean_cell(v) for v in vals]

            # skip row if all *data* cells are empty (ignore CHK and Q)
            data_cells = [cleaned[i] for i, name in enumerate(cols) if name not in ("CHK", "Q")]
            if (not data_cells) or all(x == "" for x in data_cells):
                continue

            row_idx = self.tree.rowCount()
            self.tree.insertRow(row_idx)

            for c, v in enumerate(cleaned):
                it = QTableWidgetItem(v)
                it.setTextAlignment(Qt.AlignCenter)

                # CHK column not editable
                if c == 0:
                    it.setFlags(it.flags() & ~Qt.ItemIsEditable)

                # Make the Q column text big & clear (before we replace with a widget later)
                if q_col >= 0 and c == q_col:
                    f = it.font()
                    f.setPointSize(16)
                    f.setBold(True)
                    it.setFont(f)

                self.tree.setItem(row_idx, c, it)

            # map view row -> original preview_rows index
            self._excel_iid_to_index[row_idx] = idx

            # ensure row tall enough so future qty widget isn’t clipped
            try:
                if self.tree.rowHeight(row_idx) < 34:
                    self.tree.setRowHeight(row_idx, 34)
            except Exception:
                pass

        # advance
        self._excel_batch_pos = end

        # adapt batch size to keep UI responsive
        if t0 is not None:
            dt_ms = (time.perf_counter() - t0) * 1000.0
            tgt = float(getattr(self, "_excel_batch_target_ms", 14.0))
            cur = int(getattr(self, "_excel_batch_size", 500))
            mn = int(getattr(self, "_excel_batch_min", 100))
            mx = int(getattr(self, "_excel_batch_max", 5000))
            if dt_ms < tgt * 0.6:
                self._excel_batch_size = min(mx, int(cur * 1.5) or mn)
            elif dt_ms > tgt * 1.4:
                self._excel_batch_size = max(mn, int(cur * 0.7) or mn)
        # qty widgets will be attached when the final batch completes (see early return above)



    def _edit_qty_excel(self, row, col):
        # Find the current index of the Q column (fallback to 1 if missing)
        try:
            q_col = self.tree._columns.index("Q")
        except ValueError:
            q_col = 1

        if col != q_col:
            return

        item = self.tree.item(row, q_col)
        current = item.text() if item else "1"

        ent = QLineEdit()
        ent.setText(str(current))
        ent.setAlignment(Qt.AlignCenter)
        self.tree.setCellWidget(row, q_col, ent)
        ent.setFocus()

        committed = {"done": False}
        def commit():
            if committed["done"]:
                return
            committed["done"] = True
            try:
                n = max(0, int(float(ent.text())))
            except Exception:
                n = 0
            # write back to table
            if self.tree.item(row, q_col) is None:
                self.tree.setItem(row, q_col, QTableWidgetItem(str(n)))
            else:
                self.tree.item(row, q_col).setText(str(n))
            # write back to model
            idx = self._excel_iid_to_index.get(row)
            if idx is not None:
                self.preview_qty[idx] = n
            self.tree.removeCellWidget(row, q_col)
            self._excel_update_header_checkbox()

        ent.returnPressed.connect(commit)
        ent.editingFinished.connect(commit)


    def _manual_writeback_qty(self, row, n):
        vals = (self.stage.item(row, 1).text(), self.stage.item(row, 2).text(), self.stage.item(row, 3).text())
        key = vals
        for i, r in enumerate(self.staged_rows):
            if (r.get("BARCODE",""), r.get("BRAND",""), r.get("ITEM","")) == key:
                if i < len(self.staged_qty): self.staged_qty[i] = n
                break

    def _on_stage_double_click(self, row, col):
        if col == 0:
            current_item = self.stage.item(row, col)
            current_text = current_item.text() if current_item else "1"

            ent = QLineEdit()
            ent.setText(str(current_text))

            # make the inline QTY editor big & clear
            f = ent.font()
            f.setPointSize(18)   # larger text
            f.setBold(True)      # bold for clarity
            ent.setFont(f)
            ent.setAlignment(Qt.AlignCenter)
            ent.setFixedHeight(36)  # taller editor so text isn't clipped

            self.stage.setCellWidget(row, col, ent)
            ent.setFocus()

            committed = {"done": False}
            def commit():
                if committed["done"]:
                    return
                committed["done"] = True
                try:
                    n = max(0, int(float(ent.text() or 1)))
                except Exception:
                    n = 1
                if self.stage.item(row, col) is None:
                    self.stage.setItem(row, col, QTableWidgetItem(str(n)))
                else:
                    self.stage.item(row, col).setText(str(n))
                self._manual_writeback_qty(row, n)
                self.stage.removeCellWidget(row, col)

            ent.returnPressed.connect(commit)
            ent.editingFinished.connect(commit)
            return

        if not hasattr(self, "mform") or self.mform is None:
            return

        def cell_text(c: int) -> str:
            it = self.stage.item(row, c)
            return it.text() if it else ""

        r = {
            "BARCODE": cell_text(1),
            "BRAND": cell_text(2),
            "ITEM": cell_text(3),
            "REG": cell_text(4),
            "PROMO": cell_text(5),
            "START_DATE": cell_text(6),
            "END_DATE": cell_text(7),
            "COOP": cell_text(8),
        }
        self.mform.fill(r)
        self._editing_from_stage = True
        self._editing_from_stage_key = (r["BARCODE"], r["BRAND"], r["ITEM"])



    def _collect_from_tree(self)->List[Dict[str,str]]:
        if self._current_gen_source == "excel":
            out: List[Dict[str,str]] = []
            if self._excel_checked_keys:
                indices = [i for i in range(len(self.preview_rows))
                           if self._excel_row_keys[i] in self._excel_checked_keys]
            else:
                vis = self._excel_visible_indices()
                indices = vis if vis else list(range(len(self.preview_rows)))
            for i in indices:
                rec = self.preview_rows[i]
                try:
                    q = int(float(self.preview_qty[i] or 1))
                except Exception:
                    q = 1
                q = max(1, q)
                base = {
                    # legacy fields (keep)
                    "BARCODE": rec.get("BARCODE",""),
                    "BRAND":   rec.get("BRAND",""),
                    "ITEM":    rec.get("ITEM",""),
                    "REG":     price_text(rec.get("REG","")),
                    "PROMO":   price_text(rec.get("PROMO","")),
                    "START_DATE": rec.get("START_DATE",""),
                    "END_DATE":   rec.get("END_DATE",""),
                    "COOP":    price_text(rec.get("COOP","")),
                    "SECTION": rec.get("SECTION",""),

                    # fresh fields (new)
                    "PLU":                 rec.get("PLU",""),
                    "ARABIC_DESCRIPTION":  rec.get("ARABIC_DESCRIPTION",""),
                    "ENGLISH_DESCRIPTION": rec.get("ENGLISH_DESCRIPTION",""),
                    "REGULAR_PRICE":       price_text(rec.get("REGULAR_PRICE","")),
                    "PROMO_PRICE":         price_text(rec.get("PROMO_PRICE","")),
                }

                # Bridge: when Fresh mode is ON, mirror Fresh → legacy if legacy empty
                if FRESH_SECTION_ACTIVE:
                    if not base["ITEM"] and base["ENGLISH_DESCRIPTION"]:
                        base["ITEM"] = base["ENGLISH_DESCRIPTION"]
                    if not base["REG"] and base["REGULAR_PRICE"]:
                        base["REG"] = base["REGULAR_PRICE"]
                    if not base["PROMO"] and base["PROMO_PRICE"]:
                        base["PROMO"] = base["PROMO_PRICE"]

                    if not base["BRAND"] and base["ARABIC_DESCRIPTION"]:
                        base["BRAND"] = base["ARABIC_DESCRIPTION"]
    

                for _ in range(q):
                    out.append(dict(base))
            return out

        out: List[Dict[str,str]] = []
        for row in range(self.stage.rowCount()):
            try: q = max(0, int(float(self.stage.item(row, 0).text() or 1)))
            except Exception:
                q = 0
            if q <= 0:
                continue
            r = {
                "BARCODE": self.stage.item(row, 1).text(),
                "BRAND": self.stage.item(row, 2).text(),
                "ITEM": self.stage.item(row, 3).text(),
                "REG": price_text(self.stage.item(row, 4).text()),
                "PROMO": price_text(self.stage.item(row, 5).text()),
                "START_DATE": self.stage.item(row, 6).text(),
                "END_DATE": self.stage.item(row, 7).text(),
                "COOP": price_text(self.stage.item(row, 8).text()),
                "SECTION": ""
            }
            for _ in range(q):
                out.append(dict(r))
        if not out and self.staged_rows:
            for i, rec in enumerate(self.staged_rows):
                q = self.staged_qty[i] if i < len(self.staged_qty) else 1
                if q <= 0:
                    continue
                r = {
                    "BARCODE": rec.get("BARCODE", ""),
                    "BRAND": rec.get("BRAND", ""),
                    "ITEM": rec.get("ITEM", ""),
                    "REG": price_text(rec.get("REG", "")),
                    "PROMO": price_text(rec.get("PROMO", "")),
                    "START_DATE": rec.get("START_DATE", ""),
                    "END_DATE": rec.get("END_DATE", ""),
                    "COOP": price_text(rec.get("COOP", "")),
                    "SECTION": rec.get("SECTION", ""),
                }
                for _ in range(q):
                    out.append(dict(r))
        return out

    def _on_generate(self, source):
        rows = self._collect_from_tree()
        if not rows:
            QMessageBox.information(self, "Generate", "No items with Qty > 0 (or none checked).")
            return
        if not self.selected_template:
            QMessageBox.information(self, "Generate", "Please pick a template button first.")
            return
        tpl = self.selected_template or {}
        has_positions = isinstance(tpl.get("positions"), dict) and bool(tpl["positions"])
        tmp = tempfile.gettempdir()
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        out = os.path.join(tmp, f"labels_{self.selected_template_name or 'template'}_{ts}.pdf")
        if has_positions:
            try:
                render_page_JSON(out, tpl, rows)
            except Exception as e:
                QMessageBox.critical(self, "Render Error", f"Could not render template:\n{e}")
                return
            open_file(out)
            return
        QMessageBox.warning(
            self,
            "Template Not Supported",
            f"Template '{self.selected_template_name or ''}' has no ''positions'. "
            "Paste a Template Builder JSON with 'positions', 'styles', and 'active_headers'."
        )

    def _build_manual(self):
        self.clear()
        self._preferred_size = "large"
        self.set_large()

        top = QFrame(objectName="Card")
        top_layout = QHBoxLayout(top)
        self.main_layout.addWidget(top)
        back = QLabel("‹ Back")
        back.setObjectName("Back")
        back.setCursor(QCursor(Qt.PointingHandCursor))
        top_layout.addWidget(back)
        back.mousePressEvent = lambda e: self._build_home()
        title = QLabel("Manual Entry")
        title.setObjectName("Title")
        top_layout.addWidget(title)

        body = QFrame(objectName="Card")
        body_layout = QVBoxLayout(body)
        self.main_layout.addWidget(body)

        # --- Search row ---
        search_wrap = QWidget()
        search_layout = QHBoxLayout(search_wrap)
        search_layout.addStretch()
        search_layout.addWidget(QLabel("Search", objectName="Small"))

        self.s_edit = QLineEdit()
        self.s_edit.setFont(QFont("Arial", 16))
        self.s_edit.setFixedHeight(36)
        self.s_edit.setMaximumWidth(420)
        search_layout.addWidget(self.s_edit)

        clear_search = QPushButton("✖")
        clear_search.setFixedWidth(24)
        search_layout.addWidget(clear_search)

        paste_btn = QPushButton("📋")
        paste_btn.setFixedWidth(30)
        paste_btn.clicked.connect(self._multi_paste_clipboard)
        search_layout.addWidget(paste_btn)

        start_btn = QPushButton("©")
        start_btn.setFixedWidth(30)
        start_btn.clicked.connect(self._multi_start_stepthrough)
        search_layout.addWidget(start_btn)

        arrow_btn = QPushButton("➤")
        arrow_btn.setFixedWidth(30)
        arrow_btn.clicked.connect(self._open_paste_panel)
        search_layout.addWidget(arrow_btn)
        # ---- Live search wiring (ONLY on Manual screen) ----
        # If you already have a Debouncer class, keep this; otherwise we can wire directly.
        if not hasattr(self, "_debounce_manual_search") or self._debounce_manual_search is None:
            self._debounce_manual_search = Debouncer(150, self._on_search_typing, self)

        # Text change -> live search
        # Safely disconnect only our slot if it was connected
        try:
            self.s_edit.textChanged.connect(self._on_search_typing)
        except (TypeError, RuntimeError):
            pass

        self.s_edit.textChanged.connect(lambda t: self._debounce_manual_search.call(t))

        # Clear button should also clear results via the same path
        clear_search.clicked.connect(
            lambda: (self.s_edit.setText(""), self._debounce_manual_search.call(""))
        )


        # ⬇️ Add this block (Strict Manual toggle) —
        self._strict_btn = QPushButton()
        self._strict_btn.setObjectName("FreshToggle")  # reuse the green toggle style
        is_on = bool(getattr(self, "_strict_manual_on", True))
        self._strict_btn.setText("Strict Manual: ON" if is_on else "Strict Manual: OFF")
        self._strict_btn.setProperty("active", is_on)
        self._strict_btn.setFixedHeight(32)
        self._strict_btn.setMinimumWidth(170)
        self._strict_btn.clicked.connect(self._toggle_strict_manual)
        # refresh style so the green/gray state shows immediately
        self._strict_btn.style().unpolish(self._strict_btn)
        self._strict_btn.style().polish(self._strict_btn)

        # place it on the left side of the search row
        search_layout.addWidget(self._strict_btn)


        search_layout.addStretch()
        body_layout.addWidget(search_wrap)

        # --- Debounced search wiring (NEW) ---
        # create once, reuse across screen rebuilds
        if not hasattr(self, "_debounce_manual_search") or self._debounce_manual_search is None:
            self._debounce_manual_search = Debouncer(150, self._on_search_typing, self)

        # on text typing -> debounce
        self.s_edit.textChanged.connect(lambda t: self._debounce_manual_search.call(t))
        # on clear button -> also go through debouncer
        clear_search.clicked.connect(lambda: (self.s_edit.setText(""), self._debounce_manual_search.call("")))

        # Keep Enter behavior exactly the same
        self.s_edit.returnPressed.connect(self._on_search_enter)

        # --- Live hits table ---
        self._hits_columns = (
            "BARCODE", "BRAND", "ITEM", "REG", "PROMO",
            "START_DATE", "END_DATE", "COOP",
        )
        self.hits = QTableWidget(0, len(self._hits_columns))
        self.hits.setObjectName("LiveHits")
        self.hits.setHorizontalHeaderLabels(self._hits_columns)

        header = self.hits.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)   # default for all

        # widen ITEM
        item_col = self._hits_columns.index("ITEM")
        header.setSectionResizeMode(item_col, QHeaderView.Interactive)
        self.hits.setColumnWidth(item_col, 200)

        # shrink REG / PROMO / COOP
        for name, w in (("REG", 70), ("PROMO", 70), ("COOP", 70)):
            col = self._hits_columns.index(name)
            header.setSectionResizeMode(col, QHeaderView.Interactive)
            self.hits.setColumnWidth(col, w)

        self.hits.setAlternatingRowColors(True)
        self.hits.setMinimumHeight(80)
        self.hits.verticalHeader().setDefaultSectionSize(18)
        body_layout.addWidget(self.hits)
        self.hits.cellDoubleClicked.connect(lambda r, c: self._stack_selected())

        # --- Manual form ---
        formwrap = QFrame(objectName="Card")
        formwrap_layout = QVBoxLayout(formwrap)
        body_layout.addWidget(formwrap)

        self.mform = ManualForm(
            formwrap,
            self._manual_add,
            lambda: self._build_generate(source="manual"),
            lambda: (self.mform.clear(), self.s_edit.setText(""), self._on_search_typing(self.s_edit.text())),
            self._clear_stage,
        )
        formwrap_layout.addWidget(self.mform)

        body_layout.addWidget(QLabel("Your List", objectName="Small"))

        # --- Staged list table ---
        # --- Staged list table ---
        self.stage = FilterableTable(
            ("QTY", "BARCODE", "BRAND", "ITEM", "REG", "PROMO", "START", "END", "COOP"),
            enable_filters=False
        )
        self.stage.setObjectName("StageTable")
        self.stage.setMinimumHeight(700)
        body_layout.addWidget(self.stage)
        self.stage.cellDoubleClicked.connect(self._on_stage_double_click)

        # Reset width-tuned flag for this table build
        self._stage_widths_tuned = False

        # === Preserve staged data across revisits to Manual; only init if missing ===
        if not hasattr(self, "staged_rows"):
            self.staged_rows = []
        if not hasattr(self, "staged_qty"):
            self.staged_qty = []

        self._manual_search_edit = QLineEdit()
        self._manual_refresh_stage_table()


        # Autofill wiring as before
        for widget, field in ((self.mform.e_bar, "BARCODE"), (self.mform.e_brand, "BRAND"), (self.mform.e_item, "ITEM")):
            widget.textChanged.connect(lambda t, f=field: self._manual_field_typing(f))
            widget.returnPressed.connect(lambda f=field: self._manual_field_enter(f))



    def _refresh_db_cache(self):
        """
        Build a fast in-memory search index from the CSV once.
        Auto-refreshes when the CSV file mtime changes.
        """
        try:
            mtime = os.path.getmtime(_db_path())
        except Exception:
            mtime = None

        rows = load_db_rows()  # full CSV → list[dict]
        cache = []

        searchable = set(self._searchable_fields())

        # Prefer latest rows first (keep behavior of reversed(load_db_rows()))
        for r in reversed(rows):
            bc = (r.get("BARCODE", "") or "").strip().lower()
            br = (r.get("BRAND", "") or "").strip().lower()
            it = (r.get("ITEM", "") or "").strip().lower()

            fields_lower = {}
            for h in searchable:
                v = r.get(h, "")
                if v is None:
                    continue
                s = str(v).strip().lower()
                if s:
                    fields_lower[h] = s

            # token sets for BRAND/ITEM matching bucket
            tokens_brand = set(re.findall(r"[A-Za-z0-9]+", br))
            tokens_item  = set(re.findall(r"[A-Za-z0-9]+", it))

            cache.append({
                "row": r,
                "bc": bc,
                "br": br,
                "it": it,
                "fields": fields_lower,
                "tok_brand": tokens_brand,
                "tok_item": tokens_item,
            })

        self._db_cache = cache
        self._db_cache_mtime = mtime


    def _matches_for_query(self, q: str) -> List[Dict[str, str]]:
        """
        Cached, ranked search (exact barcode, barcode suffix, exact any, token-in-brand/item, substr-other).
        Returns de-duplicated results (newest first) from CSV cache.
        """
        q = self._normalize_search_text((q or "").strip())
        if not q:
            return []

        ql = q.lower()
        is_digit = q.isdigit()

        # Refresh cache if file changed or not built
        try:
            cur_mtime = os.path.getmtime(_db_path())
        except Exception:
            cur_mtime = None
        if not hasattr(self, "_db_cache") or getattr(self, "_db_cache_mtime", None) != cur_mtime:
            self._refresh_db_cache()

        cache = getattr(self, "_db_cache", []) or []
        if not cache:
            return []

        exact_barcode: List[Dict[str, str]] = []
        suffix_barcode: List[Dict[str, str]] = []
        exact_any: List[Dict[str, str]] = []
        substr_brand_item: List[Dict[str, str]] = []
        substr_other: List[Dict[str, str]] = []
        seen = set()

        def key_of(r):
            return (
                (r.get("BARCODE", "") or "").strip(),
                (r.get("BRAND", "") or "").strip(),
                (r.get("ITEM", "") or "").strip(),
            )

        def add_once(bucket, r):
            k = key_of(r)
            if k not in seen:
                seen.add(k)
                bucket.append(r)

        for e in cache:
            r = e["row"]
            bc = e["bc"]; br = e["br"]; it = e["it"]

            # 1) exact barcode
            if bc and ql == bc:
                add_once(exact_barcode, r)
                continue

            # 2) numeric barcode suffix (≥5 digits)
            if is_digit and len(q) >= 5 and bc.endswith(ql):
                add_once(suffix_barcode, r)
                continue

            # 3) exact match in ANY searchable field
            matched = False
            for _h, val in e["fields"].items():
                if ql == val:
                    add_once(exact_any, r)
                    matched = True
                    break
            if matched:
                continue

            # 4) token present in BRAND / ITEM
            if (ql in e["tok_brand"]) or (ql in e["tok_item"]):
                add_once(substr_brand_item, r)
                continue

            # 5) substring in other searchable fields (incl. BRAND/ITEM fallback)
            for _h, val in e["fields"].items():
                if ql in val:
                    add_once(substr_other, r)
                    break

        # Concatenate buckets by rank
        out = exact_barcode + suffix_barcode + exact_any + substr_brand_item + substr_other
        return out



    def _is_strong_match(self, q: str, rec: Dict[str, str], field: Optional[str] = None) -> bool:
        qn = self._normalize_search_text((q or "").strip())
        if not qn: return False
        bc = (rec.get("BARCODE", "") or "").strip()
        br = (rec.get("BRAND", "") or "").strip()
        it = (rec.get("ITEM", "") or "").strip()
        ql = qn.lower(); bcl = bc.lower(); brl = br.lower(); itl = it.lower()
        if field == "BARCODE": return bool(bc) and (ql == bcl)
        if field == "BRAND":   return bool(br) and (ql == brl)
        if field == "ITEM":    return bool(it) and (ql == itl)
        return (bool(bc) and ql == bcl) or (bool(br) and ql == brl) or (bool(it) and ql == itl)

    def _unique_live_match(self, q: str, field: Optional[str] = None) -> Optional[Dict[str, str]]:
        m = self._matches_for_query(q)
        if len(m) != 1: return None
        rec = m[0]
        return rec if self._is_strong_match(q, rec, field=field) else None

    def _normalize_search_text(self, s: str) -> str:
        """Lowercase, collapse spaces; keeps digits unchanged for barcode suffix match."""
        s = (s or "").strip()
        s = re.sub(r"\s+", " ", s)
        return s

    def _on_search_typing(self, text: str):
        """Debounced live search → fill self.hits table."""
        if not hasattr(self, "hits") or self.hits is None:
            return
        rows = self._matches_for_query(text)
        cols = self._hits_columns
        self.hits.setRowCount(0)
        for r in rows[:500]:  # cap UI
            i = self.hits.rowCount()
            self.hits.insertRow(i)
            values = (
                r.get("BARCODE", ""),
                r.get("BRAND", ""),
                r.get("ITEM", ""),
                r.get("REG", ""),
                r.get("PROMO", ""),
                r.get("START_DATE", ""),
                r.get("END_DATE", ""),
                r.get("COOP", ""),
            )
            for c, v in enumerate(values):
                it = QTableWidgetItem(str(v or ""))
                it.setTextAlignment(Qt.AlignCenter)
                self.hits.setItem(i, c, it)

    def _on_search_enter(self):
        if not self._search_enter_armed:
            txt = self._normalize_search_text(self.s_edit.text().strip())
            matches = self._matches_for_query(txt)
            rec = self._unique_live_match(txt) or (matches[0] if matches else None)
            if rec:
                self._autofill_once(rec)
                self._search_enter_armed = True
        else:
            self._manual_add(self.mform.values())
            self._search_enter_armed = False

    def _manual_field_typing(self, field: str):
        if field == "BARCODE":
            q = self._normalize_search_text(self.mform.e_bar.text())
        elif field == "BRAND":
            q = self.mform.e_brand.text()
        else:
            q = self.mform.e_item.text()
        self._populate_hits(q)
        self._manual_enter_armed = False
        rec = self._unique_live_match(q, field=field)
        if rec:
            self._autofill_once(rec)

    def _manual_field_enter(self, field: str):
        if not self._manual_enter_armed:
            if field == "BARCODE":
                q = self._normalize_search_text(self.mform.e_bar.text().strip())
            elif field == "BRAND":
                q = self.mform.e_brand.text().strip()
            else:
                q = self.mform.e_item.text().strip()
            matches = self._matches_for_query(q)
            rec = self._unique_live_match(q) or (matches[0] if matches else None)
            if rec:
                self._autofill_once(rec)
                self._manual_enter_armed = True
        else:
            self._manual_add(self.mform.values())
            self._manual_enter_armed = False

    def _populate_hits(self, q: str):
        # Work with the live hits table, not the Excel tree
        if not hasattr(self, "hits") or self.hits is None:
            return

        # Clear when empty query
        self.hits.setRowCount(0)
        q = (q or "").strip()
        if not q:
            return

        # Get matches
        try:
            matches = self._matches_for_query(self._normalize_search_text(q)) or []
        except Exception:
            matches = []

        # Fill the hits table (cap for speed)
        cols = list(getattr(self, "_hits_columns", (
            "BARCODE","BRAND","ITEM","REG","PROMO","START_DATE","END_DATE","COOP"
        )))
        for rec in matches[:100]:
            row_idx = self.hits.rowCount()
            self.hits.insertRow(row_idx)
            for c, key in enumerate(cols):
                val = ""
                try:
                    val = str((rec or {}).get(key, "") or "")
                except Exception:
                    val = ""
                it = QTableWidgetItem(val)
                it.setTextAlignment(Qt.AlignCenter)
                # make ID-ish columns read-only
                if key in ("BARCODE","BRAND","ITEM"):
                    it.setFlags(it.flags() & ~Qt.ItemIsEditable)
                self.hits.setItem(row_idx, c, it)


    def _autofill_once(self, rec: Dict[str, str]):
        key = (rec.get("BARCODE", "").strip(), rec.get("BRAND", "").strip(), rec.get("ITEM", "").strip())
        if self._last_autofill_key == key:
            return
        self._fill_manual_from_record(rec)
        self._last_autofill_key = key

    def _first_hit_record(self) -> Optional[Dict[str, str]]:
        row = self.hits.currentRow()
        if row < 0:
            return None
        bc = self.hits.item(row, 0).text()
        br = self.hits.item(row, 1).text()
        it = self.hits.item(row, 2).text()
        for r in reversed(load_db_rows()):
            if r.get("BARCODE", "") == bc and r.get("BRAND", "") == br and r.get("ITEM", "") == it:
                return r
        return None

    def _fill_manual_from_record(self, rec: Dict[str, str]):
        r = {
            "BARCODE": clean_barcode(rec.get("BARCODE", "")),
            "BRAND": rec.get("BRAND", ""),
            "ITEM": rec.get("ITEM", ""),
            "REG": price_text(rec.get("REG", "")),
            "PROMO": price_text(rec.get("PROMO", "")),
            "START_DATE": date_only(rec.get("START_DATE", "")),
            "END_DATE": date_only(rec.get("END_DATE", "")),
            "COOP": price_text(rec.get("COOP", "")),
        }
        self.mform.fill(r)

    def _stack_selected(self):
        rec = self._first_hit_record()
        if rec:
            self._fill_manual_from_record(rec)
            self._manual_add(self.mform.values())
            self._search_enter_armed = False
            self._manual_enter_armed = False
            self._last_autofill_key = None

    def _searchable_fields(self) -> List[str]:
        """
        Read Header Manager and return fields marked as searchable.
        Falls back to CORE + defaults if config missing.
        """
        try:
            cfg = load_headers_cfg()
            return [k for k, v in (cfg or {}).items() if isinstance(v, dict) and v.get("searchable", True)]
        except Exception:
            # Safe default
            return ["BARCODE", "BRAND", "ITEM", "SECTION", "REG", "PROMO", "START_DATE", "END_DATE", "COOP",
                    "PLU", "ARABIC_DESCRIPTION", "ENGLISH_DESCRIPTION", "REGULAR_PRICE", "PROMO_PRICE"]

    def _manual_add(self, vals: Dict[str, str], clear_form: bool = False, clear_search: bool = False):
        # read current toggle (default True if missing)
        strict_on = bool(getattr(self, "_strict_manual_on", True))

        barcode = vals.get("BARCODE", "").strip()
        brand = vals.get("BRAND", "").strip()
        item = vals.get("ITEM", "").strip()
        reg = price_text(vals.get("REG", ""))
        promo = price_text(vals.get("PROMO", ""))

        # compute "required" set exactly like your current rule
        required_missing = []
        if not barcode: required_missing.append("BARCODE")
        if not brand:   required_missing.append("BRAND")
        if not item:    required_missing.append("ITEM")
        if not (reg or promo): required_missing.append("REG or PROMO")

        # When Strict Manual is ON -> behave exactly as before (block if missing)
        if strict_on and required_missing:
            QMessageBox.warning(self, "Missing fields", "Missing required fields: " + ", ".join(required_missing))
            return

        # Build row (allow blanks if Strict OFF)
        r = {
            "BARCODE": clean_barcode(barcode) if barcode else "",
            "BRAND": brand,
            "ITEM": item,
            "REG": reg,
            "PROMO": promo,
            "START_DATE": date_only(vals.get("START_DATE", "")),
            "END_DATE": date_only(vals.get("END_DATE", "")),
            "COOP": price_text(vals.get("COOP", "")),
            "SECTION": ""
        }

        # If Strict OFF, mark as local-only and DO NOT save to DB
        # (this ensures nothing created while toggle is OFF goes into your DB)
        if not strict_on:
            r["_local_only"] = True

        # Upsert/replace in staged list (same key logic)
        key = lambda d: (d.get("BARCODE", "").strip(), d.get("BRAND", "").strip(), d.get("ITEM", "").strip())
        replaced = False
        for i, existing in enumerate(self.staged_rows):
            if key(existing) == key(r):
                # preserve _local_only if it already existed OR if we're adding in non-strict mode
                local_only = existing.get("_local_only") or r.get("_local_only")
                self.staged_rows[i] = {**existing, **r}
                if local_only:
                    self.staged_rows[i]["_local_only"] = True
                replaced = True
                break
        if not replaced:
            self.staged_rows.append(r)
            self.staged_qty.append(1)

        # Refresh the staged table
        self._manual_refresh_stage_table()

        # Only save to DB when Strict is ON
        if strict_on:
            allrows = [{
                "BARCODE": r["BARCODE"], "BRAND": r["BRAND"], "ITEM": r["ITEM"], "REG": r["REG"], "PROMO": r["PROMO"],
                "START_DATE": r["START_DATE"], "END_DATE": r["END_DATE"], "COOP": r["COOP"], "SECTION": r["SECTION"]
            }]
            upsert_db_rows(allrows)

        # Reset UI state as before
        self.mform.clear()
        self.s_edit.setText("")
        self._on_search_typing(self.s_edit.text())
        self._search_enter_armed = False
        self._manual_enter_armed = False
        self._last_autofill_key = None
        self._editing_from_stage = False
        if self._multi_mode_active:
            self._multi_advance_after_add()


    def _clear_stage(self):
        if not hasattr(self, "stage"):
            self.staged_rows = []
            self.staged_qty = []
            return
        sel_rows = [i.row() for i in self.stage.selectedIndexes()]
        sel_rows = list(set(sel_rows))
        if not sel_rows:
            self.staged_rows = []
            self.staged_qty = []
            self.stage.setRowCount(0)
            return
        selected_keys = set()
        for row in sel_rows:
            vals = (
                self.stage.item(row, 1).text(),
                self.stage.item(row, 2).text(),
                self.stage.item(row, 3).text()
            )
            selected_keys.add(vals)
        new_rows = []
        new_qty = []
        for i, r in enumerate(self.staged_rows):
            k = (r.get("BARCODE", "").strip(), r.get("BRAND", "").strip(), r.get("ITEM", "").strip())
            if k in selected_keys:
                continue
            new_rows.append(r)
            new_qty.append(self.staged_qty[i] if i < len(self.staged_qty) else 1)
        self.staged_rows = new_rows
        self.staged_qty = new_qty
        self._manual_refresh_stage_table()

    def _parse_multi_lines(self, text: str) -> List[Tuple[str, int]]:
        seen = set(); out = []
        for raw in (text or "").splitlines():
            line = raw.strip()
            if not line: continue
            tok, qty = line, 1
            if "," in line:
                parts = [p.strip() for p in line.split(",", 1)]
                tok = parts[0]
                try: qty = max(1, int(float(parts[1])))
                except Exception: qty = 1
            if tok not in seen:
                seen.add(tok); out.append((tok, qty))
        return out

    def _match_token_to_record(self, tok: str) -> Optional[Dict[str, str]]:
        q = self._normalize_search_text(tok)
        matches = self._matches_for_query(q)
        return matches[0] if matches else None

    def _multi_paste_clipboard(self):
        app = QApplication.instance()
        clip = app.clipboard().text()
        self._multi_mode_active = False
        self._multi_found_queue = []
        self._multi_found_qty = []
        self._multi_unfound_tokens = []
        self._multi_index = 0
        self.staged_rows = []
        self.staged_qty = []
        self.stage.setRowCount(0)
        self.hits.setRowCount(0)
        self._last_autofill_key = None
        pairs = self._parse_multi_lines(clip)
        if not pairs:
            QMessageBox.information(self, "Multiple Selection", "Nothing to paste.")
            return
        found = []; found_qty = []; unfound = []
        for tok, q in pairs:
            rec = self._match_token_to_record(tok)
            if rec:
                found.append(rec); found_qty.append(q)
            else:
                unfound.append(tok)
        self._multi_found_queue = found
        self._multi_found_qty = found_qty
        self._multi_unfound_tokens = unfound
        self._multi_index = 0
        self._multi_mode_active = False
        QMessageBox.information(
            self,
            "Multiple Selection",
            f"Pasted: {len(pairs)}\nFound in DB: {len(found)}\nNot found: {len(unfound)}\n\nPress © to start."
        )
        if found:
            self._multi_start_stepthrough()

    def _manual_open_unfound_token(self, tok: str):
        self.mform.clear()
        self.mform.e_bar.setText(tok)

    def _multi_start_stepthrough(self):
        if not self._multi_found_queue and not self._multi_unfound_tokens:
            QMessageBox.information(self, "Multiple Selection", "No pasted list. Click 📋 to paste first.")
            return
        self._multi_mode_active = True
        self._multi_index = 0
        if self._multi_found_queue:
            self._fill_manual_from_record(self._multi_found_queue[0])
        else:
            tok = self._multi_unfound_tokens[0]
            self._manual_open_unfound_token(tok)
            self._multi_mode_active = False

    def _multi_advance_after_add(self):
        if not self._multi_mode_active:
            return
        self._multi_index += 1
        if self._multi_index < len(self._multi_found_queue):
            self._fill_manual_from_record(self._multi_found_queue[self._multi_index])
        else:
            self._multi_mode_active = False
            if self._multi_unfound_tokens:
                tok = self._multi_unfound_tokens[0]
                self._manual_open_unfound_token(tok)

    def _find_original_by_key(self, key_row: Dict[str, str]) -> Optional[Dict[str, str]]:
        barcode = key_row.get("BARCODE", "").strip()
        candidates = self.preview_rows + self.staged_rows
        if barcode:
            for r in candidates:
                if r.get("BARCODE", "").strip() == barcode:
                    return r
        kb = (
            key_row.get("BARCODE", "").strip(),
            key_row.get("BRAND", "").strip(),
            key_row.get("ITEM", "").strip()
        )
        for r in candidates:
            if (
                r.get("BARCODE", "").strip(),
                r.get("BRAND", "").strip(),
                r.get("ITEM", "").strip()
            ) == kb:
                return r
        return None

    def _draw_divider_page(self, canvas_obj, section_name: str):
        _ensure_reportlab()
        W, H = A4
        canvas_obj.setFillColor(black)
        canvas_obj.setFont("Helvetica-Bold", 28)
        canvas_obj.drawCentredString(W / 2.0, H * 0.72, (section_name or "(Unknown)").upper())
        canvas_obj.showPage()

    # --- REPLACE the whole smart_ai_render() with this fixed version ---
    def smart_ai_render(self):
        _ensure_reportlab()
        _ensure_arabic_font()
        if not self.selected_template:
            QMessageBox.information(self, "Smart AI", "Pick a 6-slot template first.")
            return
        tpl = self.selected_template
        has_positions = isinstance(tpl.get("positions"), dict) and bool(tpl["positions"])
        if not has_positions:
            QMessageBox.warning(self, "Smart AI", "Selected template has no 'positions'.")
            return

        positions = _coerce_int_keys(tpl.get("positions", {}))
        slots = _grid_order_positions(positions)
        per_page = len(slots) if slots else 6

        # ✅ FIX: determine layout so clipping is correct (prevents half-page cutoff)
        layout_mode = _infer_layout_mode(positions)

        expanded_rows = self._collect_from_tree()
        if not expanded_rows:
            QMessageBox.information(self, "Smart AI", "No items (QTY > 0) in the list.")
            return

        def resolve_section(rec: Dict[str, str]) -> str:
            sec = rec.get("SECTION", "").strip()
            if not sec:
                orig = self._find_original_by_key(rec)
                sec = (orig.get("SECTION", "") if orig else "").strip()
            return sec if sec else "(Unknown)"

        section_order: List[str] = []
        by_section: Dict[str, List[Dict[str, str]]] = {}
        for r in expanded_rows:
            sec = resolve_section(r)
            if sec not in by_section:
                by_section[sec] = []
                section_order.append(sec)
            by_section[sec].append(r)

        pages = []
        page = []
        used = 0
        for idx, sec in enumerate(section_order):
            rows = by_section[sec]
            for rec in rows:
                if used >= per_page:
                    pages.append(page); page = []; used = 0
                page.append(("item", rec, sec)); used += 1
            if idx < len(section_order) - 1:
                next_sec = section_order[idx + 1]
                remaining = per_page - used
                if remaining >= 2:
                    page.append(("divider", sec, next_sec)); used += 1
                elif remaining == 1:
                    page.append(("divider", sec, next_sec)); pages.append(page); page = []; used = 0
                else:
                    if page: pages.append(page)
                    page = [("divider", sec, next_sec)]; used = 1
        if page: pages.append(page)

        active = tpl.get("active_headers", {}) or {}
        styles = tpl.get("styles", {}) or {}
        tmp = tempfile.gettempdir()
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        out = os.path.join(tmp, f"smartai_{self.selected_template_name or 'template'}_{ts}.pdf")
        c = pdfgen_canvas.Canvas(out, pagesize=A4)

        def _slot_anchor_xy(header_map: Dict[str, dict]) -> Tuple[float, float, str]:
            for key in ("ITEM", "PROMO", "BRAND", "REG"):
                pos = header_map.get(key)
                if isinstance(pos, dict):
                    try:
                        return float(pos.get("x", 0.0)), float(pos.get("y", 0.0)), key
                    except Exception:
                        pass
            for key, pos in header_map.items():
                if isinstance(pos, dict):
                    try:
                        return float(pos.get("x", 0.0)), float(pos.get("y", 0.0)), key
                    except Exception:
                        pass
            return 105.0, 148.0, ""

        def _page_header_text(entries: List[Tuple[str, object, object]]) -> str:
            seen_order: List[str] = []
            def _add(s: str):
                if s not in seen_order:
                    seen_order.append(s)
            for kind, a, b in entries:
                if kind == "item": _add(str(b))
            for kind, a, b in entries:
                if kind == "divider":
                    _add(str(a)); _add(str(b))
            return " / ".join(seen_order)

        for entries in pages:
            header_txt = _page_header_text(entries)
            if header_txt:
                c.setFont("Helvetica-Bold", 11); c.setFillColor(black)
                c.drawCentredString(A4[0] / 2.0, A4[1] - 8 * mm, header_txt)

            for slot_index, (_row_idx, _side_idx, header_map) in enumerate(slots):
                if slot_index >= len(entries):
                    break

                kind, a, b = entries[slot_index]

                c.saveState()
                try:
                    # ✅ FIX: use the correct side index and the computed layout
                    _safe_clip_to_layout(c, _side_idx, layout_mode, margin_mm=2.0)
                except Exception:
                    pass

                if kind == "item":
                    rec = a  # type: ignore[assignment]
                    for hname, pos in header_map.items():
                        if active and (hname in active) and (not active[hname]):
                            continue
                        if not pos.get("visible", True):
                            continue

                        x = float(pos.get("x", 0.0))
                        y = float(pos.get("y", 0.0))
                        align = (pos.get("align", "left") or "left").strip().lower()
                        st = styles.get(
                            hname,
                            {
                                "family": "Helvetica",
                                "size": 12,
                                "bold": False,
                                "italic": False,
                                "underline": False,
                                "strike": False,
                                "color": "#000000",
                            },
                        )


                        # Force Arabic field to use the registered Arabic font
                        # (we'll define _ARABIC_FONT_NAME in the next step)
                        if hname == "ARABIC_DESCRIPTION":
                            st = {**st, "family": _ARABIC_FONT_NAME}

                        text = _value_for_header(rec, hname)

                        # width caps from template (optional)
                        max_w_mm = float(pos.get("max_w_mm", 0) or 0)
                        margin_mm_val = float(pos.get("margin_mm", 2.0))

                        # choose side-aware width only for true two-column layouts
                        side_for_width = _side_idx if layout_mode == "two_col" else -1

                        if hname in ("PROMO", "REG"):
                            _draw_price_fitting(
                                c, side_for_width, x, y, text, st, align,
                                max_w_mm=max_w_mm, margin_mm=margin_mm_val
                            )
                        elif hname in ("ITEM", "ENGLISH_DESCRIPTION"):
                            _draw_text_2line_shrink_left(
                                c, x, y, text, st,
                                max_w_mm=max_w_mm, margin_mm=margin_mm_val
                            )
                        else:
                            _draw_text_fitting(
                                c, side_for_width, x, y, text, st, align,
                                max_w_mm=max_w_mm, margin_mm=margin_mm_val
                            )

                else:
                    from_sec = str(a); to_sec = str(b)
                    x_mm, y_mm, _ = _slot_anchor_xy(header_map)
                    line1 = f"— {from_sec} complete —"
                    line2 = f"{to_sec} starts"
                    st1 = {"family": "Helvetica", "size": 12, "bold": True, "italic": False,
                           "underline": False, "strike": False, "color": "#000000"}
                    st2 = {"family": "Helvetica", "size": 11, "bold": False, "italic": True,
                           "underline": False, "strike": False, "color": "#000000"}
                    _draw_text(c, x_mm, y_mm - 2.5, line1, st1, "center")
                    _draw_text(c, x_mm, y_mm + 2.5, line2, st2, "center")


                c.restoreState()

            c.showPage()
        c.save()
        open_file(out)

    def _open_paste_panel(self):
        if self._paste_panel:
            self._paste_panel.show()
            return
        p = QDialog(self)
        p.setWindowTitle("Paste Items")
        p.setModal(True)
        self._paste_panel = p
        frm = QFrame()
        frm_layout = QVBoxLayout(frm)
        p.setLayout(frm_layout)
        self._paste_count_label = QLabel("Items: 0 (showing 0)")
        frm_layout.addWidget(self._paste_count_label)
        self._paste_list = QListWidget()
        frm_layout.addWidget(self._paste_list)
        ctr = QHBoxLayout()
        frm_layout.addLayout(ctr)
        paste_btn = QPushButton("📋 Paste")
        paste_btn.setFixedWidth(100)
        paste_btn.clicked.connect(self._panel_paste_from_clipboard)
        ctr.addWidget(paste_btn)
        add_btn = QPushButton("＋ Add")
        add_btn.setFixedWidth(80)
        add_btn.clicked.connect(self._panel_add_line)
        ctr.addWidget(add_btn)
        clear_btn = QPushButton("🗑 Clear")
        clear_btn.setFixedWidth(80)
        clear_btn.clicked.connect(self._panel_clear)
        ctr.addWidget(clear_btn)
        del_line_btn = QPushButton("Delete line")
        del_line_btn.setFixedWidth(120)
        del_line_btn.clicked.connect(self._panel_delete_line)
        ctr.addWidget(del_line_btn)
        collapse_btn = QPushButton("C")
        collapse_btn.setFixedWidth(60)
        collapse_btn.clicked.connect(self._panel_collapse_and_start)
        ctr.addWidget(collapse_btn)
        self._panel_refresh_preview()
        p.resize(400, 300)
        p.move(self.geometry().center() - p.rect().center())
        p.exec()

    def _panel_refresh_preview(self):
        self._paste_list.clear()
        for tok in self._paste_items[:10]:
            self._paste_list.addItem(tok)
        total = len(self._paste_items)
        shown = min(10, total)
        self._paste_count_label.setText(f"Items: {total} (showing {shown})")

    def _panel_paste_from_clipboard(self):
        app = QApplication.instance()
        clip = app.clipboard().text()
        pairs = self._parse_multi_lines(clip)
        self._paste_items = [tok for tok, _ in pairs]
        self._panel_refresh_preview()

    def _panel_add_line(self):
        val, ok = QInputDialog.getText(self._paste_panel, "Add Line", "Enter a code or token:")
        if ok and val:
            self._paste_items.append(val.strip())
            self._panel_refresh_preview()

    def _panel_delete_line(self):
        sel = self._paste_list.selectedItems()
        if not sel:
            return
        idx = self._paste_list.row(sel[0])
        if idx < len(self._paste_items):
            del self._paste_items[idx]
        self._panel_refresh_preview()

    def _panel_clear(self):
        self._paste_items = []
        self._panel_refresh_preview()

    def _panel_collapse_and_start(self):
        self._multi_mode_active = False
        self._multi_found_queue = []
        self._multi_found_qty = []
        self._multi_unfound_tokens = []
        self._multi_index = 0
        for tok in self._paste_items:
            rec = self._match_token_to_record(tok)
            if rec:
                self._multi_found_queue.append(rec)
                self._multi_found_qty.append(1)
            else:
                self._multi_unfound_tokens.append(tok)
        self._paste_panel.close()
# === CHUNK 3: ensure manual-entry normalization uppercases ASCII in ITEM/BRAND/UOM ===

def _apply_ascii_upper_core_fields(rec: dict) -> dict:
    """
    Mutates and returns rec: uppercases only ASCII a–z in ITEM/BRAND/UOM.
    Requires _upper_english from Chunk 1.
    """
    if not isinstance(rec, dict):
        return rec
    try:
        rec["ITEM"]  = _upper_english(rec.get("ITEM", ""))
        rec["BRAND"] = _upper_english(rec.get("BRAND", ""))
        rec["UOM"]   = _upper_english(rec.get("UOM", ""))
    except Exception:
        # Never fail normalization because of casing
        pass
    return rec

# Wrap the app’s normalization functions *if they exist*, preserving names/signatures.
# This keeps behavior identical, with only the added post-step for the three fields.
def _wrap_norm_func(func):
    def _wrapped(*args, **kwargs):
        out = func(*args, **kwargs)
        # Common patterns: either returns a dict (single row) or a list of dicts (batch)
        if isinstance(out, dict):
            return _apply_ascii_upper_core_fields(out)
        if isinstance(out, (list, tuple)):
            for i, r in enumerate(out):
                if isinstance(r, dict):
                    out[i] = _apply_ascii_upper_core_fields(r)
            return out
        return out
    _wrapped.__name__ = func.__name__
    _wrapped.__doc__ = func.__doc__
    return _wrapped

# Try the most likely normalization entry points without assuming they all exist.
for _fname in (
    "normalize_row_for_save",
    "normalize_row",
    "_normalize_row",
    "pre_save_normalize",
    "_pre_save_normalize",
):
    try:
        _orig = globals().get(_fname)
        if callable(_orig):
            globals()[f"__orig__{_fname}"] = _orig
            globals()[_fname] = _wrap_norm_func(_orig)
    except Exception:
        # If anything goes wrong, leave the original untouched.
        pass

        self._multi_start_stepthrough()

# file: file21oct.py  (drop-in: silent splash + post-show warmup)

from PySide6 import QtWidgets, QtGui, QtCore

# ---------------- Config toggles ----------------
SHOW_SPLASH: bool = False                 # False = no splash at all
SHOW_SPLASH_STATUS: bool = False         # False = no "Warming up..." text
WARMUP_DURING_SPLASH: bool = False       # False = warmup after window shows

# ---------------- Splash (no status by default) ----------------
class Splash(QtWidgets.QSplashScreen):
    def __init__(self):
        super().__init__(QtGui.QPixmap(), QtCore.Qt.WindowStaysOnTopHint)
        self.setWindowTitle("Starting…")
        self.setFixedSize(420, 200)
        if SHOW_SPLASH_STATUS:
            self.showMessage(
                "Launching…",
                QtCore.Qt.AlignHCenter | QtCore.Qt.AlignBottom,
                QtGui.QColor(0, 0, 0),
            )
        self.show()

    def update_text(self, txt: str) -> None:
        if SHOW_SPLASH_STATUS:
            self.showMessage(
                txt,
                QtCore.Qt.AlignHCenter | QtCore.Qt.AlignBottom,
                QtGui.QColor(0, 0, 0),
            )
            QtWidgets.QApplication.processEvents()


if __name__ == "__main__":
    from PySide6.QtCore import qVersion, Qt
    from PySide6.QtGui import QGuiApplication, QIcon
    from PySide6.QtWidgets import QApplication
    import importlib, threading, sys

    if qVersion().startswith("5."):
        QGuiApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        QGuiApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)

    try:
        app.setWindowIcon(QIcon(resource_path("app.ico")))
    except Exception:
        pass

    apply_styles(app)

    window = App()
    window.show()  # show immediately

    # Silent warm-up AFTER the window is visible (no UI messages)
    def _warmup():
        try:
            importlib.import_module("pandas")
            try:
                importlib.import_module("openpyxl")
            except Exception:
                pass
        except Exception:
            pass

    threading.Thread(target=_warmup, daemon=True).start()

    sys.exit(app.exec())
