"""
mbot_manager.py — Multi-mBot window manager for Silkroad Online (vSRO 110)
Requires: Python 3.11+, PyQt6, pywin32, pywinauto  (Windows only)
"""

import sys
import os
import re
import base64
import json
import ctypes
import subprocess
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from typing import Optional

from PyQt6.QtCore import Qt, QTimer, pyqtSignal
from PyQt6.QtGui import QColor, QPalette, QPainter, QBrush
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton, QFrame,
    QVBoxLayout, QHBoxLayout, QGridLayout, QStackedWidget, QScrollArea,
    QLineEdit, QCheckBox, QComboBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QAbstractItemView, QPlainTextEdit, QMessageBox,
    QFileDialog, QProgressBar,
)

# ---------------------------------------------------------------------------
# Qt warning suppressor — must run before QApplication
# ---------------------------------------------------------------------------
class _QtWarningFilter:
    _suppress = [
        "OleInitialize", "SetProcessDpiAwarenessContext",
        "DPI_AWARENESS_CONTEXT", "qt.conf", "QWindowsContext",
        "qt.qpa.window", "setHighDpiScaleFactorRoundingPolicy",
    ]
    def write(self, msg):
        if sys.__stderr__ and not any(s in msg for s in self._suppress):
            sys.__stderr__.write(msg)
    def flush(self):
        if sys.__stderr__:
            sys.__stderr__.flush()

sys.stderr = _QtWarningFilter()

# ---------------------------------------------------------------------------
# Win32 lazy-init — no-op stubs when not on Windows
# ---------------------------------------------------------------------------
WIN32_AVAILABLE = False
win32gui = win32con = win32com = win32process = None
findwindows = always_wait_until = PWTimeoutError = auto = None


def init_win32_modules() -> bool:
    global WIN32_AVAILABLE, win32gui, win32con, win32com, win32process
    global findwindows, always_wait_until, PWTimeoutError, auto
    if WIN32_AVAILABLE:
        return True
    try:
        import sys as _sys
        import warnings as _w
        _sys.coinit_flags = 2
        _w.filterwarnings("ignore", message="Apply externally defined coinit_flags*")

        import win32gui as _g, win32con as _c, win32com.client as _com, win32process as _p
        from pywinauto import findwindows as _fw
        from pywinauto.timings import always_wait_until as _awu, TimeoutError as _te
        import uiautomation as _a

        win32gui, win32con, win32com, win32process = _g, _c, _com, _p
        findwindows, always_wait_until, PWTimeoutError, auto = _fw, _awu, _te, _a
        WIN32_AVAILABLE = True
        return True
    except ImportError:
        return False


# ---------------------------------------------------------------------------
# Theme tokens
# ---------------------------------------------------------------------------
DARK = {
    "bg_window":    "#1f1f22", "bg_panel":  "#26262a", "bg_deep":  "#161618",
    "bg_input":     "#1a1a1d", "bg_hover":  "#2e2e33", "bg_active":"#34343a",
    "bg_titlebar":  "#1a1a1d",
    "border":       "#36363c", "border_light":"#2a2a2f",
    "text":         "#d4d4d8", "text_dim":  "#8a8a92", "text_mute":"#5f5f67",
    "accent":       "#5f8edf", "accent_hover":"#7aa3e6","accent_dim":"#3d5a8d",
    "danger":       "#d35d5d", "success":   "#6dc28a", "warn":     "#d6b35a",
    "hp":           "#d35d5d", "mp":        "#5f8edf",
}
T = DARK  # shorthand alias


def make_stylesheet(t: dict) -> str:
    return f"""
        QMainWindow, QWidget {{
            background:{t['bg_window']}; color:{t['text']};
            font-family:"Segoe UI","Helvetica Neue",system-ui,sans-serif; font-size:12px;
        }}
        QFrame#TitleBar {{ background:{t['bg_titlebar']}; border-bottom:1px solid {t['border']}; }}
        QFrame#Sidebar  {{ background:{t['bg_deep']};     border-right:1px solid {t['border']};  }}
        QPushButton#NavItem {{
            background:transparent; color:{t['text_dim']}; border:none;
            border-left:2px solid transparent; padding:12px 4px; text-align:center;
        }}
        QPushButton#NavItem:hover  {{ background:{t['bg_hover']};  color:{t['text']};  }}
        QPushButton#NavItem[active="true"] {{
            background:{t['bg_panel']}; color:{t['text']}; border-left:2px solid {t['accent']};
        }}
        QFrame#Col {{ background:{t['bg_window']}; border-right:1px solid {t['border']}; }}
        QLabel#ColHeader {{
            background:{t['bg_deep']}; color:{t['text_dim']}; border-bottom:1px solid {t['border']};
            padding:8px 12px; font-size:11px; font-weight:600; letter-spacing:0.5px;
        }}
        QLabel#Pill {{
            background:{t['bg_input']}; border:1px solid {t['border']}; border-radius:8px;
            padding:1px 6px; color:{t['text_dim']}; font-size:10px;
        }}
        QLabel#PanelTitle  {{ font-size:14px; font-weight:600; color:{t['text']};     }}
        QLabel#PanelSub    {{ font-size:11px;                  color:{t['text_dim']}; }}
        QLabel#SectionLabel {{ color:{t['text_mute']}; font-size:10px; font-weight:600; letter-spacing:0.5px; }}
        QFrame#MbotRow {{
            background:{t['bg_input']}; border:1px solid {t['border_light']}; border-radius:3px;
        }}
        QFrame#MbotRow[selected="true"] {{ background:{t['bg_hover']};  border:1px solid {t['accent']};     }}
        QFrame#MbotRow[focused="true"]  {{ background:{t['bg_active']}; border:1px solid {t['accent']};     }}
        QFrame#MbotRow:hover            {{                               border:1px solid {t['accent_dim']}; }}
        QFrame#CharCard, QFrame#SignupCard {{
            background:{t['bg_input']}; border:1px solid {t['border_light']}; border-radius:3px;
        }}
        QFrame#SignupCard {{ background:{t['bg_deep']}; border:1px solid {t['border']}; }}
        QPushButton {{
            background:{t['bg_panel']}; border:1px solid {t['border']};
            border-radius:2px; color:{t['text']}; padding:6px 10px;
        }}
        QPushButton:hover    {{ background:{t['bg_hover']};  border:1px solid {t['accent_dim']}; }}
        QPushButton:pressed  {{ background:{t['bg_active']}; }}
        QPushButton:disabled {{ color:{t['text_mute']}; }}
        QPushButton[primary="true"] {{
            background:{t['accent']};       border:1px solid {t['accent']};       color:white;
        }}
        QPushButton[primary="true"]:hover {{
            background:{t['accent_hover']}; border:1px solid {t['accent_hover']};
        }}
        QPushButton[danger="true"]:hover {{ border:1px solid {t['danger']}; color:{t['danger']}; }}
        QPushButton#ChatTab {{
            background:{t['bg_panel']}; border:1px solid {t['border']}; color:{t['text_dim']}; padding:4px 10px;
        }}
        QPushButton#ChatTab[active="true"] {{
            background:{t['accent']}; border:1px solid {t['accent']}; color:white;
        }}
        QLineEdit, QComboBox, QPlainTextEdit {{
            background:{t['bg_input']}; border:1px solid {t['border']}; border-radius:2px;
            color:{t['text']}; padding:4px 8px; selection-background-color:{t['accent']};
        }}
        QLineEdit:focus, QComboBox:focus {{ border:1px solid {t['accent']}; }}
        QComboBox::drop-down {{ border:none; width:18px; }}
        QCheckBox {{ color:{t['text']}; spacing:5px; }}
        QCheckBox::indicator {{
            width:13px; height:13px; border:1px solid {t['border']};
            background:{t['bg_input']}; border-radius:2px;
        }}
        QCheckBox::indicator:checked {{ background:{t['accent']}; border:1px solid {t['accent']}; image:none; }}
        QTableWidget {{
            background:{t['bg_deep']}; border:1px solid {t['border']};
            gridline-color:{t['border_light']}; color:{t['text']};
        }}
        QHeaderView::section {{
            background:{t['bg_panel']}; color:{t['text_dim']}; border:none;
            border-right:1px solid {t['border']}; border-bottom:1px solid {t['border']};
            padding:6px 10px; font-weight:600; font-size:10px; letter-spacing:0.5px;
        }}
        QTableWidget::item          {{ padding:6px 10px; }}
        QTableWidget::item:hover    {{ background:{t['bg_hover']}; }}
        QTableWidget::item:selected {{ background:{t['bg_active']}; color:{t['text']}; }}
        QTableWidget::item:checked  {{ background:{t['bg_active']}; border-left:2px solid {t['accent']}; }}
        QTableWidget::indicator {{
            width:14px; height:14px; border:1px solid {t['border_light']};
            border-radius:2px; background:{t['bg_input']};
        }}
        QTableWidget::indicator:checked       {{ background:{t['accent']};    border-color:{t['accent']};     image:none; }}
        QTableWidget::indicator:unchecked:hover {{ border-color:{t['accent_dim']}; }}
        QScrollBar:vertical, QScrollBar:horizontal {{
            background:{t['bg_deep']}; width:12px; height:12px;
        }}
        QScrollBar::handle:vertical, QScrollBar::handle:horizontal {{
            background:{t['border']}; border-radius:4px; min-height:20px; min-width:20px;
        }}
        QScrollBar::handle:vertical:hover, QScrollBar::handle:horizontal:hover {{
            background:{t['text_mute']};
        }}
        QScrollBar::add-line, QScrollBar::sub-line {{ background:none; height:0; width:0; }}
        QProgressBar {{
            background:{t['bg_deep']}; border:1px solid {t['border']}; border-radius:2px;
            height:10px; text-align:center; color:{t['text']}; font-size:9px;
        }}
        QProgressBar#HP::chunk {{ background:{t['hp']}; }}
        QProgressBar#MP::chunk {{ background:{t['mp']}; }}
        QProgressBar#XP::chunk {{ background:{t['accent']}; }}
        QPlainTextEdit#Log, QPlainTextEdit#ChatStream, QPlainTextEdit#InvLog {{
            background:{t['bg_deep']}; color:{t['text']};
            font-family:"JetBrains Mono","Consolas",monospace; font-size:11px;
        }}
    """


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------
@dataclass
class MbotInfo:
    """UI snapshot of one live mBot window."""
    id: int
    window_name: str
    char: str
    is_dc: bool
    hp: float = 0.0
    mp: float = 0.0
    kph: str = "–"

    @property
    def status(self) -> str:
        return "offline" if self.is_dc else "training"


def _parse_char_name(title: str) -> tuple[str, bool]:
    is_dc = "- DC" in title
    m = re.search(r"\[(.+?)(?:\s+-\s+DC)?\]", title)
    return (m.group(1) if m else title), is_dc


# Global live state
_live_windows: list = []
_live_mbots:   list[MbotInfo] = []


def now_ts() -> str:
    return datetime.now().strftime("%H:%M:%S")


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
ACCOUNTS_FILE       = "accounts.json"
CHAT_BUTTON_TEXTS   = ["Allchat","PM","Party","Guild","Global","Academy","GM","Union","Unique"]
INVENTORY_OPTIONS   = ["Avatar","Fellow","Guildstorage","Inventory","Pet","Storage"]

# ---------------------------------------------------------------------------
# TCVN3 → Unicode lookup
# ---------------------------------------------------------------------------
TCVN3_TO_UNICODE: dict[str, str] = {
    'µ':'à','¸':'á','¶':'ả','·':'ã','¹':'ạ',
    '¨':'ă','»':'ằ','¾':'ắ','¼':'ẳ','½':'ẵ','Æ':'ặ',
    '©':'â','Ç':'ầ','Ê':'ấ','È':'ẩ','É':'ẫ','Ë':'ậ',
    '®':'đ',
    'Ì':'è','Ð':'é','Î':'ẻ','Ï':'ẽ','Ñ':'ẹ',
    'ª':'ê','Ò':'ề','Õ':'ế','Ó':'ể','Ô':'ễ','Ö':'ệ',
    '×':'ì','Ý':'í','Ø':'ỉ','Ü':'ĩ','Þ':'ị',
    'ß':'ò','ã':'ó','á':'ỏ','â':'õ','ä':'ọ',
    '«':'ô','å':'ồ','è':'ố','æ':'ổ','ç':'ỗ','é':'ộ',
    '¬':'ơ','ê':'ờ','í':'ớ','ë':'ở','ì':'ỡ','î':'ợ',
    'ï':'ù','ó':'ú','ñ':'ủ','ò':'ũ','ô':'ụ',
    '­':'ư','õ':'ừ','ø':'ứ','ö':'ử','÷':'ữ','ù':'ự',
    'ú':'ỳ','ý':'ý','û':'ỷ','ü':'ỹ','þ':'ỵ',
    '§':'Đ','£':'Ê','¤':'Ô','¥':'Ơ','¦':'Ư',
}

def tcvn3_to_unicode_text(text: str) -> str:
    return ''.join(TCVN3_TO_UNICODE.get(ch, ch) for ch in text)


# ---------------------------------------------------------------------------
# Account persistence
# ---------------------------------------------------------------------------
_accounts: list = []


def load_accounts() -> list:
    if os.path.exists(ACCOUNTS_FILE):
        with open(ACCOUNTS_FILE, "r") as f:
            return json.load(f)
    return []


def save_accounts() -> None:
    with open(ACCOUNTS_FILE, "w") as f:
        json.dump(_accounts, f, indent=4)


# ---------------------------------------------------------------------------
# Win32 helpers
# ---------------------------------------------------------------------------
def extract_progress_bar(num_string: str) -> float:
    try:
        cur, tot = num_string.split("/")
        c = int(cur.replace(",","").strip()); t = int(tot.replace(",","").strip())
        return c * 100 / t if t else 0
    except Exception:
        return 0


def click_confirmation(
    class_name: str = "#32770", title: str = "Confirmation",
    text: str = "&Yes", is_re: bool = False,
    timeout: float = 1, retry_interval: float = 0.1,
) -> bool:
    if not WIN32_AVAILABLE:
        return False
    try:
        @always_wait_until(timeout, retry_interval)
        def _wait():
            els = (findwindows.find_elements(class_name=class_name, title_re=title)
                   if is_re else
                   findwindows.find_elements(class_name=class_name, title=title))
            for el in els:
                for child in el.children():
                    if child.name == text:
                        win32gui.SendMessage(child.handle, win32con.BM_CLICK, 0, 0)
                        return True
            return False
        return _wait()
    except PWTimeoutError:
        return False


# ---------------------------------------------------------------------------
# MBotWindow — wraps one live mBot window; all Win32 calls live here
# ---------------------------------------------------------------------------
class MBotWindow:
    def __init__(self, element):
        self.mbot = element
        self.name: str = ""
        # Cached element references (populated lazily)
        self._delay_edit         = None
        self._save_settings_btn  = None
        self._log_off_btn        = None
        self._start_client_btn   = None
        self._kill_client_btn    = None
        self._show_hide_cli_btn  = None
        self._reset_btn          = None
        self._stats_section      = None
        self._hp_value           = None
        self._mp_value           = None
        self._cur_pos_btn        = None
        self._start_train_btn    = None
        self._stop_train_btn     = None
        self._inv_combo          = None
        self._inv_refresh_btn    = None
        self._inv_items          = None
        self._clear_btn          = None
        self._log_edit           = None
        self._drops_cb           = None
        self._who_atk_cb         = None
        self._spy_player_cb      = None
        self._spy_refresh_btn    = None
        self._spy_combo          = None
        self._spy_log            = None
        self._chat_buttons: dict = {}

    def __str__(self): return f"MBotWindow({self.mbot.name})"

    def is_valid(self) -> bool:
        return WIN32_AVAILABLE and win32gui.IsWindow(self.mbot.handle)

    def _children(self):
        return self.mbot.children()

    # ── Element lookup helpers ────────────────────────────────────────────
    def _find_by_name(self, name):
        if not self.is_valid(): return None
        return next((c for c in self._children() if c.name == name), None)

    def _find_after(self, name):
        """Return the element immediately preceding a given name."""
        if not self.is_valid(): return None
        children = self._children()
        for i, child in enumerate(children):
            nxt = children[i+1] if i+1 < len(children) else None
            if nxt and nxt.name == name:
                return child
        return None

    def _find_nth(self, name, offset):
        """Return the element at (index_of_name + offset)."""
        if not self.is_valid(): return None
        children = self._children()
        for i, child in enumerate(children):
            if child.name == name:
                idx = i + offset
                return children[idx] if idx < len(children) else None
        return None

    # ── Stats ─────────────────────────────────────────────────────────────
    def get_hp(self) -> float | None:
        if not self.is_valid(): return None
        self._hp_value = self._hp_value or self._find_nth("HP", 6)
        return extract_progress_bar(self._hp_value.name) if self._hp_value else None

    def get_mp(self) -> float | None:
        if not self.is_valid(): return None
        self._mp_value = self._mp_value or self._find_nth("MP", 6)
        return extract_progress_bar(self._mp_value.name) if self._mp_value else None

    def get_name(self) -> str:
        if not self.is_valid(): return self.mbot.name
        if not self.name:
            el = self._find_nth("Hide client after relogin", 1)
            if el:
                parts = el.name.split(":")
                self.name = parts[1].strip() if len(parts) > 1 and parts[0] == "Name" else parts[0].strip()
            else:
                self.name = self.mbot.name
        return self.name

    def get_kills_per_hour(self) -> str:
        self._stats_section = self._stats_section or self._find_nth("Stop training", 2)
        if not self._stats_section:
            return "–"
        text = self._stats_section.name
        section = next((s for s in text.split("\n\n") if s.startswith("Per hour")), "")
        for line in section.splitlines():
            if line.startswith("Kills:"):
                return line.split(":")[1].strip().split(".")[0].strip()
        return "–"

    # ── Text read helper ──────────────────────────────────────────────────
    def _get_edit_content(self, handle) -> str:
        if not WIN32_AVAILABLE: return ""
        length = win32gui.SendMessage(handle, win32con.WM_GETTEXTLENGTH, 0, 0)
        buf = ctypes.create_unicode_buffer(length + 1)
        win32gui.SendMessage(handle, win32con.WM_GETTEXT, length + 1, buf)
        return tcvn3_to_unicode_text("\n".join(buf.value.splitlines()[-100:]))

    def get_chat_content(self, button_name: str) -> str | None:
        if not self.is_valid(): return None
        if button_name not in self._chat_buttons:
            offset = CHAT_BUTTON_TEXTS.index(button_name) + 1
            self._chat_buttons[button_name] = self._find_nth("Use colored chat", offset)
        btn = self._chat_buttons.get(button_name)
        return self._get_edit_content(btn.handle) if btn else None

    # ── Settings ──────────────────────────────────────────────────────────
    def set_delay(self, _is_default: bool = True) -> None:
        if not self.is_valid(): return
        self._delay_edit = self._delay_edit or self._find_after("minutes before relogin")
        if not self._delay_edit: return
        h = self._delay_edit.handle
        win32gui.SendMessage(h, win32con.WM_SETTEXT, 0, "")
        win32gui.SendMessage(h, win32con.WM_SETTEXT, 0, "999")

    def save_settings(self) -> None:
        if not self.is_valid(): return
        self._save_settings_btn = self._save_settings_btn or self._find_by_name("Save settings")
        if self._save_settings_btn:
            win32gui.SendMessage(self._save_settings_btn.handle, win32con.BM_CLICK, 0, 0)

    # ── Window controls ───────────────────────────────────────────────────
    def log_off(self) -> None:
        if not self.is_valid(): return
        self._log_off_btn = self._log_off_btn or self._find_by_name("Log Off")
        if self._log_off_btn:
            win32gui.PostMessage(self._log_off_btn.handle, win32con.BM_CLICK, 0, 0)
            click_confirmation()

    def start_client(self) -> None:
        if not self.is_valid(): return
        self._start_client_btn = self._start_client_btn or self._find_by_name("Start Client")
        if self._start_client_btn:
            win32gui.PostMessage(self._start_client_btn.handle, win32con.BM_CLICK, 0, 0)

    def kill_client(self) -> None:
        if not self.is_valid(): return
        self._kill_client_btn = self._kill_client_btn or self._find_by_name("Kill Client")
        if self._kill_client_btn:
            win32gui.PostMessage(self._kill_client_btn.handle, win32con.BM_CLICK, 0, 0)
            click_confirmation()

    def kill_mbot(self) -> None:
        if not self.is_valid(): return
        win32gui.PostMessage(self.mbot.handle, win32con.WM_CLOSE, 0, 0)
        click_confirmation()

    def show_hide_mbot(self) -> None:
        if not self.is_valid(): return
        h = self.mbot.handle
        flag = 0 if win32gui.IsWindowVisible(h) else 5
        ctypes.windll.user32.ShowWindow(h, flag)

    def show_hide_client(self) -> None:
        if not self.is_valid(): return
        self._show_hide_cli_btn = self._show_hide_cli_btn or self._find_by_name("Show / Hide Client")
        if self._show_hide_cli_btn:
            win32gui.PostMessage(self._show_hide_cli_btn.handle, win32con.BM_CLICK, 0, 0)

    def reset_mbot(self) -> None:
        if not self.is_valid(): return
        self._reset_btn = self._reset_btn or self._find_by_name("Reset")
        if self._reset_btn:
            win32gui.PostMessage(self._reset_btn.handle, win32con.BM_CLICK, 0, 0)

    def get_current_position(self) -> None:
        if not self.is_valid(): return
        self._cur_pos_btn = self._cur_pos_btn or self._find_by_name("Get current position")
        if self._cur_pos_btn:
            win32gui.PostMessage(self._cur_pos_btn.handle, win32con.BM_CLICK, 0, 0)

    def start_training(self) -> None:
        if not self.is_valid(): return
        self._start_train_btn = self._start_train_btn or self._find_by_name("Start training")
        if self._start_train_btn:
            win32gui.PostMessage(self._start_train_btn.handle, win32con.BM_CLICK, 0, 0)

    def stop_training(self) -> None:
        if not self.is_valid(): return
        self._stop_train_btn = self._stop_train_btn or self._find_by_name("Stop training")
        if self._stop_train_btn:
            win32gui.PostMessage(self._stop_train_btn.handle, win32con.BM_CLICK, 0, 0)

    # ── Inventory ─────────────────────────────────────────────────────────
    def set_inventory_combo(self, index: int) -> None:
        if not self.is_valid(): return
        self._inv_combo = self._inv_combo or self._find_nth("Inventory", 1)
        if self._inv_combo:
            win32gui.SendMessage(self._inv_combo.handle, win32con.CB_SETCURSEL, index, 0)

    def refresh_inventory(self) -> None:
        if not self.is_valid(): return
        self._inv_refresh_btn = self._inv_refresh_btn or self._find_nth("Inventory", 2)
        if self._inv_refresh_btn:
            win32gui.PostMessage(self._inv_refresh_btn.handle, win32con.BM_CLICK, 0, 0)

    def get_inventory_items(self) -> list[str]:
        if not self.is_valid(): return []
        self._inv_items = self._inv_items or self._find_nth("Inventory", 3)
        if not self._inv_items: return []
        h = self._inv_items.handle
        count = win32gui.SendMessage(h, win32con.LB_GETCOUNT, 0, 0)
        if count <= 0: return []
        items = []
        for i in range(count):
            length = win32gui.SendMessage(h, win32con.LB_GETTEXTLEN, i, 0)
            if length <= 0: continue
            buf = ctypes.create_unicode_buffer(length + 1)
            win32gui.SendMessage(h, win32con.LB_GETTEXT, i, buf)
            items.append(tcvn3_to_unicode_text(buf.value))
        return items

    # ── Log ───────────────────────────────────────────────────────────────
    def get_log(self) -> str | None:
        if not self.is_valid(): return None
        self._log_edit = self._log_edit or self._find_nth("Weaponswitch", 1)
        return self._get_edit_content(self._log_edit.handle) if self._log_edit else None

    def clear_log(self) -> None:
        if not self.is_valid(): return
        self._clear_btn = self._clear_btn or self._find_by_name("Clear")
        if self._clear_btn:
            win32gui.PostMessage(self._clear_btn.handle, win32con.BM_CLICK, 0, 0)

    # ── Checkboxes ────────────────────────────────────────────────────────
    def _get_cb(self, handle) -> bool:
        return win32gui.SendMessage(handle, win32con.BM_GETCHECK, 0, 0) == win32con.BST_CHECKED

    def _set_cb(self, handle, desired: bool) -> None:
        if self._get_cb(handle) != desired:
            win32gui.PostMessage(handle, win32con.BM_CLICK, 0, 0)

    def get_drops_checkbox_state(self) -> bool:
        if not self.is_valid(): return False
        self._drops_cb = self._drops_cb or self._find_by_name("Drops")
        return self._get_cb(self._drops_cb.handle) if self._drops_cb else False

    def set_drops_checkbox_state(self, desired: bool) -> None:
        if not self.is_valid(): return
        self._drops_cb = self._drops_cb or self._find_by_name("Drops")
        if self._drops_cb: self._set_cb(self._drops_cb.handle, desired)

    def get_who_attacked_you_checkbox_state(self) -> bool:
        if not self.is_valid(): return False
        self._who_atk_cb = self._who_atk_cb or self._find_by_name("Players who attacked you")
        return self._get_cb(self._who_atk_cb.handle) if self._who_atk_cb else False

    def set_who_attacked_you_checkbox_state(self, desired: bool) -> None:
        if not self.is_valid(): return
        self._who_atk_cb = self._who_atk_cb or self._find_by_name("Players who attacked you")
        if self._who_atk_cb: self._set_cb(self._who_atk_cb.handle, desired)

    # ── Spy / Active buffs ────────────────────────────────────────────────
    def set_spy_player_checkbox_state(self) -> None:
        if not self.is_valid(): return
        self._spy_player_cb = self._spy_player_cb or self._find_nth("Spy", 6)
        if self._spy_player_cb:
            if win32gui.SendMessage(self._spy_player_cb.handle, win32con.BM_GETCHECK, 0, 0) != win32con.BST_CHECKED:
                win32gui.PostMessage(self._spy_player_cb.handle, win32con.BM_CLICK, 0, 0)

    def refresh_spy(self) -> None:
        if not self.is_valid(): return
        self._spy_refresh_btn = self._spy_refresh_btn or self._find_nth("Spy", 5)
        if self._spy_refresh_btn:
            win32gui.PostMessage(self._spy_refresh_btn.handle, win32con.BM_CLICK, 0, 0)

    def get_active_buffs(self) -> list[str] | None:
        if not self.is_valid(): return None
        self._spy_combo = self._spy_combo or self._find_nth("Spy", 10)
        self._spy_log   = self._spy_log   or self._find_nth("Spy", 11)
        if not self._spy_combo or not self._spy_log: return None
        pattern = re.compile(rf"^Name:\s+{re.escape(self.get_name())}$")
        count = win32gui.SendMessage(self._spy_combo.handle, win32con.CB_GETCOUNT, 0, 0)
        for _ in range(count):
            win32gui.SendMessage(self._spy_combo.handle, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
            content = self._get_edit_content(self._spy_log.handle)
            result = []; found = collecting = False
            for line in content.splitlines():
                if pattern.search(line): found = True; continue
                if found:
                    if line.startswith("Active buffs:"): collecting = True; continue
                    if collecting: result.append(line.lstrip("\t"))
            if found: return result
        return None


# ---------------------------------------------------------------------------
# ProcessMbotsMixin — sequential QTimer-based action runner
# ---------------------------------------------------------------------------
class ProcessMbotsMixin:
    def process_mbots(self, mbot_list: list, actions: list[tuple]) -> None:
        def _run(index: int = 0) -> None:
            if index >= len(mbot_list): return
            mbot = mbot_list[index]
            acc_ms = 0
            for delay_ms, method in actions:
                acc_ms += delay_ms
                QTimer.singleShot(acc_ms, lambda m=mbot, fn=method: fn(m))
            total = sum(d for d, _ in actions)
            QTimer.singleShot(total + 100, lambda: _run(index + 1))
        _run(0)


# ---------------------------------------------------------------------------
# Window scan
# ---------------------------------------------------------------------------
def scan_mbot_windows() -> list:
    if not WIN32_AVAILABLE:
        return []
    try:
        raw = findwindows.find_elements(class_name="#32770", visible_only=False, title_re=r".*[Mm][Bb]ot.*")
        if not raw:
            all_w = findwindows.find_elements(class_name="#32770", visible_only=False)
            raw = [el for el in all_w if "mbot" in el.name.lower()]
        return [MBotWindow(el) for el in sorted(raw, key=lambda e: e.name)]
    except Exception:
        return []


# ---------------------------------------------------------------------------
# Reusable widgets
# ---------------------------------------------------------------------------

class StatusDot(QWidget):
    _COLORS = {
        "training": "#6dc28a", "idle": "#d6b35a",
        "dead":     "#d35d5d", "offline": "#5f5f67",
    }
    def __init__(self, status="offline", size=8, parent=None):
        super().__init__(parent)
        self.status = status
        self.setFixedSize(size, size)

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setBrush(QBrush(QColor(self._COLORS.get(self.status, "#5f5f67"))))
        p.setPen(Qt.PenStyle.NoPen)
        p.drawEllipse(0, 0, self.width(), self.height())


class MbotRow(QFrame):
    clicked = pyqtSignal(int)
    toggled = pyqtSignal(int, bool)

    def __init__(self, mbot: MbotInfo, multi=True):
        super().__init__()
        self.mbot  = mbot
        self.multi = multi
        self.setObjectName("MbotRow")
        self.setProperty("selected", False)
        self.setProperty("focused", False)
        self.setFixedHeight(28)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        lay = QHBoxLayout(self); lay.setContentsMargins(10, 4, 10, 4); lay.setSpacing(8)
        display = f"{mbot.char} - DC" if mbot.is_dc else mbot.char
        color   = T['text_mute'] if mbot.is_dc else T['text']
        name    = QLabel(display); name.setStyleSheet(f"font-size:12px; color:{color};")
        lay.addWidget(name, 1)

    def _repaint(self):
        self.style().unpolish(self); self.style().polish(self)

    def set_selected(self, v): self.setProperty("selected", v); self._repaint()
    def set_focused(self, v):  self.setProperty("focused",  v); self._repaint()

    def mousePressEvent(self, e):
        self.clicked.emit(self.mbot.id)
        if self.multi:
            if e.modifiers() & Qt.KeyboardModifier.ControlModifier:
                self.toggled.emit(self.mbot.id, not self.property("selected"))
            else:
                self.toggled.emit(self.mbot.id, True)
        super().mousePressEvent(e)


class MbotListColumn(QFrame):
    selection_changed = pyqtSignal(list)
    focus_changed     = pyqtSignal(int)

    def __init__(self, multi=True, initial_selected=None, initial_focus=None):
        super().__init__()
        self.setObjectName("Col")
        self.multi    = multi
        self.selected = list(initial_selected or [])
        self.focused  = initial_focus
        self.rows:    dict[int, MbotRow] = {}
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

        root = QVBoxLayout(self); root.setContentsMargins(0,0,0,0); root.setSpacing(0)

        # Header
        hdr = QFrame()
        hdr.setStyleSheet(f"background:{T['bg_deep']}; border-bottom:1px solid {T['border']};")
        hl = QHBoxLayout(hdr); hl.setContentsMargins(12,8,12,8); hl.setSpacing(8)
        title = QLabel("MBOTS"); title.setObjectName("ColHeader")
        title.setStyleSheet("background:transparent; border:none; padding:0;")
        self._count_pill = QLabel("0"); self._count_pill.setObjectName("Pill")
        hl.addWidget(title); hl.addStretch(1); hl.addWidget(self._count_pill)
        root.addWidget(hdr)

        # Quick-select toolbar (multi-select only)
        if multi:
            qa = QHBoxLayout(); qa.setContentsMargins(8,6,8,6); qa.setSpacing(6)
            sa = QPushButton("Select all"); sa.clicked.connect(self.select_all)
            cl = QPushButton("Clear");      cl.clicked.connect(self.clear)
            qa.addWidget(sa); qa.addWidget(cl)
            wrap = QFrame(); wrap.setLayout(qa)
            wrap.setStyleSheet(f"border-bottom:1px solid {T['border_light']};")
            root.addWidget(wrap)

        # Scrollable row area
        self._scroll = QScrollArea(); self._scroll.setWidgetResizable(True)
        self._scroll.setFrameShape(QFrame.Shape.NoFrame)
        self._body = QWidget()
        self._bl   = QVBoxLayout(self._body)
        self._bl.setContentsMargins(8,8,8,8); self._bl.setSpacing(4)
        self._scroll.setWidget(self._body)
        root.addWidget(self._scroll, 1)
        self.setFixedWidth(240)
        self.reload()

    def reload(self):
        existing = {m.id for m in _live_mbots}
        self.selected = [i for i in self.selected if i in existing]
        if self.focused not in existing:
            self.focused = _live_mbots[0].id if _live_mbots else None
        # Clear rows
        while self._bl.count():
            item = self._bl.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        self.rows.clear()
        for m in _live_mbots:
            row = MbotRow(m, self.multi)
            row.clicked.connect(self._on_focus)
            row.toggled.connect(self._on_toggle)
            row.set_selected(m.id in self.selected)
            row.set_focused(m.id == self.focused)
            self.rows[m.id] = row
            self._bl.addWidget(row)
        self._bl.addStretch(1)
        self._count_pill.setText(str(len(_live_mbots)))

    def _on_focus(self, mid):
        self.focused = mid
        for rid, row in self.rows.items():
            row.set_focused(rid == mid)
        self.focus_changed.emit(mid)
        self.setFocus()

    def _on_toggle(self, mid, on):
        mods = QApplication.keyboardModifiers()
        if self.multi and not (mods & Qt.KeyboardModifier.ControlModifier):
            self.selected = [mid] if on else []
            for rid, row in self.rows.items():
                row.set_selected(rid == mid and on)
        else:
            if on and mid not in self.selected:
                self.selected.append(mid)
            elif not on and mid in self.selected:
                self.selected.remove(mid)
            self.rows[mid].set_selected(on)
        self.selection_changed.emit(self.selected)

    def keyPressEvent(self, e):
        ids = [m.id for m in _live_mbots]
        if not ids: return super().keyPressEvent(e)
        cur = ids.index(self.focused) if self.focused in ids else 0
        if e.key() == Qt.Key.Key_Down:
            self._on_focus(ids[min(len(ids)-1, cur+1)]); return
        if e.key() == Qt.Key.Key_Up:
            self._on_focus(ids[max(0, cur-1)]); return
        if e.key() == Qt.Key.Key_Space and self.multi:
            self._on_toggle(self.focused, self.focused not in self.selected); return
        super().keyPressEvent(e)

    def select_all(self):
        self.selected = [m.id for m in _live_mbots]
        for row in self.rows.values(): row.set_selected(True)
        self.selection_changed.emit(self.selected)

    def clear(self):
        self.selected = []
        for row in self.rows.values(): row.set_selected(False)
        self.selection_changed.emit(self.selected)


# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------

class CharCard(QFrame):
    def __init__(self, mbot: MbotInfo):
        super().__init__()
        self.mbot_id = mbot.id
        self.setObjectName("CharCard")
        self.setFixedHeight(34)

        lay = QHBoxLayout(self); lay.setContentsMargins(10,5,10,5); lay.setSpacing(8)
        color = T['text_mute'] if mbot.is_dc else T['text']
        name  = QLabel(mbot.char); name.setStyleSheet(f"font-weight:600; font-size:12px; color:{color};")
        name.setFixedWidth(68); lay.addWidget(name)

        def _bar(label: str, color_key: str):
            lbl = QLabel(label); lbl.setStyleSheet(f"color:{T['text_mute']}; font-size:10px;"); lbl.setFixedWidth(16)
            bar = QProgressBar(); bar.setRange(0, 100); bar.setValue(0)
            bar.setTextVisible(False); bar.setFixedHeight(7)
            bar.setStyleSheet(
                f"QProgressBar{{background:{T['bg_deep']};border:1px solid {T['border']};border-radius:2px;}}"
                f"QProgressBar::chunk{{background:{T[color_key]};border-radius:2px;}}"
            )
            return lbl, bar

        hp_lbl, self.hp_bar = _bar("HP", "hp")
        mp_lbl, self.mp_bar = _bar("MP", "mp")
        lay.addWidget(hp_lbl); lay.addWidget(self.hp_bar, 2)
        lay.addWidget(mp_lbl); lay.addWidget(self.mp_bar, 2)

        self.kph_lbl = QLabel(); self.kph_lbl.setFixedWidth(80); lay.addWidget(self.kph_lbl)
        self.refresh(mbot)

    def refresh(self, mbot: MbotInfo):
        v = 0 if mbot.is_dc else int(mbot.hp)
        self.hp_bar.setValue(v)
        v = 0 if mbot.is_dc else int(mbot.mp)
        self.mp_bar.setValue(v)
        self.kph_lbl.setText(
            f"<span style='color:{T['text_mute']};font-size:10px;'>K/h</span> "
            f"<span style='color:{T['accent']};font-family:\"JetBrains Mono\",monospace;"
            f"font-size:11px;font-weight:600;'>{mbot.kph}</span>"
        )


class DashboardPanel(ProcessMbotsMixin, QWidget):
    log_event = pyqtSignal(str, str)

    _BUTTON_GRID = [
        # (id, label, kind, row, col)
        ("refresh",     "Refresh mBots",    None,  0, 0),
        ("showHide",    "Show/Hide mBots",  None,  0, 1),
        ("killBot",     "Kill mBots",       None,  0, 2),
        ("startClient", "Start client",     None,  1, 0),
        ("showHideCli", "Show/Hide Client", None,  1, 1),
        ("killClient",  "Kill client",      None,  1, 2),
        ("logoff",      "Log Off",          None,  1, 3),
        ("reset",       "Reset",            None,  2, 0),
        ("getPos",      "Get Position",     None,  2, 1),
        ("startTrain",  "Start Training",   None,  2, 2),
        ("stopTrain",   "Stop Training",    None,  2, 3),
    ]

    def __init__(self):
        super().__init__()
        self._dc_pending_next = None
        root = QVBoxLayout(self); root.setContentsMargins(0,0,0,0); root.setSpacing(0)

        # ── Top: mbot list + character cards ─────────────────────────────
        top = QHBoxLayout(); top.setContentsMargins(0,0,0,0); top.setSpacing(0)

        self.list_col = MbotListColumn(multi=True, initial_selected=[1], initial_focus=1)
        self.list_col.selection_changed.connect(self._on_sel)
        top.addWidget(self.list_col)

        col_chars = QFrame(); col_chars.setObjectName("Col")
        cc = QVBoxLayout(col_chars); cc.setContentsMargins(0,0,0,0); cc.setSpacing(0)
        hdr_chars = QFrame()
        hdr_chars.setStyleSheet(f"background:{T['bg_deep']}; border-bottom:1px solid {T['border']};")
        hc = QHBoxLayout(hdr_chars); hc.setContentsMargins(12,8,12,8)
        hc.addWidget(QLabel("CHARACTERS")); hc.addStretch(1)
        self.char_pill = QLabel("0 online"); self.char_pill.setObjectName("Pill")
        hc.addWidget(self.char_pill)
        cc.addWidget(hdr_chars)
        sc = QScrollArea(); sc.setWidgetResizable(True); sc.setFrameShape(QFrame.Shape.NoFrame)
        self._char_inner  = QWidget()
        self._char_layout = QVBoxLayout(self._char_inner)
        self._char_layout.setContentsMargins(10,10,10,10); self._char_layout.setSpacing(6)
        self._char_layout.addStretch(1)
        sc.setWidget(self._char_inner)
        cc.addWidget(sc, 1)
        top.addWidget(col_chars, 1)
        self._char_cards: dict[int, CharCard] = {}

        top_w = QWidget(); top_w.setLayout(top)
        root.addWidget(top_w, 1)

        # ── Bottom: action buttons ────────────────────────────────────────
        actions = QFrame()
        actions.setStyleSheet(f"background:{T['bg_deep']}; border-top:1px solid {T['border']};")
        af = QVBoxLayout(actions); af.setContentsMargins(0,0,0,0); af.setSpacing(0)

        act_hdr = QFrame()
        act_hdr.setStyleSheet(f"background:{T['bg_panel']}; border-bottom:1px solid {T['border']};")
        ah = QHBoxLayout(act_hdr); ah.setContentsMargins(12,6,12,6)
        ah.addWidget(QLabel("ACTIONS"))
        self.sel_pill = QLabel(f"{len(self.list_col.selected)} selected"); self.sel_pill.setObjectName("Pill")
        ah.addWidget(self.sel_pill); ah.addStretch(1)
        af.addWidget(act_hdr)

        grid_w = QWidget()
        grid = QGridLayout(grid_w); grid.setContentsMargins(10,8,10,8); grid.setSpacing(6)
        for c in range(4): grid.setColumnStretch(c, 1)
        for bid, label, kind, row, col in self._BUTTON_GRID:
            btn = QPushButton(label)
            if kind: btn.setProperty(kind, True)
            btn.style().unpolish(btn); btn.style().polish(btn)
            btn.clicked.connect(lambda _, b=bid, l=label: self._fire(b, l))
            grid.addWidget(btn, row, col)
        af.addWidget(grid_w)
        root.addWidget(actions)

        QTimer.singleShot(0, self._refresh_list)

    def _on_sel(self, sel):
        self.sel_pill.setText(f"{len(sel)} selected")

    def _selected_windows(self) -> list[MBotWindow]:
        sel_ids = set(self.list_col.selected)
        return [w for w, m in zip(_live_windows, _live_mbots) if m.id in sel_ids]

    def _do_scan(self):
        global _live_windows, _live_mbots
        if not WIN32_AVAILABLE:
            self.log_event.emit("win32/pywinauto not available — cannot scan mBot windows", "err")
            return
        try:
            all_32770 = findwindows.find_elements(class_name="#32770", visible_only=False)
            self.log_event.emit(f"Found {len(all_32770)} total #32770 windows", "info")
            for el in all_32770[:5]:
                self.log_event.emit(f"  Window: '{el.name}'", "info")
        except Exception as ex:
            self.log_event.emit(f"Error scanning windows: {ex}", "err"); return
        _live_windows = scan_mbot_windows()
        _live_mbots   = []
        for i, w in enumerate(_live_windows):
            char, is_dc = _parse_char_name(w.mbot.name)
            _live_mbots.append(MbotInfo(id=i+1, window_name=w.mbot.name, char=char, is_dc=is_dc))
        self._known_names = sorted(w.mbot.name for w in _live_windows if w.mbot.name)
        self._rebuild_ui()
        if callable(getattr(self, "_on_scan_done", None)):
            self._on_scan_done()
        self.log_event.emit(f"Refresh done — {len(_live_mbots)} mBots matched", "ok")

    def _rebuild_ui(self):
        self.list_col.reload()
        while self._char_layout.count():
            item = self._char_layout.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        self._char_cards.clear()
        for m in _live_mbots:
            card = CharCard(m)
            self._char_cards[m.id] = card
            self._char_layout.addWidget(card)
        self._char_layout.addStretch(1)
        online = sum(1 for m in _live_mbots if not m.is_dc)
        self.char_pill.setText(f"{online} online")
        if self._dc_pending_next:
            cb = self._dc_pending_next; self._dc_pending_next = None
            QTimer.singleShot(1000, cb)

    def _refresh_list(self):
        self._do_scan()

    def _fire(self, bid: str, label: str):
        sel   = self._selected_windows()
        names = ", ".join(m.char for m in _live_mbots if m.id in self.list_col.selected) or "(none)"
        kind  = "warn" if (bid.startswith("kill") or bid == "stopTrain") else (
                "ok"   if bid.startswith("start") else "info")

        if bid == "refresh":
            self._refresh_list(); return
        if not sel:
            self.log_event.emit(f"{label} — no mBots selected", "warn"); return

        if bid == "killBot":
            if QMessageBox.question(None, "Confirm", f"Close mBot(s): {names}?") != QMessageBox.StandardButton.Yes: return
            self.process_mbots(sel, [(0, lambda m: m.kill_mbot())])
            QTimer.singleShot(2000, self._refresh_list)
        elif bid == "killClient":
            if QMessageBox.question(None, "Confirm", f"Kill client(s): {names}?") != QMessageBox.StandardButton.Yes: return
            self.process_mbots(sel, [(0, lambda m: m.kill_client())])
        elif bid == "showHide":
            self.process_mbots(sel, [(0, lambda m: m.show_hide_mbot())])
        elif bid == "showHideCli":
            self.process_mbots(sel, [(0, lambda m: m.show_hide_client())])
        elif bid == "startClient":
            self.process_mbots(sel, [(0, lambda m: m.start_client())])
        elif bid == "logoff":
            self.process_mbots(sel, [(0, lambda m: m.log_off())])
        elif bid == "reset":
            self.process_mbots(sel, [(0, lambda m: m.reset_mbot())])
        elif bid == "getPos":
            self.process_mbots(sel, [(0, lambda m: m.get_current_position())])
        elif bid == "startTrain":
            self.process_mbots(sel, [(0, lambda m: m.start_training())])
        elif bid == "stopTrain":
            self.process_mbots(sel, [(0, lambda m: m.stop_training())])

        self.log_event.emit(f"{label} → {names}", kind)


# ---------------------------------------------------------------------------
# Account
# ---------------------------------------------------------------------------

class AccountPanel(ProcessMbotsMixin, QWidget):
    log_event = pyqtSignal(str, str)

    def __init__(self):
        super().__init__()
        self._item_changed_connected = False
        self.pending_login: tuple = ()

        root = QVBoxLayout(self); root.setContentsMargins(16,14,16,14); root.setSpacing(12)

        # Header
        head = QHBoxLayout(); head.setSpacing(12)
        text_col = QVBoxLayout(); text_col.setSpacing(2)
        text_col.addWidget(QLabel("Accounts", objectName="PanelTitle"))
        self.sub_label = QLabel(); self.sub_label.setObjectName("PanelSub")
        text_col.addWidget(self.sub_label)
        head.addLayout(text_col, 1)
        login_btn = QPushButton("  Login selected  ")
        login_btn.setProperty("primary", True); login_btn.style().unpolish(login_btn); login_btn.style().polish(login_btn)
        login_btn.setFixedHeight(32); login_btn.clicked.connect(self._login_selected)
        head.addWidget(login_btn, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        root.addLayout(head)

        # Toolbar
        tb = QHBoxLayout(); tb.setSpacing(6)
        sa = QPushButton("Select all");      sa.clicked.connect(self._select_all)
        ca = QPushButton("Clear all");       ca.clicked.connect(self._clear_all)
        rm = QPushButton("Remove selected"); rm.setProperty("danger", True)
        rm.style().unpolish(rm); rm.style().polish(rm)
        rm.clicked.connect(self._remove_selected)
        self.sel_pill = QLabel("0 selected"); self.sel_pill.setObjectName("Pill")
        tb.addWidget(sa); tb.addWidget(ca); tb.addWidget(rm); tb.addStretch(1); tb.addWidget(self.sel_pill)
        root.addLayout(tb)

        # Table
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["", "#", "Username", "Character", "mBot file path"])
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setShowGrid(False)
        h = self.table.horizontalHeader()
        for i, mode in enumerate([
            QHeaderView.ResizeMode.ResizeToContents,
            QHeaderView.ResizeMode.ResizeToContents,
            QHeaderView.ResizeMode.ResizeToContents,
            QHeaderView.ResizeMode.ResizeToContents,
            QHeaderView.ResizeMode.Stretch,
        ]): h.setSectionResizeMode(i, mode)
        self.table.setMinimumHeight(220)
        root.addWidget(self.table)
        self._refresh_table()

        # Add-account card — wrapped in scroll so buttons stay visible at small heights
        # Add-account card — 2-column compact layout, no scroll needed
        card = QFrame(); card.setObjectName("SignupCard")
        cl = QVBoxLayout(card); cl.setContentsMargins(12,10,12,10); cl.setSpacing(6)

        title_row = QHBoxLayout(); title_row.setSpacing(8)
        title_row.addWidget(QLabel("Add account", styleSheet="font-size:12px; font-weight:600;"))
        title_row.addStretch(1)
        cl.addLayout(title_row)

        self.in_user = QLineEdit(placeholderText="Username")
        self.in_pass = QLineEdit(placeholderText="Password")
        self.in_pass.setEchoMode(QLineEdit.EchoMode.Password)
        self.in_char = QLineEdit(placeholderText="Character (exact, case-sensitive)")
        self.in_path = QLineEdit(placeholderText=r"C:\MBot\mbot.exe")
        browse_btn = QPushButton("Browse…"); browse_btn.clicked.connect(self._browse_mbot)
        browse_btn.setFixedWidth(70)

        # Row 1: Username | Password
        r1 = QHBoxLayout(); r1.setSpacing(8)
        r1.addWidget(self.in_user, 1); r1.addWidget(self.in_pass, 1)
        cl.addLayout(r1)

        # Row 2: Character | mBot path + Browse
        r2 = QHBoxLayout(); r2.setSpacing(8)
        r2.addWidget(self.in_char, 1); r2.addWidget(self.in_path, 1); r2.addWidget(browse_btn)
        cl.addLayout(r2)

        # Row 3: buttons
        r3 = QHBoxLayout(); r3.setSpacing(6)
        add_btn = QPushButton("Add account"); add_btn.setProperty("primary", True)
        add_btn.style().unpolish(add_btn); add_btn.style().polish(add_btn)
        add_btn.clicked.connect(self._add)
        clr_btn = QPushButton("Clear"); clr_btn.clicked.connect(self._clear_form)
        r3.addWidget(add_btn); r3.addWidget(clr_btn); r3.addStretch(1)
        cl.addLayout(r3)

        root.addWidget(card)

    def _refresh_table(self):
        self.table.blockSignals(True)
        self.table.setRowCount(len(_accounts))
        for i, a in enumerate(_accounts):
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
            chk.setCheckState(Qt.CheckState.Unchecked)
            chk.setData(Qt.ItemDataRole.UserRole, i)
            self.table.setItem(i, 0, chk)
            self.table.setItem(i, 1, QTableWidgetItem(str(i+1)))
            self.table.setItem(i, 2, QTableWidgetItem(a.get("username", "")))
            self.table.setItem(i, 3, QTableWidgetItem(a.get("character", "")))
            path_it = QTableWidgetItem(a.get("mbot_file_path", ""))
            path_it.setForeground(QColor(T['text_dim']))
            path_it.setToolTip(a.get("mbot_file_path", ""))
            self.table.setItem(i, 4, path_it)

        self.table.blockSignals(False)
        self.table.resizeRowsToContents()
        if not self._item_changed_connected:
            self.table.itemChanged.connect(lambda it: it.column() == 0 and self._update_pill())
            self._item_changed_connected = True
        self._update_pill()
        self.sub_label.setText(f"{len(_accounts)} saved accounts. Each is bound to a .mbot profile file.")

    def _update_pill(self):
        n = sum(1 for r in range(self.table.rowCount())
                if (it := self.table.item(r, 0)) and it.checkState() == Qt.CheckState.Checked)
        self.sel_pill.setText(f"{n} selected")

    def _select_all(self):
        self.table.blockSignals(True)
        for r in range(self.table.rowCount()):
            it = self.table.item(r, 0)
            if it: it.setCheckState(Qt.CheckState.Checked)
        self.table.blockSignals(False); self._update_pill()

    def _clear_all(self):
        self.table.blockSignals(True)
        for r in range(self.table.rowCount()):
            it = self.table.item(r, 0)
            if it: it.setCheckState(Qt.CheckState.Unchecked)
        self.table.blockSignals(False); self._update_pill()

    def _remove_selected(self):
        indices = self._selected_indices()
        if not indices:
            return
        names = ", ".join(_accounts[i]["username"] for i in indices if i < len(_accounts))
        if QMessageBox.question(self, "Confirm", f"Delete {len(indices)} account(s): {names}?") != QMessageBox.StandardButton.Yes:
            return
        for i in sorted(indices, reverse=True):
            if i < len(_accounts):
                _accounts.pop(i)
        save_accounts()
        self._refresh_table()
        self.log_event.emit(f"Removed {len(indices)} account(s)", "warn")

    def _selected_indices(self) -> list[int]:
        return [
            self.table.item(r, 0).data(Qt.ItemDataRole.UserRole)
            for r in range(self.table.rowCount())
            if (it := self.table.item(r, 0)) and it.checkState() == Qt.CheckState.Checked
        ]

    # ── Login sequence ────────────────────────────────────────────────────
    def _login_selected(self):
        indices = self._selected_indices()
        if not indices:
            QMessageBox.information(self, "No selection", "Please select at least one account to log in.")
            return
        names = ", ".join(_accounts[i]["username"] for i in indices if i < len(_accounts))
        self.log_event.emit(f"Starting login sequence → {names}", "ok")
        if WIN32_AVAILABLE:
            self.pending_login = tuple(indices)
            self._start_mbot_client(0)
        else:
            self.log_event.emit("Win32 not available — login sequence skipped (not on Windows)", "warn")

    def _ensure_firewall(self, exe_path: str, username: str) -> None:
        rule_name = f"{username}_{os.path.basename(exe_path)}"
        try:
            check = subprocess.run(
                ["netsh", "advfirewall", "firewall", "show", "rule", f"name={rule_name}"],
                capture_output=True, text=True)
            if check.returncode == 0 and "No rules match" not in check.stdout:
                self.log_event.emit(f"Firewall rule already exists for {rule_name}", "info"); return
            subprocess.run([
                "netsh", "advfirewall", "firewall", "add", "rule",
                f"name={rule_name}", "dir=in", "action=allow",
                f"program={exe_path}", "profile=public", "enable=yes",
            ], capture_output=True)
            self.log_event.emit(f"Firewall rule added for {rule_name} (Public inbound)", "ok")
        except Exception as e:
            self.log_event.emit(f"Firewall setup failed: {e}", "warn")

    def _start_mbot_client(self, index: int) -> None:
        if index >= len(self.pending_login): return
        idx = self.pending_login[index]
        if idx >= len(_accounts): return
        acc      = _accounts[idx]
        username = acc["username"]
        for title in [f"[{username}] mBot v1.12b (vSRO 110)",
                      f"[{username} - DC] mBot v1.12b (vSRO 110)"]:
            if findwindows.find_elements(class_name="#32770", title=title, visible_only=False):
                self.log_event.emit(f"mBot already open for {username}, skipping", "info")
                QTimer.singleShot(1000, lambda: self._start_mbot_client(index + 1)); return
        mbot_path = acc.get("mbot_file_path", "")
        if not mbot_path or not os.path.exists(mbot_path):
            self.log_event.emit(f"mBot path not found for {username}: {mbot_path}", "err")
            QTimer.singleShot(500, lambda: self._start_mbot_client(index + 1)); return
        folder   = os.path.normpath(os.path.dirname(mbot_path))
        vsro_exe = os.path.join(folder, "mBot_vSRO110.exe")
        if os.path.exists(vsro_exe):
            self._ensure_firewall(vsro_exe, username)
        subprocess.Popen(mbot_path, cwd=folder)
        self.log_event.emit(f"Launched mBot for {username}", "info")
        QTimer.singleShot(3000, lambda: self._start_client_sro(index))

    def _start_client_sro(self, index: int) -> None:
        username  = _accounts[self.pending_login[index]]["username"]
        mbot_list = findwindows.find_elements(class_name="#32770", title="mBot v1.12b (vSRO 110)")
        if not mbot_list:
            QTimer.singleShot(1000, lambda: self._start_client_sro(index)); return
        MBotWindow(mbot_list[0]).start_client()
        self.log_event.emit(f"Start Client sent for {username}", "info")
        QTimer.singleShot(10_000, lambda: self._login_sro(index))

    def _login_sro(self, index: int) -> None:
        windows = findwindows.find_elements(class_name="CLIENT", title="SRO_Client")
        if not windows:
            QTimer.singleShot(2000, lambda: self._login_sro(index)); return
        handle = windows[0].handle
        ctypes.windll.user32.ShowWindow(handle, 5)
        win32gui.SetWindowPos(handle, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                              win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        left, top, right, bottom = win32gui.GetWindowRect(handle)
        cx = left + (right - left) // 2
        cy = top  + (bottom - top) // 2
        QTimer.singleShot(2000, lambda: self._login_click_center(index, cx, cy))

    def _login_click_center(self, index, cx, cy):
        auto.Click(cx, cy)
        QTimer.singleShot(2000, lambda: self._login_click_server(index, cx, cy))

    def _login_click_server(self, index, cx, cy):
        auto.Click(cx, cy - 125)
        QTimer.singleShot(2000, lambda: self._login_choose_server(index, cx, cy))

    def _login_choose_server(self, index, cx, cy):
        auto.Click(cx - 50, cy + 200)
        QTimer.singleShot(1500, lambda: self._login_enter_credentials(index))

    def _login_enter_credentials(self, index: int) -> None:
        acc      = _accounts[self.pending_login[index]]
        username = acc["username"]
        password = base64.b64decode(acc["password"]).decode("utf-8")
        for key in ('{Tab}', username, '{Tab}', password, '{Enter}'):
            auto.SendKeys(key, interval=0.08)
        self.log_event.emit(f"Credentials sent for {username}", "ok")
        QTimer.singleShot(15_000, lambda: self._start_training_sro(index))

    def _start_training_sro(self, index: int) -> None:
        acc       = _accounts[self.pending_login[index]]
        character = acc.get("character", acc["username"])
        windows   = findwindows.find_elements(class_name="CLIENT", title=character)
        if not windows:
            QTimer.singleShot(5000, lambda: self._start_training_sro(index)); return
        win32gui.SetWindowPos(windows[0].handle, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                              win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        mbot_list = findwindows.find_elements(class_name="#32770",
                                              title=f"[{character}] mBot v1.12b (vSRO 110)")
        if mbot_list:
            MBotWindow(mbot_list[0]).start_training()
        QTimer.singleShot(5000, lambda: self._hide_and_next(index))

    def _hide_and_next(self, index: int) -> None:
        acc       = _accounts[self.pending_login[index]]
        character = acc.get("character", acc["username"])
        mbot_list = findwindows.find_elements(class_name="#32770",
                                              title=f"[{character}] mBot v1.12b (vSRO 110)")
        if mbot_list:
            self.process_mbots([MBotWindow(mbot_list[0])], [
                (0,   lambda m: m.start_training()),
                (100, lambda m: m.show_hide_client()),
                (100, lambda m: m.show_hide_mbot()),
            ])
        self.log_event.emit(f"Login complete for {character}", "ok")
        QTimer.singleShot(2000, lambda: self._start_mbot_client(index + 1))

    # ── CRUD ──────────────────────────────────────────────────────────────
    def _browse_mbot(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select mBot executable", "",
            "Applications (*.exe);;All files (*)",
            options=QFileDialog.Option.DontUseNativeDialog)
        if path: self.in_path.setText(os.path.normpath(path))

    def _add(self):
        u    = self.in_user.text().strip()
        p    = self.in_pass.text().strip()
        char = self.in_char.text().strip()
        path = self.in_path.text().strip()
        if not u or not p:
            QMessageBox.warning(self, "Missing fields", "Username and password are required."); return
        if not char:
            QMessageBox.warning(self, "Missing fields", "Character name is required."); return
        if any(a["username"] == u for a in _accounts):
            QMessageBox.critical(self, "Error", "Username already exists!"); return
        _accounts.append({
            "username": u,
            "password": base64.b64encode(p.encode()).decode(),
            "character": char,
            "mbot_file_path": path,
        })
        save_accounts()
        self._refresh_table()
        self.log_event.emit(f"Added account '{u}' ({char})", "ok")
        self._clear_form()

    def _clear_form(self):
        for w in (self.in_user, self.in_pass, self.in_char, self.in_path):
            w.clear()

    def _remove(self, idx):
        if idx >= len(_accounts): return
        name = _accounts[idx]["username"]
        if QMessageBox.question(self, "Confirm", f"Delete account '{name}'?") != QMessageBox.StandardButton.Yes: return
        _accounts.pop(idx)
        save_accounts()
        self._refresh_table()
        self.log_event.emit(f"Removed account '{name}'", "warn")


# ---------------------------------------------------------------------------
# Chat
# ---------------------------------------------------------------------------

class ChatPanel(QWidget):
    def __init__(self):
        super().__init__()
        self.active_ch   = "Allchat"
        self.tab_buttons: dict[str, QPushButton] = {}
        self._chat_timer: QTimer | None = None

        root = QHBoxLayout(self); root.setContentsMargins(0,0,0,0); root.setSpacing(0)

        self.list_col = MbotListColumn(multi=False, initial_focus=None)
        self.list_col.focus_changed.connect(self._on_focus_changed)
        root.addWidget(self.list_col)

        right = QFrame(); right.setObjectName("Col")
        rl = QVBoxLayout(right); rl.setContentsMargins(0,0,0,0); rl.setSpacing(0)

        # Header
        hdr = QFrame()
        hdr.setStyleSheet(f"background:{T['bg_deep']}; border-bottom:1px solid {T['border']};")
        hh = QHBoxLayout(hdr); hh.setContentsMargins(12,8,12,8)
        hh.addWidget(QLabel("CHAT")); hh.addStretch(1)
        self.chat_pill = QLabel("—"); self.chat_pill.setObjectName("Pill")
        hh.addWidget(self.chat_pill)
        rl.addWidget(hdr)

        # Channel tabs
        tabs = QFrame()
        tabs.setStyleSheet(f"background:{T['bg_deep']}; border-bottom:1px solid {T['border']};")
        tw = QHBoxLayout(tabs); tw.setContentsMargins(10,8,10,8); tw.setSpacing(4)
        for ch in CHAT_BUTTON_TEXTS:
            b = QPushButton(ch); b.setObjectName("ChatTab")
            b.setProperty("active", ch == self.active_ch)
            b.style().unpolish(b); b.style().polish(b)
            b.clicked.connect(lambda _, c=ch: self._set_channel(c))
            self.tab_buttons[ch] = b; tw.addWidget(b)
        tw.addStretch(1)
        rl.addWidget(tabs)

        self.stream = QPlainTextEdit(); self.stream.setObjectName("ChatStream")
        self.stream.setReadOnly(True)
        rl.addWidget(self.stream, 1)
        root.addWidget(right, 1)

    def _focused_window(self) -> MBotWindow | None:
        fid = self.list_col.focused
        return next((w for w, m in zip(_live_windows, _live_mbots) if m.id == fid), None)

    def _on_focus_changed(self, mid):
        self.chat_pill.setText(next((m.char for m in _live_mbots if m.id == mid), "—"))
        self._start_chat_poll()

    def _set_channel(self, ch):
        self.active_ch = ch
        for c, b in self.tab_buttons.items():
            b.setProperty("active", c == ch)
            b.style().unpolish(b); b.style().polish(b)
        self._start_chat_poll()

    def _start_chat_poll(self):
        if self._chat_timer: self._chat_timer.stop()
        self._do_poll()

    def _do_poll(self):
        w = self._focused_window()
        if w:
            content = w.get_chat_content(self.active_ch)
            if content is not None and self.stream.toPlainText() != content:
                self.stream.setPlainText(content)
                self.stream.verticalScrollBar().setValue(self.stream.verticalScrollBar().maximum())
        self._chat_timer = QTimer(self)
        self._chat_timer.setSingleShot(True)
        self._chat_timer.timeout.connect(self._do_poll)
        self._chat_timer.start(20_000)


# ---------------------------------------------------------------------------
# Inventory
# ---------------------------------------------------------------------------

class InventoryPanel(QWidget):
    def __init__(self):
        super().__init__()
        root = QHBoxLayout(self); root.setContentsMargins(0,0,0,0); root.setSpacing(0)

        self.list_col = MbotListColumn(multi=False, initial_focus=1)
        self.list_col.focus_changed.connect(self._refresh)
        root.addWidget(self.list_col)

        right = QFrame(); right.setObjectName("Col")
        rl = QVBoxLayout(right); rl.setContentsMargins(0,0,0,0); rl.setSpacing(0)

        hdr = QFrame()
        hdr.setStyleSheet(f"background:{T['bg_deep']}; border-bottom:1px solid {T['border']};")
        hh = QHBoxLayout(hdr); hh.setContentsMargins(12,8,12,8)
        hh.addWidget(QLabel("INVENTORY & LOG")); hh.addStretch(1)
        self.head_pill = QLabel("—"); self.head_pill.setObjectName("Pill")
        hh.addWidget(self.head_pill)
        rl.addWidget(hdr)

        # 3-column body
        body_layout = QHBoxLayout(); body_layout.setContentsMargins(0,0,0,0); body_layout.setSpacing(0)

        # Inventory column
        inv_wrap = QFrame()
        inv_wrap.setStyleSheet(f"background:{T['bg_window']}; border-right:1px solid {T['border']};")
        iv = QVBoxLayout(inv_wrap); iv.setContentsMargins(0,0,0,0); iv.setSpacing(0)
        inv_hdr = QFrame()
        inv_hdr.setStyleSheet(f"background:{T['bg_panel']}; border-bottom:1px solid {T['border']};")
        ih = QHBoxLayout(inv_hdr); ih.setContentsMargins(10,5,10,5); ih.setSpacing(8)
        inv_title = QLabel("INVENTORY")
        inv_title.setStyleSheet(f"color:{T['text_dim']};font-weight:600;font-size:10px;letter-spacing:0.5px;")
        self.inv_combo = QComboBox(); self.inv_combo.addItems(INVENTORY_OPTIONS)
        self.inv_combo.setCurrentText("Inventory"); self.inv_combo.setFixedWidth(110)
        self.inv_combo.currentTextChanged.connect(self._refresh)
        ih.addWidget(inv_title); ih.addStretch(1); ih.addWidget(self.inv_combo)
        iv.addWidget(inv_hdr)
        self.inv_log = QPlainTextEdit(); self.inv_log.setObjectName("InvLog"); self.inv_log.setReadOnly(True)
        iv.addWidget(self.inv_log, 1)
        body_layout.addWidget(inv_wrap, 1)

        # Active buffs column
        self.buff_col = self._plain_column("ACTIVE BUFFS")
        body_layout.addWidget(self.buff_col["wrap"], 1)

        # Event log column
        ev_wrap = QFrame()
        ev_wrap.setStyleSheet(f"background:{T['bg_window']}; border-right:1px solid {T['border']};")
        ev = QVBoxLayout(ev_wrap); ev.setContentsMargins(0,0,0,0); ev.setSpacing(0)
        ev_hdr = QFrame()
        ev_hdr.setStyleSheet(f"background:{T['bg_panel']}; border-bottom:1px solid {T['border']};")
        eh = QHBoxLayout(ev_hdr); eh.setContentsMargins(10,5,10,5); eh.setSpacing(10)
        ev_title = QLabel("EVENT LOG")
        ev_title.setStyleSheet(f"color:{T['text_dim']};font-weight:600;font-size:10px;letter-spacing:0.5px;")
        ev_clr = QPushButton("Clear"); ev_clr.setFixedHeight(20)
        ev_clr.setStyleSheet("padding:1px 8px; font-size:10px;")
        ev_clr.clicked.connect(lambda: self.ev_log.clear())
        eh.addWidget(ev_title); eh.addStretch(1); eh.addWidget(ev_clr)
        ev.addWidget(ev_hdr)
        self.ev_log = QPlainTextEdit(); self.ev_log.setObjectName("InvLog"); self.ev_log.setReadOnly(True)
        ev.addWidget(self.ev_log, 1)
        body_layout.addWidget(ev_wrap, 1)

        body_w = QWidget(); body_w.setLayout(body_layout)
        rl.addWidget(body_w, 1)
        root.addWidget(right, 1)
        self._refresh()

    @staticmethod
    def _plain_column(title: str) -> dict:
        wrap = QFrame()
        wrap.setStyleSheet(f"background:{T['bg_window']}; border-right:1px solid {T['border']};")
        v = QVBoxLayout(wrap); v.setContentsMargins(0,0,0,0); v.setSpacing(0)
        hdr = QLabel(title)
        hdr.setStyleSheet(
            f"background:{T['bg_panel']};color:{T['text_dim']};"
            f"border-bottom:1px solid {T['border']};"
            f"padding:6px 12px;font-weight:600;font-size:10px;letter-spacing:0.5px;")
        v.addWidget(hdr)
        log = QPlainTextEdit(); log.setObjectName("InvLog"); log.setReadOnly(True)
        v.addWidget(log, 1)
        return {"wrap": wrap, "log": log}

    def _refresh(self, *_):
        focus_id = self.list_col.focused
        focus_m  = next((m for m in _live_mbots if m.id == focus_id), None)
        focus_w  = next((w for w, m in zip(_live_windows, _live_mbots) if m.id == focus_id), None)
        self.head_pill.setText(focus_m.char if focus_m else "—")

        inv_type = self.inv_combo.currentText()
        inv_idx  = INVENTORY_OPTIONS.index(inv_type) if inv_type in INVENTORY_OPTIONS else 3
        inv_html = [f"<div style='margin-bottom:4px'><span style='color:{T['accent']};font-size:10px;"
                    f"font-weight:600;'>[{inv_type}]</span></div>"]

        if focus_w:
            focus_w.set_inventory_combo(inv_idx)
            focus_w.refresh_inventory()
            raw_items = focus_w.get_inventory_items()
            totals: dict[str, int] = defaultdict(int)
            slots:  dict[str, int] = defaultdict(int)
            for line in raw_items:
                m = re.search(r':\s*(.*?)\s*\((\d+)\s+pieces\)', line)
                if m:
                    totals[m.group(1)] += int(m.group(2))
                    slots[m.group(1)]  += 1
            if totals:
                for item in sorted(totals):
                    inv_html.append(
                        f"<div><span style='color:{T['text']}'>{item}</span>"
                        f"<span style='color:{T['text_mute']}'> — </span>"
                        f"<span style='color:{T['accent']};font-family:monospace'>{totals[item]}</span>"
                        f"<span style='color:{T['text_mute']}'>pcs / {slots[item]} slots</span></div>"
                    )
            else:
                inv_html.append(f"<div style='color:{T['text_mute']}'>No stackable items found.</div>")
        else:
            inv_html.append(f"<div style='color:{T['text_mute']}'>No mBot selected or not running.</div>")

        self.inv_log.clear(); self.inv_log.appendHtml("".join(inv_html))

        # Buffs
        buff_html = []
        if focus_w:
            focus_w.set_spy_player_checkbox_state()
            focus_w.refresh_spy()
            buffs = focus_w.get_active_buffs()
            buff_html = (
                [f"<div style='color:{T['accent']}'>{b}</div>" for b in buffs]
                if buffs else
                [f"<div style='color:{T['text_mute']}'>No buffs found.</div>"]
            )
        else:
            buff_html = [f"<div style='color:{T['text_mute']}'>No mBot selected.</div>"]
        self.buff_col["log"].clear(); self.buff_col["log"].appendHtml("".join(buff_html))

        # Event log
        if focus_w:
            raw = focus_w.get_log() or ""
            ev_html = [f"<div style='color:{T['text_dim']}'>{line}</div>"
                       for line in raw.splitlines() if line.strip()]
            self.ev_log.clear()
            self.ev_log.appendHtml("".join(ev_html) if ev_html
                                   else f"<div style='color:{T['text_mute']}'>No log entries.</div>")
        else:
            self.ev_log.clear()
            self.ev_log.appendHtml(f"<div style='color:{T['text_mute']}'>No mBot selected.</div>")



# ---------------------------------------------------------------------------
# Log
# ---------------------------------------------------------------------------

class LogPanel(QWidget):
    def __init__(self):
        super().__init__()
        root = QVBoxLayout(self); root.setContentsMargins(0,0,0,0); root.setSpacing(0)

        hdr = QFrame()
        hdr.setStyleSheet(f"background:{T['bg_panel']}; border-bottom:1px solid {T['border']};")
        lh = QHBoxLayout(hdr); lh.setContentsMargins(12,6,12,6)
        lh.addWidget(QLabel("LOG", styleSheet=f"color:{T['text_dim']};font-weight:600;font-size:10px;letter-spacing:0.5px;"))
        self.log_count = QLabel("0 entries"); self.log_count.setObjectName("Pill")
        clr_btn = QPushButton("Clear"); clr_btn.clicked.connect(self.clear)
        lh.addWidget(self.log_count); lh.addStretch(1); lh.addWidget(clr_btn)
        root.addWidget(hdr)

        self.log = QPlainTextEdit(); self.log.setObjectName("Log"); self.log.setReadOnly(True)
        root.addWidget(self.log, 1)

    def append(self, msg: str, kind: str = "info", who: Optional[str] = None):
        colors = {
            "info": T['text_dim'], "ok": T['success'], "warn": T['warn'],
            "err":  T['danger'],   "accent": T['accent'],
        }
        color    = colors.get(kind, T['text_dim'])
        prefix   = f"<span style='color:{T['text_mute']}'>[{now_ts()}]</span> "
        who_html = f"<span style='color:{T['accent']}'>{who}:</span> " if who else ""
        self.log.appendHtml(f"{prefix}{who_html}<span style='color:{color}'>{msg}</span>")
        n = int(self.log_count.text().split()[0]) + 1
        self.log_count.setText(f"{n} entries")
        self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())

    def clear(self):
        self.log.clear(); self.log_count.setText("0 entries")

# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class MainWindow(ProcessMbotsMixin, QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MBot Manager")
        self.resize(900, 620)
        self.setStyleSheet(make_stylesheet(DARK))
        self._known_names: list[str] = []

        central = QWidget(); self.setCentralWidget(central)
        root = QVBoxLayout(central); root.setContentsMargins(0,0,0,0); root.setSpacing(0)

        # Title bar
        title_bar = QFrame(); title_bar.setObjectName("TitleBar"); title_bar.setFixedHeight(32)
        tb = QHBoxLayout(title_bar); tb.setContentsMargins(12,0,0,0); tb.setSpacing(8)
        self.title_text = QLabel()
        tb.addWidget(self.title_text); tb.addStretch(1)
        root.addWidget(title_bar)
        self._update_title(0)

        # Main row: sidebar + content
        main = QHBoxLayout(); main.setContentsMargins(0,0,0,0); main.setSpacing(0)

        sidebar = QFrame(); sidebar.setObjectName("Sidebar"); sidebar.setFixedWidth(96)
        sl = QVBoxLayout(sidebar); sl.setContentsMargins(0,4,0,0); sl.setSpacing(0)

        self.stack = QStackedWidget()
        self.dash  = DashboardPanel()
        self.acc   = AccountPanel()
        self.chat  = ChatPanel()
        self.inv   = InventoryPanel()
        self.log_panel = LogPanel()
        for w in (self.dash, self.acc, self.chat, self.inv, self.log_panel):
            self.stack.addWidget(w)
        self.dash.log_event.connect(self._append_log)
        self.acc.log_event.connect(self._append_log)

        def _on_scan_done():
            self.chat.list_col.reload()
            self.inv.list_col.reload()
        self.dash._on_scan_done = _on_scan_done

        self.nav_buttons: list[QPushButton] = []
        for i, label in enumerate(["Dashboard","Account","Chat","Inventory","Log"]):
            b = QPushButton(label); b.setObjectName("NavItem")
            b.setProperty("active", i == 0)
            b.style().unpolish(b); b.style().polish(b)
            b.clicked.connect(lambda _, idx=i: self._switch(idx))
            sl.addWidget(b); self.nav_buttons.append(b)
        sl.addStretch(1)
        main.addWidget(sidebar)

        right = QFrame()
        rl = QVBoxLayout(right); rl.setContentsMargins(0,0,0,0); rl.setSpacing(0)
        rl.addWidget(self.stack, 1)


        main.addWidget(right, 1)
        root.addLayout(main, 1)

        # Initial messages
        self._append_log("Welcome to MBot Manager v0.1.0", "accent")
        self._append_log(f"Loaded {len(_accounts)} accounts from {ACCOUNTS_FILE}", "info")
        if not WIN32_AVAILABLE:
            self._append_log("win32/pywinauto not available — running in UI-only mode", "warn")

        # HP/MP/KPH poll every 1 s
        self._hp_mp_timer = QTimer(self)
        self._hp_mp_timer.timeout.connect(self._poll_hp_mp)
        self._hp_mp_timer.start(1000)

        # Window change detection every 60 s
        self._scan_timer = QTimer(self)
        self._scan_timer.timeout.connect(self._scan_if_changed)
        self._scan_timer.start(60_000)

    def _update_title(self, online: int):
        self.title_text.setText(
            f"<b>MBot Manager</b> <span style='color:{T['text_mute']}'>v0.1.0</span>"
            f"  <span style='background:rgba(0,0,0,0.2);padding:2px 8px;border-radius:3px;'>"
            f"<span style='color:{T['success']}'>●</span> {online} mBots online</span>"
        )

    def _scan_if_changed(self):
        if not WIN32_AVAILABLE: return
        global _live_windows, _live_mbots
        try:
            new_windows = scan_mbot_windows()
        except Exception as e:
            self._append_log(f"[scan] error: {e}", "err"); return
        new_names = sorted(w.mbot.name for w in new_windows if w.mbot.name)
        if new_names == self._known_names: return
        self._known_names = new_names[:]     # store a copy to avoid aliasing
        _live_windows = new_windows
        _live_mbots   = []
        for i, w in enumerate(_live_windows):
            char, is_dc = _parse_char_name(w.mbot.name)
            _live_mbots.append(MbotInfo(id=i+1, window_name=w.mbot.name, char=char, is_dc=is_dc))
        self.dash._rebuild_ui()
        self.chat.list_col.reload()
        self.inv.list_col.reload()

    def _poll_hp_mp(self):
        if not _live_windows: return
        changed = False
        for w, m in zip(_live_windows, _live_mbots):
            try:
                hp  = w.get_hp();  kph = w.get_kills_per_hour(); mp = w.get_mp()
                if hp  is not None: m.hp  = hp;  changed = True
                if mp  is not None: m.mp  = mp;  changed = True
                if kph:             m.kph = kph; changed = True
            except Exception as e:
                self._append_log(f"Poll error [{m.char}]: {e}", "err")
        if changed:
            for mid, card in self.dash._char_cards.items():
                mbot = next((x for x in _live_mbots if x.id == mid), None)
                if mbot: card.refresh(mbot)
            self._update_title(sum(1 for m in _live_mbots if not m.is_dc))

    def _switch(self, idx):
        self.stack.setCurrentIndex(idx)
        for i, b in enumerate(self.nav_buttons):
            b.setProperty("active", i == idx)
            b.style().unpolish(b); b.style().polish(b)

    def _append_log(self, msg: str, kind: str = "info", who: Optional[str] = None):
        self.log_panel.append(msg, kind, who)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    global _accounts, _live_windows, _live_mbots

    import traceback as _tb
    def _excepthook(exc_type, exc_val, exc_tb):
        msg = "".join(_tb.format_exception(exc_type, exc_val, exc_tb))
        print(msg, file=sys.__stderr__)
        try:
            with open("crash.log", "a") as f: f.write(msg + "\n")
        except Exception:
            pass
    sys.excepthook = _excepthook

    autologin = "--autologin" in sys.argv
    os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"

    _accounts     = load_accounts()
    app           = QApplication(sys.argv)
    init_win32_modules()
    app.setStyle("Fusion")

    pal = app.palette()
    pal.setColor(QPalette.ColorRole.Window,     QColor(T['bg_window']))
    pal.setColor(QPalette.ColorRole.Base,       QColor(T['bg_input']))
    pal.setColor(QPalette.ColorRole.Text,       QColor(T['text']))
    pal.setColor(QPalette.ColorRole.WindowText, QColor(T['text']))
    pal.setColor(QPalette.ColorRole.Button,     QColor(T['bg_panel']))
    pal.setColor(QPalette.ColorRole.ButtonText, QColor(T['text']))
    app.setPalette(pal)

    _live_windows = scan_mbot_windows()
    _live_mbots   = []
    for i, w in enumerate(_live_windows):
        char, is_dc = _parse_char_name(w.mbot.name)
        _live_mbots.append(MbotInfo(id=i+1, window_name=w.mbot.name, char=char, is_dc=is_dc))

    w = MainWindow(); w.show()

    if autologin and _accounts:
        w._append_log(f"--autologin: selecting all {len(_accounts)} accounts and logging in", "ok")
        w.acc._select_all()
        QTimer.singleShot(500, w.acc._login_selected)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()