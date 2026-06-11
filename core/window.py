"""
core/window.py — Win32 mBot window interaction layer.

All pywin32/pywinauto/uiautomation calls are isolated here.
No Qt dependency. Safe to import in a subprocess.
"""

import re
import ctypes
import subprocess
from dataclasses import dataclass, field
from typing import Optional

# ---------------------------------------------------------------------------
# Win32 lazy-init
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
# TCVN3 → Unicode
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
# Constants
# ---------------------------------------------------------------------------
CHAT_BUTTON_TEXTS  = ["Allchat", "PM", "Party", "Guild", "Global", "Academy", "GM", "Union", "Unique"]
INVENTORY_OPTIONS  = ["Avatar", "Fellow", "Guildstorage", "Inventory", "Pet", "Storage"]


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------
@dataclass
class MbotInfo:
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


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def extract_progress_bar(num_string: str) -> float:
    try:
        cur, tot = num_string.split("/")
        c = int(cur.replace(",", "").strip())
        t = int(tot.replace(",", "").strip())
        return c * 100 / t if t else 0
    except Exception:
        return 0


def click_confirmation(
    class_name: str = "#32770",
    title: str = "Confirmation",
    text: str = "&Yes",
    is_re: bool = False,
    timeout: float = 1,
    retry_interval: float = 0.1,
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
# MBotWindow
# ---------------------------------------------------------------------------
class MBotWindow:
    def __init__(self, element):
        self.mbot = element
        self.name: str = ""
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

    def __str__(self):
        return f"MBotWindow({self.mbot.name})"

    def is_valid(self) -> bool:
        return WIN32_AVAILABLE and win32gui.IsWindow(self.mbot.handle)

    def _children(self):
        return self.mbot.children()

    def _find_by_name(self, name):
        if not self.is_valid():
            return None
        return next((c for c in self._children() if c.name == name), None)

    def _find_after(self, name):
        if not self.is_valid():
            return None
        children = self._children()
        for i, child in enumerate(children):
            nxt = children[i + 1] if i + 1 < len(children) else None
            if nxt and nxt.name == name:
                return child
        return None

    def _find_nth(self, name, offset):
        if not self.is_valid():
            return None
        children = self._children()
        for i, child in enumerate(children):
            if child.name == name:
                idx = i + offset
                return children[idx] if idx < len(children) else None
        return None

    # ── Stats ─────────────────────────────────────────────────────────────
    def get_hp(self) -> Optional[float]:
        if not self.is_valid():
            return None
        self._hp_value = self._hp_value or self._find_nth("HP", 6)
        return extract_progress_bar(self._hp_value.name) if self._hp_value else None

    def get_mp(self) -> Optional[float]:
        if not self.is_valid():
            return None
        self._mp_value = self._mp_value or self._find_nth("MP", 6)
        return extract_progress_bar(self._mp_value.name) if self._mp_value else None

    def get_name(self) -> str:
        if not self.is_valid():
            return self.mbot.name
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

    # ── Text helper ────────────────────────────────────────────────────────
    def _get_edit_content(self, handle) -> str:
        if not WIN32_AVAILABLE:
            return ""
        length = win32gui.SendMessage(handle, win32con.WM_GETTEXTLENGTH, 0, 0)
        buf = ctypes.create_unicode_buffer(length + 1)
        win32gui.SendMessage(handle, win32con.WM_GETTEXT, length + 1, buf)
        return tcvn3_to_unicode_text("\n".join(buf.value.splitlines()[-100:]))

    def get_chat_content(self, button_name: str) -> Optional[str]:
        if not self.is_valid():
            return None
        if button_name not in self._chat_buttons:
            offset = CHAT_BUTTON_TEXTS.index(button_name) + 1
            self._chat_buttons[button_name] = self._find_nth("Use colored chat", offset)
        btn = self._chat_buttons.get(button_name)
        return self._get_edit_content(btn.handle) if btn else None

    # ── Settings ──────────────────────────────────────────────────────────
    def set_delay(self, _is_default: bool = True) -> None:
        if not self.is_valid():
            return
        self._delay_edit = self._delay_edit or self._find_after("minutes before relogin")
        if not self._delay_edit:
            return
        h = self._delay_edit.handle
        win32gui.SendMessage(h, win32con.WM_SETTEXT, 0, "")
        win32gui.SendMessage(h, win32con.WM_SETTEXT, 0, "999")

    def save_settings(self) -> None:
        if not self.is_valid():
            return
        self._save_settings_btn = self._save_settings_btn or self._find_by_name("Save settings")
        if self._save_settings_btn:
            win32gui.SendMessage(self._save_settings_btn.handle, win32con.BM_CLICK, 0, 0)

    # ── Window controls ───────────────────────────────────────────────────
    def log_off(self) -> None:
        if not self.is_valid():
            return
        self._log_off_btn = self._log_off_btn or self._find_by_name("Log Off")
        if self._log_off_btn:
            win32gui.PostMessage(self._log_off_btn.handle, win32con.BM_CLICK, 0, 0)
            click_confirmation()

    def start_client(self) -> None:
        if not self.is_valid():
            return
        self._start_client_btn = self._start_client_btn or self._find_by_name("Start Client")
        if self._start_client_btn:
            win32gui.PostMessage(self._start_client_btn.handle, win32con.BM_CLICK, 0, 0)

    def kill_client(self) -> None:
        if not self.is_valid():
            return
        self._kill_client_btn = self._kill_client_btn or self._find_by_name("Kill Client")
        if self._kill_client_btn:
            win32gui.PostMessage(self._kill_client_btn.handle, win32con.BM_CLICK, 0, 0)
            click_confirmation()

    def kill_mbot(self) -> None:
        if not self.is_valid():
            return
        win32gui.PostMessage(self.mbot.handle, win32con.WM_CLOSE, 0, 0)
        click_confirmation()

    def show_hide_mbot(self) -> None:
        if not self.is_valid():
            return
        h = self.mbot.handle
        flag = 0 if win32gui.IsWindowVisible(h) else 5
        ctypes.windll.user32.ShowWindow(h, flag)

    def show_hide_client(self) -> None:
        if not self.is_valid():
            return
        self._show_hide_cli_btn = self._show_hide_cli_btn or self._find_by_name("Show / Hide Client")
        if self._show_hide_cli_btn:
            win32gui.PostMessage(self._show_hide_cli_btn.handle, win32con.BM_CLICK, 0, 0)

    def reset_mbot(self) -> None:
        if not self.is_valid():
            return
        self._reset_btn = self._reset_btn or self._find_by_name("Reset")
        if self._reset_btn:
            win32gui.PostMessage(self._reset_btn.handle, win32con.BM_CLICK, 0, 0)

    def get_current_position(self) -> None:
        if not self.is_valid():
            return
        self._cur_pos_btn = self._cur_pos_btn or self._find_by_name("Get current position")
        if self._cur_pos_btn:
            win32gui.PostMessage(self._cur_pos_btn.handle, win32con.BM_CLICK, 0, 0)

    def start_training(self) -> None:
        if not self.is_valid():
            return
        self._start_train_btn = self._start_train_btn or self._find_by_name("Start training")
        if self._start_train_btn:
            win32gui.PostMessage(self._start_train_btn.handle, win32con.BM_CLICK, 0, 0)

    def stop_training(self) -> None:
        if not self.is_valid():
            return
        self._stop_train_btn = self._stop_train_btn or self._find_by_name("Stop training")
        if self._stop_train_btn:
            win32gui.PostMessage(self._stop_train_btn.handle, win32con.BM_CLICK, 0, 0)

    # ── Inventory ─────────────────────────────────────────────────────────
    def set_inventory_combo(self, index: int) -> None:
        if not self.is_valid():
            return
        self._inv_combo = self._inv_combo or self._find_nth("Inventory", 1)
        if self._inv_combo:
            win32gui.SendMessage(self._inv_combo.handle, win32con.CB_SETCURSEL, index, 0)

    def refresh_inventory(self) -> None:
        if not self.is_valid():
            return
        self._inv_refresh_btn = self._inv_refresh_btn or self._find_nth("Inventory", 2)
        if self._inv_refresh_btn:
            win32gui.PostMessage(self._inv_refresh_btn.handle, win32con.BM_CLICK, 0, 0)

    def get_inventory_items(self) -> list[str]:
        if not self.is_valid():
            return []
        self._inv_items = self._inv_items or self._find_nth("Inventory", 3)
        if not self._inv_items:
            return []
        h = self._inv_items.handle
        count = win32gui.SendMessage(h, win32con.LB_GETCOUNT, 0, 0)
        if count <= 0:
            return []
        items = []
        for i in range(count):
            length = win32gui.SendMessage(h, win32con.LB_GETTEXTLEN, i, 0)
            if length <= 0:
                continue
            buf = ctypes.create_unicode_buffer(length + 1)
            win32gui.SendMessage(h, win32con.LB_GETTEXT, i, buf)
            items.append(tcvn3_to_unicode_text(buf.value))
        return items

    # ── Log ───────────────────────────────────────────────────────────────
    def get_log(self) -> Optional[str]:
        if not self.is_valid():
            return None
        self._log_edit = self._log_edit or self._find_nth("Weaponswitch", 1)
        return self._get_edit_content(self._log_edit.handle) if self._log_edit else None

    def clear_log(self) -> None:
        if not self.is_valid():
            return
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
        if not self.is_valid():
            return False
        self._drops_cb = self._drops_cb or self._find_by_name("Drops")
        return self._get_cb(self._drops_cb.handle) if self._drops_cb else False

    def set_drops_checkbox_state(self, desired: bool) -> None:
        if not self.is_valid():
            return
        self._drops_cb = self._drops_cb or self._find_by_name("Drops")
        if self._drops_cb:
            self._set_cb(self._drops_cb.handle, desired)

    def get_who_attacked_you_checkbox_state(self) -> bool:
        if not self.is_valid():
            return False
        self._who_atk_cb = self._who_atk_cb or self._find_by_name("Players who attacked you")
        return self._get_cb(self._who_atk_cb.handle) if self._who_atk_cb else False

    def set_who_attacked_you_checkbox_state(self, desired: bool) -> None:
        if not self.is_valid():
            return
        self._who_atk_cb = self._who_atk_cb or self._find_by_name("Players who attacked you")
        if self._who_atk_cb:
            self._set_cb(self._who_atk_cb.handle, desired)

    # ── Spy / Active buffs ────────────────────────────────────────────────
    def set_spy_player_checkbox_state(self) -> None:
        if not self.is_valid():
            return
        self._spy_player_cb = self._spy_player_cb or self._find_nth("Spy", 6)
        if self._spy_player_cb:
            if win32gui.SendMessage(self._spy_player_cb.handle, win32con.BM_GETCHECK, 0, 0) != win32con.BST_CHECKED:
                win32gui.PostMessage(self._spy_player_cb.handle, win32con.BM_CLICK, 0, 0)

    def refresh_spy(self) -> None:
        if not self.is_valid():
            return
        self._spy_refresh_btn = self._spy_refresh_btn or self._find_nth("Spy", 5)
        if self._spy_refresh_btn:
            win32gui.PostMessage(self._spy_refresh_btn.handle, win32con.BM_CLICK, 0, 0)

    def get_active_buffs(self) -> Optional[list[str]]:
        if not self.is_valid():
            return None
        self._spy_combo = self._spy_combo or self._find_nth("Spy", 10)
        self._spy_log   = self._spy_log   or self._find_nth("Spy", 11)
        if not self._spy_combo or not self._spy_log:
            return None
        pattern = re.compile(rf"^Name:\s+{re.escape(self.get_name())}$")
        count = win32gui.SendMessage(self._spy_combo.handle, win32con.CB_GETCOUNT, 0, 0)
        for _ in range(count):
            win32gui.SendMessage(self._spy_combo.handle, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
            content = self._get_edit_content(self._spy_log.handle)
            result = []
            found = collecting = False
            for line in content.splitlines():
                if pattern.search(line):
                    found = True
                    continue
                if found:
                    if line.startswith("Active buffs:"):
                        collecting = True
                        continue
                    if collecting:
                        result.append(line.lstrip("\t"))
            if found:
                return result
        return None


# ---------------------------------------------------------------------------
# Window scan
# ---------------------------------------------------------------------------
def scan_mbot_windows() -> list[MBotWindow]:
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
# Dialog dismissal helpers
# ---------------------------------------------------------------------------
def dismiss_bsobj_dialogs(log_fn=None) -> int:
    if not WIN32_AVAILABLE:
        return 0
    dismissed = 0
    try:
        for el in findwindows.find_elements(title="BSObj Plugin"):
            for child in el.children():
                if child.name in ("OK", "&OK"):
                    win32gui.PostMessage(child.handle, win32con.BM_CLICK, 0, 0)
                    dismissed += 1
                    if log_fn:
                        log_fn(f"[BSObj] Dismissed dialog (handle={el.handle})", "warn")
                    break
    except Exception as e:
        if log_fn:
            log_fn(f"[BSObj] Check error: {e}", "err")
    return dismissed


def dismiss_neterror_dialogs(log_fn=None) -> None:
    if not WIN32_AVAILABLE:
        return
    try:
        for el in findwindows.find_elements(title="NetError"):
            for child in el.children():
                if child.name in ("OK", "&OK"):
                    win32gui.PostMessage(child.handle, win32con.BM_CLICK, 0, 0)
                    if log_fn:
                        log_fn(f"[NetError] Dismissed dialog (handle={el.handle})", "warn")
                    break
    except Exception as e:
        if log_fn:
            log_fn(f"[NetError] Check error: {e}", "err")


def dismiss_openerror_dialogs(log_fn=None) -> None:
    if not WIN32_AVAILABLE:
        return
    try:
        for el in findwindows.find_elements(class_name="#32770", title="Error"):
            for child in el.children():
                if child.name in ("OK", "&OK"):
                    win32gui.PostMessage(child.handle, win32con.BM_CLICK, 0, 0)
                    if log_fn:
                        log_fn(f"[Error] Dismissed dialog (handle={el.handle})", "warn")
                    break
    except Exception as e:
        if log_fn:
            log_fn(f"[Error] Check error: {e}", "err")


# ---------------------------------------------------------------------------
# Silkroad process helpers (used by update sequence)
# ---------------------------------------------------------------------------
def kill_silkroad_processes() -> None:
    try:
        subprocess.run(["taskkill", "/F", "/IM", "silkroad.exe"], capture_output=True)
        subprocess.run(["taskkill", "/F", "/IM", "Silkroad.exe"], capture_output=True)
    except Exception:
        pass


def count_silkroad_controls() -> int:
    if not WIN32_AVAILABLE:
        return 0
    try:
        import psutil
        sro_pids = {
            p.info["pid"]
            for p in psutil.process_iter(["pid", "name"])
            if p.info["name"] and p.info["name"].lower() == "silkroad.exe"
        }
        if not sro_pids:
            return 0

        def _count_children(hwnd) -> int:
            count = 0
            child = win32gui.GetWindow(hwnd, win32con.GW_CHILD)
            while child:
                count += 1 + _count_children(child)
                child = win32gui.GetWindow(child, win32con.GW_HWNDNEXT)
            return count

        best = 0

        def _enum(hwnd, _):
            nonlocal best
            if not win32gui.IsWindowVisible(hwnd):
                return
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
            except Exception:
                return
            if pid not in sro_pids:
                return
            count = _count_children(hwnd)
            if count > best:
                best = count

        win32gui.EnumWindows(_enum, None)
        return best
    except Exception:
        return 0
