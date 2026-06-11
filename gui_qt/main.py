"""
gui_qt/main.py — PyQt6 GUI for MBot Manager.

Reads all state from SQLite (via core.db).
Sends all commands via db.enqueue_command() — no direct win32 calls.

Run:
    python -m gui_qt.main
"""

import base64
import json
import os
import re
import sys
from collections import defaultdict
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

from core import db
from core.window import CHAT_BUTTON_TEXTS, INVENTORY_OPTIONS

# ---------------------------------------------------------------------------
# Qt warning suppressor
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
        if sys.__stderr__: sys.__stderr__.flush()

sys.stderr = _QtWarningFilter()

# ---------------------------------------------------------------------------
# Theme
# ---------------------------------------------------------------------------
DARK = {
    'bg_window': '#1a1a24', 'bg_panel':  '#1e1e2e', 'bg_deep':  '#13131b',
    'bg_input':  '#13131b', 'text':      '#e0e0f0', 'text_dim': '#9090a8',
    'text_mute': '#5a5a72', 'border':    '#2a2a3e', 'border_light': '#34344e',
    'accent':    '#7c6af7', 'success':   '#6dc28a', 'warn':     '#d6b35a',
    'danger':    '#d35d5d', 'hp':        '#6dc28a', 'mp':       '#5a8fd6',
}
T = DARK

ACCOUNTS_FILE  = "accounts.json"
UPDATER_FILE   = "updater.json"

_updater_paths: list[str] = []


def now_ts() -> str:
    return datetime.now().strftime("%H:%M:%S")


def load_accounts() -> list:
    if os.path.exists(ACCOUNTS_FILE):
        with open(ACCOUNTS_FILE) as f:
            return json.load(f)
    return []


def save_accounts(accounts: list) -> None:
    with open(ACCOUNTS_FILE, "w") as f:
        json.dump(accounts, f, indent=4)


def load_updater_paths() -> list[str]:
    if os.path.exists(UPDATER_FILE):
        with open(UPDATER_FILE) as f:
            return json.load(f)
    return []


def save_updater_paths() -> None:
    with open(UPDATER_FILE, "w") as f:
        json.dump(_updater_paths, f, indent=4)


def make_stylesheet(t: dict) -> str:
    return f"""
        QMainWindow, QWidget {{ background:{t['bg_window']}; color:{t['text']}; font-family:"Segoe UI","Inter",sans-serif; font-size:12px; }}
        QFrame#Sidebar {{ background:{t['bg_panel']}; border-right:1px solid {t['border']}; }}
        QFrame#TitleBar {{ background:{t['bg_deep']}; border-bottom:1px solid {t['border']}; }}
        QFrame#Col {{ background:{t['bg_window']}; border-right:1px solid {t['border']}; }}
        QFrame#CharCard {{ background:{t['bg_panel']}; border:1px solid {t['border']}; border-radius:4px; }}
        QFrame#SignupCard {{ background:{t['bg_panel']}; border:1px solid {t['border']}; border-radius:6px; }}
        QLabel {{ background:transparent; color:{t['text']}; }}
        QLabel#PanelTitle {{ font-size:16px; font-weight:600; }}
        QLabel#PanelSub {{ color:{t['text_mute']}; font-size:11px; }}
        QLabel#ColHeader {{ color:{t['text_dim']}; font-weight:600; font-size:10px; letter-spacing:0.5px; }}
        QLabel#Pill {{ background:{t['bg_deep']}; color:{t['accent']}; padding:2px 8px; border-radius:3px; font-size:10px; }}
        QPushButton {{ background:{t['bg_panel']}; color:{t['text']}; border:1px solid {t['border']}; border-radius:4px; padding:4px 12px; }}
        QPushButton:hover {{ background:{t['border']}; }}
        QPushButton[primary=true] {{ background:{t['accent']}; color:#fff; border:none; }}
        QPushButton[primary=true]:hover {{ background:#6b5ae0; }}
        QPushButton[danger=true] {{ border-color:{t['danger']}; color:{t['danger']}; }}
        QPushButton[danger=true]:hover {{ background:rgba(211,93,93,0.15); }}
        QPushButton#NavItem {{ background:transparent; border:none; border-radius:0; padding:10px 0; color:{t['text_mute']}; font-size:11px; }}
        QPushButton#NavItem:hover {{ color:{t['text']}; background:rgba(255,255,255,0.04); }}
        QPushButton#NavItem[active=true] {{ color:{t['accent']}; border-left:2px solid {t['accent']}; }}
        QPushButton#ChatTab {{ background:transparent; border:none; border-radius:3px; padding:3px 8px; color:{t['text_mute']}; font-size:10px; }}
        QPushButton#ChatTab:hover {{ color:{t['text']}; background:rgba(255,255,255,0.06); }}
        QPushButton#ChatTab[active=true] {{ color:{t['accent']}; background:rgba(124,106,247,0.12); }}
        QLineEdit, QComboBox {{ background:{t['bg_input']}; color:{t['text']}; border:1px solid {t['border']}; border-radius:4px; padding:4px 8px; }}
        QLineEdit:focus, QComboBox:focus {{ border-color:{t['accent']}; }}
        QComboBox::drop-down {{ border:none; width:20px; }}
        QTableWidget {{ background:{t['bg_input']}; color:{t['text']}; border:1px solid {t['border']}; gridline-color:transparent; }}
        QTableWidget::item {{ padding:4px 8px; border-bottom:1px solid {t['border']}; }}
        QTableWidget::item:selected {{ background:rgba(124,106,247,0.2); }}
        QHeaderView::section {{ background:{t['bg_panel']}; color:{t['text_dim']}; border:none; border-bottom:1px solid {t['border']}; padding:4px 8px; font-size:10px; letter-spacing:0.3px; }}
        QScrollArea {{ border:none; }}
        QScrollBar:vertical {{ background:{t['bg_deep']}; width:4px; }}
        QScrollBar::handle:vertical {{ background:{t['border']}; border-radius:2px; }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height:0; }}
        QPlainTextEdit#Log, QPlainTextEdit#ChatStream, QPlainTextEdit#InvLog {{
            background:{t['bg_deep']}; color:{t['text']};
            font-family:"JetBrains Mono","Consolas",monospace; font-size:11px;
        }}
        QFrame#MbotRow {{ background:transparent; border:none; border-radius:3px; }}
        QFrame#MbotRow:hover {{ background:rgba(255,255,255,0.04); }}
        QFrame#MbotRow[selected=true] {{ background:rgba(124,106,247,0.15); }}
        QFrame#MbotRow[focused=true] {{ background:rgba(124,106,247,0.08); }}
    """


# ---------------------------------------------------------------------------
# Reusable widgets
# ---------------------------------------------------------------------------
class StatusDot(QWidget):
    _COLORS = {
        "training": "#6dc28a", "idle": "#d6b35a",
        "dead": "#d35d5d",     "offline": "#5f5f67",
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

    def __init__(self, mbot: dict, multi: bool = True):
        super().__init__()
        self.mbot_id = mbot["id"]
        self.multi   = multi
        self.setObjectName("MbotRow")
        self.setProperty("selected", False)
        self.setProperty("focused",  False)
        self.setFixedHeight(28)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        lay     = QHBoxLayout(self); lay.setContentsMargins(10, 4, 10, 4); lay.setSpacing(8)
        is_dc   = bool(mbot.get("is_dc"))
        display = f"{mbot['char']} - DC" if is_dc else mbot["char"]
        color   = T['text_mute'] if is_dc else T['text']
        name    = QLabel(display)
        name.setStyleSheet(f"font-size:12px; color:{color};")
        lay.addWidget(name, 1)

    def _repaint(self):
        self.style().unpolish(self); self.style().polish(self)

    def set_selected(self, v): self.setProperty("selected", v); self._repaint()
    def set_focused(self, v):  self.setProperty("focused",  v); self._repaint()

    def mousePressEvent(self, e):
        self.clicked.emit(self.mbot_id)
        if self.multi:
            if e.modifiers() & Qt.KeyboardModifier.ControlModifier:
                self.toggled.emit(self.mbot_id, not self.property("selected"))
            else:
                self.toggled.emit(self.mbot_id, True)
        super().mousePressEvent(e)


class MbotListColumn(QFrame):
    selection_changed = pyqtSignal(list)
    focus_changed     = pyqtSignal(int)

    def __init__(self, multi: bool = True, initial_focus: Optional[int] = None):
        super().__init__()
        self.setObjectName("Col")
        self.multi    = multi
        self.selected: list[int] = []
        self.focused  = initial_focus
        self.rows:    dict[int, MbotRow] = {}

        root = QVBoxLayout(self); root.setContentsMargins(0, 0, 0, 0); root.setSpacing(0)

        hdr = QFrame()
        hdr.setStyleSheet(f"background:{T['bg_deep']}; border-bottom:1px solid {T['border']};")
        hl  = QHBoxLayout(hdr); hl.setContentsMargins(12, 8, 12, 8); hl.setSpacing(8)
        title = QLabel("MBOTS"); title.setObjectName("ColHeader")
        title.setStyleSheet("background:transparent; border:none; padding:0;")
        self._count_pill = QLabel("0"); self._count_pill.setObjectName("Pill")
        hl.addWidget(title); hl.addStretch(1); hl.addWidget(self._count_pill)
        root.addWidget(hdr)

        if multi:
            qa = QHBoxLayout(); qa.setContentsMargins(8, 6, 8, 6); qa.setSpacing(6)
            sa = QPushButton("Select all"); sa.clicked.connect(self.select_all)
            cl = QPushButton("Clear");      cl.clicked.connect(self.clear)
            qa.addWidget(sa); qa.addWidget(cl)
            wrap = QFrame(); wrap.setLayout(qa)
            wrap.setStyleSheet(f"border-bottom:1px solid {T['border_light']};")
            root.addWidget(wrap)

        self._scroll = QScrollArea(); self._scroll.setWidgetResizable(True)
        self._scroll.setFrameShape(QFrame.Shape.NoFrame)
        self._body = QWidget()
        self._bl   = QVBoxLayout(self._body)
        self._bl.setContentsMargins(8, 8, 8, 8); self._bl.setSpacing(4)
        self._scroll.setWidget(self._body)
        root.addWidget(self._scroll, 1)
        self.setFixedWidth(240)

    def reload(self, mbots: list[dict]) -> None:
        existing_ids = {m["id"] for m in mbots}
        self.selected = [i for i in self.selected if i in existing_ids]
        if self.focused is not None and self.focused not in existing_ids:
            self.focused = None

        while self._bl.count():
            item = self._bl.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        self.rows.clear()

        for m in mbots:
            row = MbotRow(m, self.multi)
            row.clicked.connect(self._on_focus)
            row.toggled.connect(self._on_toggle)
            row.set_selected(m["id"] in self.selected)
            row.set_focused(m["id"] == self.focused)
            self.rows[m["id"]] = row
            self._bl.addWidget(row)
        self._bl.addStretch(1)
        self._count_pill.setText(str(len(mbots)))

    def _on_focus(self, mid: int) -> None:
        self.focused = mid
        for rid, row in self.rows.items():
            row.set_focused(rid == mid)
        self.focus_changed.emit(mid)
        self.setFocus()

    def _on_toggle(self, mid: int, on: bool) -> None:
        mods = QApplication.keyboardModifiers()
        if self.multi and not (mods & Qt.KeyboardModifier.ControlModifier):
            self.selected = [mid] if on else []
            for rid, row in self.rows.items():
                row.set_selected(rid == mid and on)
        else:
            if on and mid not in self.selected: self.selected.append(mid)
            elif not on and mid in self.selected: self.selected.remove(mid)
            self.rows[mid].set_selected(on)
        self.selection_changed.emit(self.selected)

    def keyPressEvent(self, e) -> None:
        ids = list(self.rows.keys())
        if not ids: return super().keyPressEvent(e)
        cur = ids.index(self.focused) if self.focused in ids else 0
        if e.key() == Qt.Key.Key_Down:
            self._on_focus(ids[min(len(ids) - 1, cur + 1)]); return
        if e.key() == Qt.Key.Key_Up:
            self._on_focus(ids[max(0, cur - 1)]); return
        if e.key() == Qt.Key.Key_Space and self.multi:
            self._on_toggle(self.focused, self.focused not in self.selected); return
        super().keyPressEvent(e)

    def select_all(self) -> None:
        self.selected = list(self.rows.keys())
        for row in self.rows.values(): row.set_selected(True)
        self.selection_changed.emit(self.selected)

    def clear(self) -> None:
        self.selected = []
        for row in self.rows.values(): row.set_selected(False)
        self.selection_changed.emit(self.selected)


# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------
class CharCard(QFrame):
    def __init__(self, mbot: dict):
        super().__init__()
        self.mbot_id = mbot["id"]
        self.setObjectName("CharCard")
        self.setFixedHeight(34)

        lay = QHBoxLayout(self); lay.setContentsMargins(10, 5, 10, 5); lay.setSpacing(8)
        color = T['text_mute'] if mbot.get("is_dc") else T['text']
        name  = QLabel(mbot["char"])
        name.setStyleSheet(f"font-weight:600; font-size:12px; color:{color};")
        name.setFixedWidth(68)
        lay.addWidget(name)

        def _bar(label: str, color_key: str):
            lbl = QLabel(label)
            lbl.setStyleSheet(f"color:{T['text_mute']}; font-size:10px;")
            lbl.setFixedWidth(16)
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

        self.kph_lbl = QLabel(); self.kph_lbl.setFixedWidth(80)
        lay.addWidget(self.kph_lbl)
        self.refresh(mbot)

    def refresh(self, mbot: dict) -> None:
        is_dc = bool(mbot.get("is_dc"))
        self.hp_bar.setValue(0 if is_dc else int(mbot.get("hp", 0)))
        self.mp_bar.setValue(0 if is_dc else int(mbot.get("mp", 0)))
        kph = mbot.get("kph", "–")
        self.kph_lbl.setText(
            f"<span style='color:{T['text_mute']};font-size:10px;'>K/h</span> "
            f"<span style='color:{T['accent']};font-family:\"JetBrains Mono\",monospace;"
            f"font-size:11px;font-weight:600;'>{kph}</span>"
        )


class DashboardPanel(QWidget):
    log_event = pyqtSignal(str, str)

    _BUTTON_GRID = [
        ("refresh",      "Refresh mBots",    None,  0, 0),
        ("showHide",     "Show/Hide mBots",  None,  0, 1),
        ("killBot",      "Kill mBots",       None,  0, 2),
        ("startClient",  "Start client",     None,  1, 0),
        ("showHideCli",  "Show/Hide Client", None,  1, 1),
        ("killClient",   "Kill client",      None,  1, 2),
        ("logoff",       "Log Off",          None,  1, 3),
        ("reset",        "Reset",            None,  2, 0),
        ("getPos",       "Get Position",     None,  2, 1),
        ("startTrain",   "Start Training",   None,  2, 2),
        ("stopTrain",    "Stop Training",    None,  2, 3),
    ]

    def __init__(self):
        super().__init__()
        self._mbots: list[dict] = []
        root = QVBoxLayout(self); root.setContentsMargins(0, 0, 0, 0); root.setSpacing(0)

        top = QHBoxLayout(); top.setContentsMargins(0, 0, 0, 0); top.setSpacing(0)

        self.list_col = MbotListColumn(multi=True)
        self.list_col.selection_changed.connect(self._on_sel)
        top.addWidget(self.list_col)

        col_chars = QFrame(); col_chars.setObjectName("Col")
        cc = QVBoxLayout(col_chars); cc.setContentsMargins(0, 0, 0, 0); cc.setSpacing(0)
        hdr_chars = QFrame()
        hdr_chars.setStyleSheet(f"background:{T['bg_deep']}; border-bottom:1px solid {T['border']};")
        hc = QHBoxLayout(hdr_chars); hc.setContentsMargins(12, 8, 12, 8)
        hc.addWidget(QLabel("CHARACTERS")); hc.addStretch(1)
        self.char_pill = QLabel("0 online"); self.char_pill.setObjectName("Pill")
        hc.addWidget(self.char_pill)
        cc.addWidget(hdr_chars)
        sc = QScrollArea(); sc.setWidgetResizable(True); sc.setFrameShape(QFrame.Shape.NoFrame)
        self._char_inner  = QWidget()
        self._char_layout = QVBoxLayout(self._char_inner)
        self._char_layout.setContentsMargins(10, 10, 10, 10); self._char_layout.setSpacing(6)
        self._char_layout.addStretch(1)
        sc.setWidget(self._char_inner)
        cc.addWidget(sc, 1)
        top.addWidget(col_chars, 1)
        self._char_cards: dict[int, CharCard] = {}

        top_w = QWidget(); top_w.setLayout(top)
        root.addWidget(top_w, 1)

        actions = QFrame()
        actions.setStyleSheet(f"background:{T['bg_deep']}; border-top:1px solid {T['border']};")
        af = QVBoxLayout(actions); af.setContentsMargins(0, 0, 0, 0); af.setSpacing(0)

        act_hdr = QFrame()
        act_hdr.setStyleSheet(f"background:{T['bg_panel']}; border-bottom:1px solid {T['border']};")
        ah = QHBoxLayout(act_hdr); ah.setContentsMargins(12, 6, 12, 6)
        ah.addWidget(QLabel("ACTIONS"))
        self.sel_pill = QLabel("0 selected"); self.sel_pill.setObjectName("Pill")
        ah.addWidget(self.sel_pill); ah.addStretch(1)
        af.addWidget(act_hdr)

        grid_w = QWidget()
        grid   = QGridLayout(grid_w); grid.setContentsMargins(10, 8, 10, 8); grid.setSpacing(6)
        for c in range(4): grid.setColumnStretch(c, 1)
        for bid, label, kind, row, col in self._BUTTON_GRID:
            btn = QPushButton(label)
            if kind: btn.setProperty(kind, True)
            btn.style().unpolish(btn); btn.style().polish(btn)
            btn.clicked.connect(lambda _, b=bid, l=label: self._fire(b, l))
            grid.addWidget(btn, row, col)
        af.addWidget(grid_w)
        root.addWidget(actions)

    def update_mbots(self, mbots: list[dict]) -> None:
        self._mbots = mbots
        self.list_col.reload(mbots)

        while self._char_layout.count():
            item = self._char_layout.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        self._char_cards.clear()

        for m in mbots:
            card = CharCard(m)
            self._char_cards[m["id"]] = card
            self._char_layout.addWidget(card)
        self._char_layout.addStretch(1)

        online = sum(1 for m in mbots if not m.get("is_dc"))
        self.char_pill.setText(f"{online} online")

    def refresh_cards(self, mbots: list[dict]) -> None:
        self._mbots = mbots
        for m in mbots:
            card = self._char_cards.get(m["id"])
            if card:
                card.refresh(m)
        online = sum(1 for m in mbots if not m.get("is_dc"))
        self.char_pill.setText(f"{online} online")

    def _on_sel(self, sel: list) -> None:
        self.sel_pill.setText(f"{len(sel)} selected")

    def _selected_ids(self) -> list[int]:
        return list(self.list_col.selected)

    def _fire(self, bid: str, label: str) -> None:
        sel_ids = self._selected_ids()
        names   = ", ".join(m["char"] for m in self._mbots if m["id"] in sel_ids) or "(none)"

        if bid == "refresh":
            self.log_event.emit("Refreshing mBot list…", "info")
            return

        if not sel_ids:
            self.log_event.emit(f"{label} — no mBots selected", "warn")
            return

        if bid in ("killBot", "killClient"):
            if QMessageBox.question(None, "Confirm", f"{label}: {names}?") != QMessageBox.StandardButton.Yes:
                return

        _CMD = {
            "showHide":    "show_hide_mbot",
            "killBot":     "kill_mbot",
            "startClient": "start_client",
            "showHideCli": "show_hide_client",
            "killClient":  "kill_client",
            "logoff":      "log_off",
            "reset":       "reset",
            "getPos":      "get_position",
            "startTrain":  "start_training",
            "stopTrain":   "stop_training",
        }
        action = _CMD.get(bid)
        if not action:
            return

        for mid in sel_ids:
            db.enqueue_command(action, mid)

        kind = "warn" if bid in ("killBot", "killClient", "stopTrain") else (
               "ok"   if bid in ("startTrain", "startClient") else "info")
        self.log_event.emit(f"{label} → {names}", kind)


# ---------------------------------------------------------------------------
# Account
# ---------------------------------------------------------------------------
class AccountPanel(QWidget):
    log_event = pyqtSignal(str, str)

    def __init__(self):
        super().__init__()
        self._accounts: list[dict] = load_accounts()
        self._item_changed_connected = False

        root = QVBoxLayout(self); root.setContentsMargins(16, 14, 16, 14); root.setSpacing(12)

        head = QHBoxLayout(); head.setSpacing(12)
        text_col = QVBoxLayout(); text_col.setSpacing(2)
        text_col.addWidget(QLabel("Accounts", objectName="PanelTitle"))
        self.sub_label = QLabel(); self.sub_label.setObjectName("PanelSub")
        text_col.addWidget(self.sub_label)
        head.addLayout(text_col, 1)

        login_btn = QPushButton("  Login selected  ")
        login_btn.setProperty("primary", True); login_btn.style().unpolish(login_btn); login_btn.style().polish(login_btn)
        login_btn.setFixedHeight(32); login_btn.clicked.connect(self._login_selected)
        hide_btn = QPushButton("  Hide mBots  ")
        hide_btn.setFixedHeight(32); hide_btn.clicked.connect(self._hide_mbots)
        head.addWidget(hide_btn,  0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        head.addWidget(login_btn, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        root.addLayout(head)

        tb = QHBoxLayout(); tb.setSpacing(6)
        sa = QPushButton("Select all");      sa.clicked.connect(self._select_all)
        ca = QPushButton("Clear all");       ca.clicked.connect(self._clear_all)
        rm = QPushButton("Remove selected"); rm.setProperty("danger", True)
        rm.style().unpolish(rm); rm.style().polish(rm); rm.clicked.connect(self._remove_selected)
        self.sel_pill = QLabel("0 selected"); self.sel_pill.setObjectName("Pill")
        tb.addWidget(sa); tb.addWidget(ca); tb.addWidget(rm); tb.addStretch(1); tb.addWidget(self.sel_pill)
        root.addLayout(tb)

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

        card = QFrame(); card.setObjectName("SignupCard")
        cl = QVBoxLayout(card); cl.setContentsMargins(12, 10, 12, 10); cl.setSpacing(6)
        cl.addWidget(QLabel("Add account", styleSheet="font-size:12px; font-weight:600;"))

        self.in_user = QLineEdit(placeholderText="Username")
        self.in_pass = QLineEdit(placeholderText="Password")
        self.in_pass.setEchoMode(QLineEdit.EchoMode.Password)
        self.in_char = QLineEdit(placeholderText="Character (exact, case-sensitive)")
        self.in_path = QLineEdit(placeholderText=r"C:\MBot\mbot.exe")
        browse_btn = QPushButton("Browse…"); browse_btn.clicked.connect(self._browse_mbot)
        browse_btn.setFixedWidth(70)

        r1 = QHBoxLayout(); r1.setSpacing(8)
        r1.addWidget(self.in_user, 1); r1.addWidget(self.in_pass, 1)
        cl.addLayout(r1)
        r2 = QHBoxLayout(); r2.setSpacing(8)
        r2.addWidget(self.in_char, 1); r2.addWidget(self.in_path, 1); r2.addWidget(browse_btn)
        cl.addLayout(r2)
        r3 = QHBoxLayout(); r3.setSpacing(6)
        add_btn = QPushButton("Add account"); add_btn.setProperty("primary", True)
        add_btn.style().unpolish(add_btn); add_btn.style().polish(add_btn)
        add_btn.clicked.connect(self._add)
        clr_btn = QPushButton("Clear"); clr_btn.clicked.connect(self._clear_form)
        r3.addWidget(add_btn); r3.addWidget(clr_btn); r3.addStretch(1)
        cl.addLayout(r3)
        root.addWidget(card)

    def _refresh_table(self) -> None:
        self.table.blockSignals(True)
        self.table.setRowCount(len(self._accounts))
        for i, a in enumerate(self._accounts):
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
            chk.setCheckState(Qt.CheckState.Unchecked)
            chk.setData(Qt.ItemDataRole.UserRole, i)
            self.table.setItem(i, 0, chk)
            self.table.setItem(i, 1, QTableWidgetItem(str(i + 1)))
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
        self.sub_label.setText(f"{len(self._accounts)} saved accounts. Each is bound to a .mbot profile file.")

    def _update_pill(self) -> None:
        n = sum(1 for r in range(self.table.rowCount())
                if (it := self.table.item(r, 0)) and it.checkState() == Qt.CheckState.Checked)
        self.sel_pill.setText(f"{n} selected")

    def _selected_indices(self) -> list[int]:
        return [
            self.table.item(r, 0).data(Qt.ItemDataRole.UserRole)
            for r in range(self.table.rowCount())
            if (it := self.table.item(r, 0)) and it.checkState() == Qt.CheckState.Checked
        ]

    def _select_all(self) -> None:
        self.table.blockSignals(True)
        for r in range(self.table.rowCount()):
            it = self.table.item(r, 0)
            if it: it.setCheckState(Qt.CheckState.Checked)
        self.table.blockSignals(False); self._update_pill()

    def _clear_all(self) -> None:
        self.table.blockSignals(True)
        for r in range(self.table.rowCount()):
            it = self.table.item(r, 0)
            if it: it.setCheckState(Qt.CheckState.Unchecked)
        self.table.blockSignals(False); self._update_pill()

    def _remove_selected(self) -> None:
        indices = self._selected_indices()
        if not indices: return
        names = ", ".join(self._accounts[i]["username"] for i in indices if i < len(self._accounts))
        if QMessageBox.question(self, "Confirm", f"Delete {len(indices)} account(s): {names}?") != QMessageBox.StandardButton.Yes:
            return
        for i in sorted(indices, reverse=True):
            if i < len(self._accounts): self._accounts.pop(i)
        save_accounts(self._accounts)
        self._refresh_table()
        self.log_event.emit(f"Removed {len(indices)} account(s)", "warn")

    def _login_selected(self) -> None:
        indices = self._selected_indices()
        if not indices:
            QMessageBox.information(self, "No selection", "Please select at least one account."); return
        names = ", ".join(self._accounts[i]["username"] for i in indices if i < len(self._accounts))
        db.enqueue_command("login", params={"indices": indices})
        self.log_event.emit(f"Login queued → {names}", "ok")

    def _hide_mbots(self) -> None:
        indices = self._selected_indices()
        db.enqueue_command("hide_mbots", params={"indices": indices} if indices else None)
        self.log_event.emit("Hide mBots queued", "info")

    def _browse_mbot(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Select mBot executable", "",
            "Applications (*.exe);;All files (*)",
            options=QFileDialog.Option.DontUseNativeDialog)
        if path: self.in_path.setText(os.path.normpath(path))

    def _add(self) -> None:
        u    = self.in_user.text().strip()
        p    = self.in_pass.text().strip()
        char = self.in_char.text().strip()
        path = self.in_path.text().strip()
        if not u or not p:
            QMessageBox.warning(self, "Missing fields", "Username and password are required."); return
        if not char:
            QMessageBox.warning(self, "Missing fields", "Character name is required."); return
        if any(a["username"] == u for a in self._accounts):
            QMessageBox.critical(self, "Error", "Username already exists!"); return
        self._accounts.append({
            "username": u,
            "password": base64.b64encode(p.encode()).decode(),
            "character": char,
            "mbot_file_path": path,
        })
        save_accounts(self._accounts)
        self._refresh_table()
        self.log_event.emit(f"Added account '{u}' ({char})", "ok")
        self._clear_form()

    def _clear_form(self) -> None:
        for w in (self.in_user, self.in_pass, self.in_char, self.in_path): w.clear()


# ---------------------------------------------------------------------------
# Chat
# ---------------------------------------------------------------------------
class ChatPanel(QWidget):
    def __init__(self):
        super().__init__()
        self._focused_id: Optional[int] = None
        self._active_ch  = "Allchat"
        self.tab_buttons: dict[str, QPushButton] = {}

        root = QHBoxLayout(self); root.setContentsMargins(0, 0, 0, 0); root.setSpacing(0)

        self.list_col = MbotListColumn(multi=False)
        self.list_col.focus_changed.connect(self._on_focus)
        root.addWidget(self.list_col)

        right = QFrame(); right.setObjectName("Col")
        rl    = QVBoxLayout(right); rl.setContentsMargins(0, 0, 0, 0); rl.setSpacing(0)

        hdr = QFrame()
        hdr.setStyleSheet(f"background:{T['bg_deep']}; border-bottom:1px solid {T['border']};")
        hh  = QHBoxLayout(hdr); hh.setContentsMargins(12, 8, 12, 8)
        hh.addWidget(QLabel("CHAT")); hh.addStretch(1)
        self.chat_pill = QLabel("—"); self.chat_pill.setObjectName("Pill")
        hh.addWidget(self.chat_pill)
        rl.addWidget(hdr)

        tabs = QFrame()
        tabs.setStyleSheet(f"background:{T['bg_deep']}; border-bottom:1px solid {T['border']};")
        tw = QHBoxLayout(tabs); tw.setContentsMargins(10, 8, 10, 8); tw.setSpacing(4)
        for ch in CHAT_BUTTON_TEXTS:
            b = QPushButton(ch); b.setObjectName("ChatTab")
            b.setProperty("active", ch == self._active_ch)
            b.style().unpolish(b); b.style().polish(b)
            b.clicked.connect(lambda _, c=ch: self._set_channel(c))
            self.tab_buttons[ch] = b; tw.addWidget(b)
        tw.addStretch(1)
        rl.addWidget(tabs)

        self.stream = QPlainTextEdit(); self.stream.setObjectName("ChatStream")
        self.stream.setReadOnly(True)
        rl.addWidget(self.stream, 1)
        root.addWidget(right, 1)

    def update_mbots(self, mbots: list[dict]) -> None:
        self.list_col.reload(mbots)
        self._refresh_chat()

    def _on_focus(self, mid: int) -> None:
        self._focused_id = mid
        self._refresh_chat()

    def _set_channel(self, ch: str) -> None:
        self._active_ch = ch
        for c, b in self.tab_buttons.items():
            b.setProperty("active", c == ch)
            b.style().unpolish(b); b.style().polish(b)
        self._refresh_chat()

    def _refresh_chat(self) -> None:
        if self._focused_id is None:
            return
        content = db.get_chat(self._focused_id, self._active_ch) or ""
        if self.stream.toPlainText() != content:
            self.stream.setPlainText(content)
            self.stream.verticalScrollBar().setValue(self.stream.verticalScrollBar().maximum())


# ---------------------------------------------------------------------------
# Inventory
# ---------------------------------------------------------------------------
class InventoryPanel(QWidget):
    def __init__(self):
        super().__init__()
        self._focused_id: Optional[int] = None
        self._mbots: list[dict] = []

        root = QHBoxLayout(self); root.setContentsMargins(0, 0, 0, 0); root.setSpacing(0)

        self.list_col = MbotListColumn(multi=False)
        self.list_col.focus_changed.connect(self._on_focus)
        root.addWidget(self.list_col)

        right = QFrame(); right.setObjectName("Col")
        rl    = QVBoxLayout(right); rl.setContentsMargins(0, 0, 0, 0); rl.setSpacing(0)

        hdr = QFrame()
        hdr.setStyleSheet(f"background:{T['bg_deep']}; border-bottom:1px solid {T['border']};")
        hh  = QHBoxLayout(hdr); hh.setContentsMargins(12, 8, 12, 8)
        hh.addWidget(QLabel("INVENTORY & LOG")); hh.addStretch(1)
        self.head_pill = QLabel("—"); self.head_pill.setObjectName("Pill")
        hh.addWidget(self.head_pill)
        rl.addWidget(hdr)

        body_layout = QHBoxLayout(); body_layout.setContentsMargins(0, 0, 0, 0); body_layout.setSpacing(0)

        # Inventory column
        inv_wrap = QFrame()
        inv_wrap.setStyleSheet(f"background:{T['bg_window']}; border-right:1px solid {T['border']};")
        iv = QVBoxLayout(inv_wrap); iv.setContentsMargins(0, 0, 0, 0); iv.setSpacing(0)
        inv_hdr = QFrame()
        inv_hdr.setStyleSheet(f"background:{T['bg_panel']}; border-bottom:1px solid {T['border']};")
        ih = QHBoxLayout(inv_hdr); ih.setContentsMargins(10, 5, 10, 5); ih.setSpacing(8)
        inv_title = QLabel("INVENTORY")
        inv_title.setStyleSheet(f"color:{T['text_dim']};font-weight:600;font-size:10px;letter-spacing:0.5px;")
        self.inv_combo = QComboBox(); self.inv_combo.addItems(INVENTORY_OPTIONS)
        self.inv_combo.setCurrentText("Inventory"); self.inv_combo.setFixedWidth(110)
        refresh_inv_btn = QPushButton("↻"); refresh_inv_btn.setFixedWidth(28)
        refresh_inv_btn.clicked.connect(self._request_inventory)
        ih.addWidget(inv_title); ih.addStretch(1)
        ih.addWidget(self.inv_combo); ih.addWidget(refresh_inv_btn)
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
        ev = QVBoxLayout(ev_wrap); ev.setContentsMargins(0, 0, 0, 0); ev.setSpacing(0)
        ev_hdr = QFrame()
        ev_hdr.setStyleSheet(f"background:{T['bg_panel']}; border-bottom:1px solid {T['border']};")
        eh = QHBoxLayout(ev_hdr); eh.setContentsMargins(10, 5, 10, 5); eh.setSpacing(10)
        ev_title = QLabel("EVENT LOG")
        ev_title.setStyleSheet(f"color:{T['text_dim']};font-weight:600;font-size:10px;letter-spacing:0.5px;")
        ev_clr = QPushButton("Clear"); ev_clr.setFixedHeight(20)
        ev_clr.setStyleSheet("padding:1px 8px; font-size:10px;")
        ev_clr.clicked.connect(lambda: self.ev_log.clear())
        refresh_log_btn = QPushButton("↻"); refresh_log_btn.setFixedWidth(28)
        refresh_log_btn.clicked.connect(self._request_log)
        eh.addWidget(ev_title); eh.addStretch(1); eh.addWidget(ev_clr); eh.addWidget(refresh_log_btn)
        ev.addWidget(ev_hdr)
        self.ev_log = QPlainTextEdit(); self.ev_log.setObjectName("InvLog"); self.ev_log.setReadOnly(True)
        ev.addWidget(self.ev_log, 1)
        body_layout.addWidget(ev_wrap, 1)

        body_w = QWidget(); body_w.setLayout(body_layout)
        rl.addWidget(body_w, 1)
        root.addWidget(right, 1)

    @staticmethod
    def _plain_column(title: str) -> dict:
        wrap = QFrame()
        wrap.setStyleSheet(f"background:{T['bg_window']}; border-right:1px solid {T['border']};")
        v = QVBoxLayout(wrap); v.setContentsMargins(0, 0, 0, 0); v.setSpacing(0)
        hdr = QLabel(title)
        hdr.setStyleSheet(
            f"background:{T['bg_panel']};color:{T['text_dim']};"
            f"border-bottom:1px solid {T['border']};"
            f"padding:6px 12px;font-weight:600;font-size:10px;letter-spacing:0.5px;")
        v.addWidget(hdr)
        log = QPlainTextEdit(); log.setObjectName("InvLog"); log.setReadOnly(True)
        v.addWidget(log, 1)
        return {"wrap": wrap, "log": log}

    def update_mbots(self, mbots: list[dict]) -> None:
        self._mbots = mbots
        self.list_col.reload(mbots)

    def _on_focus(self, mid: int) -> None:
        self._focused_id = mid
        char = next((m["char"] for m in self._mbots if m["id"] == mid), "—")
        self.head_pill.setText(char)
        self._render_inventory()
        self._render_log()

    def _request_inventory(self) -> None:
        if self._focused_id is None: return
        inv_type = self.inv_combo.currentText()
        db.enqueue_command("get_inventory", self._focused_id, {"inv_type": inv_type})
        QTimer.singleShot(2000, self._render_inventory)

    def _request_log(self) -> None:
        if self._focused_id is None: return
        db.enqueue_command("get_mbot_log", self._focused_id)
        QTimer.singleShot(1000, self._render_log)

    def _render_inventory(self) -> None:
        if self._focused_id is None: return
        inv_type = self.inv_combo.currentText()
        items    = db.get_inventory(self._focused_id, inv_type)

        totals: dict[str, int] = defaultdict(int)
        slots:  dict[str, int] = defaultdict(int)
        for line in items:
            m = re.search(r':\s*(.*?)\s*\((\d+)\s+pieces\)', line)
            if m:
                totals[m.group(1)] += int(m.group(2))
                slots[m.group(1)]  += 1
            else:
                totals[line] = 1; slots[line] = 1

        html = [f"<div style='margin-bottom:4px'><span style='color:{T['accent']};font-size:10px;"
                f"font-weight:600;'>[{inv_type}]</span></div>"]
        if totals:
            for item in sorted(totals):
                html.append(
                    f"<div><span style='color:{T['text']}'>{item}</span>"
                    f"<span style='color:{T['text_mute']}'> — </span>"
                    f"<span style='color:{T['accent']};font-family:monospace'>{totals[item]}</span>"
                    f"<span style='color:{T['text_mute']}'>pcs / {slots[item]} slots</span></div>"
                )
        else:
            html.append(f"<div style='color:{T['text_mute']}'>No data. Click ↻ to refresh.</div>")
        self.inv_log.clear(); self.inv_log.appendHtml("".join(html))

    def _render_log(self) -> None:
        if self._focused_id is None: return
        raw  = db.get_mbot_log(self._focused_id)
        html = [f"<div style='color:{T['text_dim']}'>{line}</div>"
                for line in raw.splitlines() if line.strip()]
        self.ev_log.clear()
        self.ev_log.appendHtml(
            "".join(html) if html
            else f"<div style='color:{T['text_mute']}'>No log. Click ↻ to load.</div>"
        )


# ---------------------------------------------------------------------------
# Update
# ---------------------------------------------------------------------------
class UpdatePanel(QWidget):
    log_event = pyqtSignal(str, str)

    def __init__(self):
        super().__init__()
        self._item_changed_connected = False

        root = QVBoxLayout(self); root.setContentsMargins(16, 14, 16, 14); root.setSpacing(12)

        head = QHBoxLayout(); head.setSpacing(12)
        text_col = QVBoxLayout(); text_col.setSpacing(2)
        text_col.addWidget(QLabel("SRO Updater", objectName="PanelTitle"))
        self.sub_label = QLabel(); self.sub_label.setObjectName("PanelSub")
        text_col.addWidget(self.sub_label)
        head.addLayout(text_col, 1)
        self.update_btn = QPushButton("  Run Update  ")
        self.update_btn.setProperty("primary", True)
        self.update_btn.style().unpolish(self.update_btn); self.update_btn.style().polish(self.update_btn)
        self.update_btn.setFixedHeight(32); self.update_btn.clicked.connect(self._run_selected)
        head.addWidget(self.update_btn, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        root.addLayout(head)

        tb = QHBoxLayout(); tb.setSpacing(6)
        sa = QPushButton("Select all");      sa.clicked.connect(self._select_all)
        ca = QPushButton("Clear all");       ca.clicked.connect(self._clear_all)
        rm = QPushButton("Remove selected"); rm.setProperty("danger", True)
        rm.style().unpolish(rm); rm.style().polish(rm); rm.clicked.connect(self._remove_selected)
        self.sel_pill = QLabel("0 selected"); self.sel_pill.setObjectName("Pill")
        tb.addWidget(sa); tb.addWidget(ca); tb.addWidget(rm); tb.addStretch(1); tb.addWidget(self.sel_pill)
        root.addLayout(tb)

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["", "#", "Silkroad.exe path"])
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setShowGrid(False)
        h = self.table.horizontalHeader()
        h.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        h.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        h.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table.setMinimumHeight(220)
        root.addWidget(self.table)
        self._refresh_table()

        card = QFrame(); card.setObjectName("SignupCard")
        cl = QVBoxLayout(card); cl.setContentsMargins(12, 10, 12, 10); cl.setSpacing(6)
        cl.addWidget(QLabel("Add Silkroad.exe path", styleSheet="font-size:12px; font-weight:600;"))
        r1 = QHBoxLayout(); r1.setSpacing(8)
        self.in_path = QLineEdit(placeholderText=r"C:\Silkroad\Silkroad.exe")
        browse_btn   = QPushButton("Browse…"); browse_btn.setFixedWidth(70)
        browse_btn.clicked.connect(self._browse)
        r1.addWidget(self.in_path, 1); r1.addWidget(browse_btn)
        cl.addLayout(r1)
        r2 = QHBoxLayout(); r2.setSpacing(6)
        add_btn = QPushButton("Add path"); add_btn.setProperty("primary", True)
        add_btn.style().unpolish(add_btn); add_btn.style().polish(add_btn); add_btn.clicked.connect(self._add)
        clr_btn = QPushButton("Clear"); clr_btn.clicked.connect(lambda: self.in_path.clear())
        r2.addWidget(add_btn); r2.addWidget(clr_btn); r2.addStretch(1)
        cl.addLayout(r2)
        root.addWidget(card)

        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet(f"color:{T['text_mute']}; font-size:11px;")
        root.addWidget(self.status_lbl)

    def _refresh_table(self) -> None:
        self.table.blockSignals(True)
        self.table.setRowCount(len(_updater_paths))
        for i, path in enumerate(_updater_paths):
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
            chk.setCheckState(Qt.CheckState.Unchecked)
            chk.setData(Qt.ItemDataRole.UserRole, i)
            self.table.setItem(i, 0, chk)
            self.table.setItem(i, 1, QTableWidgetItem(str(i + 1)))
            path_it = QTableWidgetItem(path)
            path_it.setForeground(QColor(T['text_dim'])); path_it.setToolTip(path)
            self.table.setItem(i, 2, path_it)
        self.table.blockSignals(False); self.table.resizeRowsToContents()
        if not self._item_changed_connected:
            self.table.itemChanged.connect(lambda it: it.column() == 0 and self._update_pill())
            self._item_changed_connected = True
        self._update_pill()
        self.sub_label.setText(f"{len(_updater_paths)} path(s) configured.")

    def _update_pill(self) -> None:
        n = sum(1 for r in range(self.table.rowCount())
                if (it := self.table.item(r, 0)) and it.checkState() == Qt.CheckState.Checked)
        self.sel_pill.setText(f"{n} selected")

    def _selected_indices(self) -> list[int]:
        return [
            self.table.item(r, 0).data(Qt.ItemDataRole.UserRole)
            for r in range(self.table.rowCount())
            if (it := self.table.item(r, 0)) and it.checkState() == Qt.CheckState.Checked
        ]

    def _select_all(self) -> None:
        self.table.blockSignals(True)
        for r in range(self.table.rowCount()):
            it = self.table.item(r, 0)
            if it: it.setCheckState(Qt.CheckState.Checked)
        self.table.blockSignals(False); self._update_pill()

    def _clear_all(self) -> None:
        self.table.blockSignals(True)
        for r in range(self.table.rowCount()):
            it = self.table.item(r, 0)
            if it: it.setCheckState(Qt.CheckState.Unchecked)
        self.table.blockSignals(False); self._update_pill()

    def _remove_selected(self) -> None:
        indices = self._selected_indices()
        if not indices: return
        if QMessageBox.question(self, "Confirm", f"Remove {len(indices)} path(s)?") != QMessageBox.StandardButton.Yes:
            return
        for i in sorted(indices, reverse=True):
            if i < len(_updater_paths): _updater_paths.pop(i)
        save_updater_paths(); self._refresh_table()
        self.log_event.emit(f"Removed {len(indices)} updater path(s)", "warn")

    def _browse(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Silkroad.exe", "",
            "Applications (*.exe);;All files (*)",
            options=QFileDialog.Option.DontUseNativeDialog)
        if path: self.in_path.setText(os.path.normpath(path))

    def _add(self) -> None:
        path = self.in_path.text().strip()
        if not path:
            QMessageBox.warning(self, "Missing path", "Please enter or browse to a Silkroad.exe path."); return
        if not os.path.exists(path):
            if QMessageBox.question(self, "Path not found", f"File not found:\n{path}\n\nAdd anyway?") != QMessageBox.StandardButton.Yes:
                return
        if path in _updater_paths:
            QMessageBox.information(self, "Duplicate", "This path is already in the list."); return
        _updater_paths.append(path); save_updater_paths()
        self._refresh_table(); self.in_path.clear()
        self.log_event.emit(f"Added updater path: {path}", "ok")

    def _run_selected(self) -> None:
        indices = self._selected_indices()
        if not indices:
            QMessageBox.information(self, "No selection", "Please select at least one path."); return
        db.enqueue_command("run_update", params={"indices": indices})
        self.status_lbl.setText("Update queued…")
        self.log_event.emit(f"Update queued for {len(indices)} path(s)", "accent")

    def run_all(self) -> None:
        db.enqueue_command("run_update")
        self.status_lbl.setText("Update queued (all)…")
        self.log_event.emit("Update queued for all paths", "accent")


# ---------------------------------------------------------------------------
# Log
# ---------------------------------------------------------------------------
class LogPanel(QWidget):
    def __init__(self):
        super().__init__()
        self._last_log_id = 0
        root = QVBoxLayout(self); root.setContentsMargins(0, 0, 0, 0); root.setSpacing(0)

        hdr = QFrame()
        hdr.setStyleSheet(f"background:{T['bg_panel']}; border-bottom:1px solid {T['border']};")
        lh  = QHBoxLayout(hdr); lh.setContentsMargins(12, 6, 12, 6)
        lh.addWidget(QLabel("LOG", styleSheet=f"color:{T['text_dim']};font-weight:600;font-size:10px;letter-spacing:0.5px;"))
        self.log_count = QLabel("0 entries"); self.log_count.setObjectName("Pill")
        clr_btn = QPushButton("Clear"); clr_btn.clicked.connect(self.clear)
        lh.addWidget(self.log_count); lh.addStretch(1); lh.addWidget(clr_btn)
        root.addWidget(hdr)

        self.log = QPlainTextEdit(); self.log.setObjectName("Log"); self.log.setReadOnly(True)
        root.addWidget(self.log, 1)

    def append(self, msg: str, kind: str = "info", who: Optional[str] = None) -> None:
        colors = {
            "info": T['text_dim'], "ok": T['success'], "warn": T['warn'],
            "err":  T['danger'],   "accent": T['accent'],
        }
        color    = colors.get(kind, T['text_dim'])
        prefix   = f"<span style='color:{T['text_mute']}'>[{now_ts()}]</span> "
        who_html = f"<span style='color:{T['accent']}'>{who}:</span> " if who else ""
        self.log.appendHtml(f"{prefix}{who_html}<span style='color:{color}'>{msg}</span>")
        count = int(self.log_count.text().split()[0]) + 1
        self.log_count.setText(f"{count} entries")
        self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())

    def load_from_db(self) -> None:
        rows = db.get_recent_logs(200)
        self.log.clear(); self._last_log_id = 0
        colors = {
            "info": T['text_dim'], "ok": T['success'], "warn": T['warn'],
            "err":  T['danger'],   "accent": T['accent'],
        }
        for row in reversed(rows):
            c = colors.get(row["kind"], T['text_dim'])
            who_html = f"<span style='color:{T['accent']}'>{row['who']}:</span> " if row.get("who") else ""
            self.log.appendHtml(
                f"<span style='color:{T['text_mute']}'>[{row['ts']}]</span> "
                f"{who_html}<span style='color:{c}'>{row['msg']}</span>"
            )
            if row["id"] > self._last_log_id:
                self._last_log_id = row["id"]
        self.log_count.setText(f"{len(rows)} entries")
        self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())

    def clear(self) -> None:
        self.log.clear(); self.log_count.setText("0 entries")


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MBot Manager")
        self.resize(900, 620)
        self.setStyleSheet(make_stylesheet(DARK))
        self._mbots: list[dict] = []
        self._mbots_key: list   = []

        central = QWidget(); self.setCentralWidget(central)
        root    = QVBoxLayout(central); root.setContentsMargins(0, 0, 0, 0); root.setSpacing(0)

        # Title bar
        title_bar = QFrame(); title_bar.setObjectName("TitleBar"); title_bar.setFixedHeight(32)
        tb = QHBoxLayout(title_bar); tb.setContentsMargins(12, 0, 0, 0); tb.setSpacing(8)
        self.title_text = QLabel()
        tb.addWidget(self.title_text); tb.addStretch(1)
        root.addWidget(title_bar)
        self._update_title(0)

        # Sidebar + content
        main = QHBoxLayout(); main.setContentsMargins(0, 0, 0, 0); main.setSpacing(0)

        sidebar = QFrame(); sidebar.setObjectName("Sidebar"); sidebar.setFixedWidth(96)
        sl      = QVBoxLayout(sidebar); sl.setContentsMargins(0, 4, 0, 0); sl.setSpacing(0)

        self.stack     = QStackedWidget()
        self.dash      = DashboardPanel()
        self.acc       = AccountPanel()
        self.chat      = ChatPanel()
        self.inv       = InventoryPanel()
        self.upd       = UpdatePanel()
        self.log_panel = LogPanel()

        for w in (self.dash, self.acc, self.chat, self.inv, self.upd, self.log_panel):
            self.stack.addWidget(w)

        self.dash.log_event.connect(self._append_log)
        self.acc.log_event.connect(self._append_log)
        self.upd.log_event.connect(self._append_log)

        self.nav_buttons: list[QPushButton] = []
        for i, label in enumerate(["Dashboard", "Account", "Chat", "Inventory", "Update", "Log"]):
            b = QPushButton(label); b.setObjectName("NavItem")
            b.setProperty("active", i == 0)
            b.style().unpolish(b); b.style().polish(b)
            b.clicked.connect(lambda _, idx=i: self._switch(idx))
            sl.addWidget(b); self.nav_buttons.append(b)
        sl.addStretch(1)
        main.addWidget(sidebar)

        right = QFrame()
        rl    = QVBoxLayout(right); rl.setContentsMargins(0, 0, 0, 0); rl.setSpacing(0)
        rl.addWidget(self.stack, 1)
        main.addWidget(right, 1)
        root.addLayout(main, 1)

        # Initial data load
        self._refresh_mbots()
        self.log_panel.load_from_db()
        self._append_log("MBot Manager (Qt) ready", "accent")

        # 500 ms tick — refresh HP/MP/KPH
        self._fast_timer = QTimer(self)
        self._fast_timer.timeout.connect(self._fast_tick)
        self._fast_timer.start(500)

        # 3 s tick — check for new mbot windows + new logs
        self._slow_timer = QTimer(self)
        self._slow_timer.timeout.connect(self._slow_tick)
        self._slow_timer.start(3_000)

    # ── Tick handlers ────────────────────────────────────────────────────
    def _fast_tick(self) -> None:
        mbots = db.get_all_mbots()
        if not mbots: return
        # Cheap path: same ids → just refresh cards
        key = [(m["id"], m.get("is_dc")) for m in mbots]
        if key == self._mbots_key:
            self._mbots = mbots
            self.dash.refresh_cards(mbots)
            return
        self._refresh_all(mbots)

    def _slow_tick(self) -> None:
        mbots = db.get_all_mbots()
        key   = [(m["id"], m.get("is_dc")) for m in mbots]
        if key != self._mbots_key:
            self._refresh_all(mbots)
        # Append any new log entries
        rows = db.get_recent_logs(200)
        if rows and rows[0]["id"] > self.log_panel._last_log_id:
            self.log_panel.load_from_db()
        # Refresh chat if visible
        if self.stack.currentIndex() == 2:
            self.chat._refresh_chat()

    def _refresh_mbots(self) -> None:
        self._refresh_all(db.get_all_mbots())

    def _refresh_all(self, mbots: list[dict]) -> None:
        self._mbots      = mbots
        self._mbots_key  = [(m["id"], m.get("is_dc")) for m in mbots]
        self.dash.update_mbots(mbots)
        self.chat.update_mbots(mbots)
        self.inv.update_mbots(mbots)
        online = sum(1 for m in mbots if not m.get("is_dc"))
        self._update_title(online)

    def _update_title(self, online: int) -> None:
        self.title_text.setText(
            f"<b>MBot Manager</b> <span style='color:{T['text_mute']}'>v0.2.0</span>"
            f"  <span style='background:rgba(0,0,0,0.2);padding:2px 8px;border-radius:3px;'>"
            f"<span style='color:{T['success']}'>●</span> {online} mBots online</span>"
        )

    def _switch(self, idx: int) -> None:
        self.stack.setCurrentIndex(idx)
        for i, b in enumerate(self.nav_buttons):
            b.setProperty("active", i == idx)
            b.style().unpolish(b); b.style().polish(b)

    def _append_log(self, msg: str, kind: str = "info", who: Optional[str] = None) -> None:
        self.log_panel.append(msg, kind, who)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main() -> None:
    global _updater_paths

    import traceback as _tb

    def _excepthook(exc_type, exc_val, exc_tb):
        msg = "".join(_tb.format_exception(exc_type, exc_val, exc_tb))
        print(msg, file=sys.__stderr__)
        try:
            with open("crash.log", "a") as f: f.write(msg + "\n")
        except Exception:
            pass
    sys.excepthook = _excepthook

    os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
    _updater_paths = load_updater_paths()

    db.init_db()

    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    pal = app.palette()
    pal.setColor(QPalette.ColorRole.Window,     QColor(T['bg_window']))
    pal.setColor(QPalette.ColorRole.Base,       QColor(T['bg_input']))
    pal.setColor(QPalette.ColorRole.Text,       QColor(T['text']))
    pal.setColor(QPalette.ColorRole.WindowText, QColor(T['text']))
    pal.setColor(QPalette.ColorRole.Button,     QColor(T['bg_panel']))
    pal.setColor(QPalette.ColorRole.ButtonText, QColor(T['text']))
    app.setPalette(pal)

    w = MainWindow(); w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
