"""
sandbox_toggle.py — Windows Sandbox Toggle
Requires: Python 3.11+, PyQt6, pywin32  (Windows only)
"""

import sys
import win32gui
import win32con
from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QLabel, QHBoxLayout
from PyQt6.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve, pyqtProperty, QPoint
from PyQt6.QtGui import QFont, QColor, QPainter, QPainterPath, QLinearGradient, QFontDatabase


# ── Win32 logic ──────────────────────────────────────────────────────────────

TARGET_CLASS = "WinUIDesktopWin32WindowClass"
TARGET_TITLE = "Windows Sandbox"


def find_sandbox_hwnd() -> int:
    result = []

    def callback(hwnd, _):
        cls   = win32gui.GetClassName(hwnd)
        title = win32gui.GetWindowText(hwnd)
        if cls == TARGET_CLASS and TARGET_TITLE in title:
            result.append(hwnd)

    win32gui.EnumWindows(callback, None)
    return result[0] if result else 0


def toggle_window(hwnd: int) -> bool:
    """Returns True if window is visible after triggering the toggle."""
    if not win32gui.IsWindow(hwnd):
        return False

    if win32gui.IsWindowVisible(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
        return False
    else:
        placement = win32gui.GetWindowPlacement(hwnd)
        if placement[1] == win32con.SW_SHOWMINIMIZED:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        else:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
        win32gui.SetForegroundWindow(hwnd)
        return True


# ── Toggle Button ─────────────────────────────────────────────────────────────

class ToggleButton(QWidget):
    def __init__(self, parent=None, on_toggle=None):
        super().__init__(parent)
        self.setFixedSize(72, 36)
        self._checked = False
        self._offset = 4
        self._on_toggle = on_toggle

        self._anim = QPropertyAnimation(self, b"offset", self)
        self._anim.setDuration(220)
        self._anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    @pyqtProperty(int)
    def offset(self):
        return self._offset

    @offset.setter
    def offset(self, v):
        self._offset = v
        self.update()

    def setChecked(self, val: bool):
        self._checked = val
        self._anim.setStartValue(self._offset)
        self._anim.setEndValue(40 if val else 4)
        self._anim.start()

    def isChecked(self):
        return self._checked

    def mousePressEvent(self, _):
        self.setChecked(not self._checked)
        if self._on_toggle:
            self._on_toggle()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)

        # Track
        track = QPainterPath()
        track.addRoundedRect(0, 0, 72, 36, 18, 18)
        if self._checked:
            grad = QLinearGradient(0, 0, 72, 0)
            grad.setColorAt(0, QColor("#00c896"))
            grad.setColorAt(1, QColor("#00e6b0"))
            p.fillPath(track, grad)
        else:
            p.fillPath(track, QColor("#2a2a3a"))

        # Thumb
        p.setBrush(QColor("#ffffff"))
        p.setPen(Qt.PenStyle.NoPen)
        p.drawEllipse(self._offset, 4, 28, 28)
        p.end()


# ── Status Dot ────────────────────────────────────────────────────────────────

class StatusDot(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(10, 10)
        self._color = QColor("#444455")
        self._pulse = 0.0

        self._pulse_anim = QPropertyAnimation(self, b"pulse", self)
        self._pulse_anim.setDuration(1200)
        self._pulse_anim.setStartValue(0.0)
        self._pulse_anim.setEndValue(1.0)
        self._pulse_anim.setLoopCount(-1)
        self._pulse_anim.setEasingCurve(QEasingCurve.Type.SineCurve)

    @pyqtProperty(float)
    def pulse(self):
        return self._pulse

    @pulse.setter
    def pulse(self, v):
        self._pulse = v
        self.update()

    def setStatus(self, found: bool, visible: bool):
        if not found:
            self._color = QColor("#ff4466")
            self._pulse_anim.stop()
        elif visible:
            self._color = QColor("#00e6b0")
            self._pulse_anim.start()
        else:
            self._color = QColor("#888899")
            self._pulse_anim.stop()
        self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        if self._pulse > 0:
            halo = QColor(self._color)
            halo.setAlphaF(0.25 * (1 - self._pulse))
            p.setBrush(halo)
            p.setPen(Qt.PenStyle.NoPen)
            r = int(5 + self._pulse * 5)
            cx, cy = self.width() // 2, self.height() // 2
            p.drawEllipse(cx - r, cy - r, r * 2, r * 2)
        p.setBrush(self._color)
        p.setPen(Qt.PenStyle.NoPen)
        p.drawEllipse(1, 1, 8, 8)
        p.end()


# ── Main Window ───────────────────────────────────────────────────────────────

class SandboxToggleApp(QWidget):
    def __init__(self):
        super().__init__()
        self.hwnd = 0
        self.is_visible = False
        self._setup_ui()
        self._apply_style()

        # Poll mỗi 1.5s để check trạng thái
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._poll)
        self.timer.start(1500)
        self._poll()

    def _setup_ui(self):
        self.setWindowTitle("Sandbox Toggle")
        self.setFixedSize(320, 200)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        # Drag support
        self._drag_pos = QPoint()

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(0)

        # Card
        self.card = QWidget(self)
        self.card.setObjectName("card")
        card_layout = QVBoxLayout(self.card)
        card_layout.setContentsMargins(24, 22, 24, 22)
        card_layout.setSpacing(16)

        # Header
        header = QHBoxLayout()
        header.setSpacing(8)

        self.dot = StatusDot(self.card)
        header.addWidget(self.dot, 0, Qt.AlignmentFlag.AlignVCenter)

        title = QLabel("Windows Sandbox")
        title.setObjectName("title")
        header.addWidget(title)
        header.addStretch()

        min_btn = QPushButton("─")
        min_btn.setObjectName("closeBtn")
        min_btn.setFixedSize(24, 24)
        min_btn.clicked.connect(self.showMinimized)
        header.addWidget(min_btn)

        close_btn = QPushButton("✕")
        close_btn.setObjectName("closeBtn")
        close_btn.setFixedSize(24, 24)
        close_btn.clicked.connect(self.close)
        header.addWidget(close_btn)

        card_layout.addLayout(header)

        # Status text
        self.status_label = QLabel("Searching...")
        self.status_label.setObjectName("status")
        card_layout.addWidget(self.status_label)

        card_layout.addSpacing(4)

        # Toggle row
        toggle_row = QHBoxLayout()
        toggle_row.setSpacing(12)

        hide_lbl = QLabel("Hide")
        hide_lbl.setObjectName("toggleLabel")
        toggle_row.addWidget(hide_lbl)

        self.toggle = ToggleButton(self, on_toggle=self.on_toggle)
        toggle_row.addWidget(self.toggle)

        show_lbl = QLabel("Show")
        show_lbl.setObjectName("toggleLabel")
        toggle_row.addWidget(show_lbl)
        toggle_row.addStretch()

        card_layout.addLayout(toggle_row)

        root.addWidget(self.card)

    def _apply_style(self):
        self.setStyleSheet("""
            QWidget#card {
                background: #13131f;
                border-radius: 18px;
                border: 1px solid #2a2a40;
            }
            QLabel#title {
                font-family: 'Segoe UI', sans-serif;
                font-size: 15px;
                font-weight: 600;
                color: #e8e8f0;
                letter-spacing: 0.3px;
            }
            QLabel#status {
                font-family: 'Segoe UI', sans-serif;
                font-size: 12px;
                color: #666680;
                letter-spacing: 0.2px;
            }
            QLabel#toggleLabel {
                font-family: 'Segoe UI', sans-serif;
                font-size: 12px;
                color: #55556a;
            }
            QPushButton#closeBtn {
                background: transparent;
                border: none;
                color: #444455;
                font-size: 12px;
                border-radius: 12px;
            }
            QPushButton#closeBtn:hover {
                background: #2a2a3a;
                color: #e8e8f0;
            }
        """)

    # ── drag window ───────────────────────────────────────────────────────────
    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self._drag_pos = e.globalPosition().toPoint() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if e.buttons() == Qt.MouseButton.LeftButton and not self._drag_pos.isNull():
            self.move(e.globalPosition().toPoint() - self._drag_pos)

    # ── logic ─────────────────────────────────────────────────────────────────
    def _poll(self):
        self.hwnd = find_sandbox_hwnd()
        if not self.hwnd:
            self.dot.setStatus(False, False)
            self.status_label.setText("Process not found")
            self.status_label.setStyleSheet("color: #ff4466; font-size: 12px;")
            self.toggle.setEnabled(False)
            return

        self.toggle.setEnabled(True)
        self.is_visible = bool(win32gui.IsWindowVisible(self.hwnd))
        self.dot.setStatus(True, self.is_visible)

        # Sync toggle state no trigger animation
        if self.toggle.isChecked() != self.is_visible:
            self.toggle.setChecked(self.is_visible)

        hwnd_hex = hex(self.hwnd)
        state    = "is visible" if self.is_visible else "is hidden"
        self.status_label.setText(f"HWND {hwnd_hex} — {state}")
        self.status_label.setStyleSheet("color: #555570; font-size: 12px;")

    def on_toggle(self):
        if not self.hwnd:
            return
        self.is_visible = toggle_window(self.hwnd)
        self.dot.setStatus(True, self.is_visible)
        state = "is visible" if self.is_visible else "is hidden"
        self.status_label.setText(f"HWND {hex(self.hwnd)} — {state}")


# ── Entry ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = SandboxToggleApp()
    win.show()
    sys.exit(app.exec())