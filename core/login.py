"""
core/login.py — Account login sequence.

Converts the QTimer-based AccountPanel flow to a blocking thread:
  1. Launch all selected mBot executables
  2. For each account: find matching mBot window, click Start Client
  3. Wait for SRO_Client window, click through server select, enter credentials
  4. Hide windows, move to next account
  5. After all accounts: wait 60 s, start training on all mBots

No Qt dependency.
"""

import base64
import ctypes
import os
import subprocess
import time
import threading
from typing import Callable, Optional

from core.window import (
    MBotWindow,
    WIN32_AVAILABLE,
    win32gui,
    win32con,
    win32process,
    findwindows,
    auto,
)

LogFn = Callable[[str, str], None]

_NOOP_LOG: LogFn = lambda msg, kind: None

# ---------------------------------------------------------------------------
# Firewall helper
# ---------------------------------------------------------------------------
def ensure_firewall(exe_path: str, username: str, log: LogFn = _NOOP_LOG) -> None:
    rule_name = f"{username}_{os.path.basename(exe_path)}"
    try:
        check = subprocess.run(
            ["netsh", "advfirewall", "firewall", "show", "rule", f"name={rule_name}"],
            capture_output=True, text=True,
        )
        if check.returncode == 0 and "No rules match" not in check.stdout:
            log(f"Firewall rule already exists for {rule_name}", "info")
            return
        subprocess.run([
            "netsh", "advfirewall", "firewall", "add", "rule",
            f"name={rule_name}", "dir=in", "action=allow",
            f"program={exe_path}", "profile=public", "enable=yes",
        ], capture_output=True)
        log(f"Firewall rule added for {rule_name}", "ok")
    except Exception as e:
        log(f"Firewall setup failed: {e}", "warn")


# ---------------------------------------------------------------------------
# Hide mBot windows by process name
# ---------------------------------------------------------------------------
def hide_mbots_by_path(accounts: list[dict], indices: list[int], log: LogFn = _NOOP_LOG) -> int:
    if not WIN32_AVAILABLE:
        log("Win32 not available — cannot hide mBot windows", "warn")
        return 0
    try:
        import psutil
    except ImportError:
        log("psutil not available", "warn")
        return 0

    hidden = 0
    for idx in indices:
        if idx >= len(accounts):
            continue
        mbot_path = accounts[idx].get("mbot_file_path", "")
        if not mbot_path:
            continue
        file_name = os.path.basename(mbot_path).lower()
        try:
            pids = {
                p.info["pid"]
                for p in psutil.process_iter(["pid", "name"])
                if p.info["name"] and p.info["name"].lower() == file_name
            }

            def _enum(hwnd, _data, _pids=pids):
                nonlocal hidden
                if not win32gui.IsWindowVisible(hwnd):
                    return
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                except Exception:
                    return
                if pid not in _pids:
                    return
                ctypes.windll.user32.ShowWindow(hwnd, 0)
                hidden += 1
                log(f"Hidden: '{win32gui.GetWindowText(hwnd)}'", "info")

            win32gui.EnumWindows(_enum, None)
        except Exception as e:
            log(f"Hide error for {file_name}: {e}", "err")

    log(f"Hide mBots — {hidden} window(s) hidden", "ok" if hidden else "warn")
    return hidden


# ---------------------------------------------------------------------------
# Start training on all found mBot windows
# ---------------------------------------------------------------------------
def start_training_all(log: LogFn = _NOOP_LOG) -> None:
    if not WIN32_AVAILABLE:
        return
    try:
        mbot_list = findwindows.find_elements(class_name="#32770")
        for elem in mbot_list:
            w = MBotWindow(elem)
            w.start_training()
            time.sleep(0.1)
            w.start_training()
        log(f"start_training sent to {len(mbot_list)} mBot(s)", "ok")
    except Exception as e:
        log(f"start_training_all error: {e}", "err")


# ---------------------------------------------------------------------------
# Login sequence (blocking — run in a daemon thread)
# ---------------------------------------------------------------------------
class LoginRunner:
    """Runs the full login sequence for a set of accounts in a daemon thread."""

    def __init__(self):
        self._lock    = threading.Lock()
        self._running = False

    def run(self, accounts: list[dict], indices: list[int], log: LogFn = _NOOP_LOG) -> bool:
        with self._lock:
            if self._running:
                log("Login already running", "warn")
                return False
            self._running = True
        t = threading.Thread(
            target=self._sequence,
            args=(accounts, indices, log),
            daemon=True,
            name="login",
        )
        t.start()
        return True

    @property
    def running(self) -> bool:
        with self._lock:
            return self._running

    # ── Internal steps ────────────────────────────────────────────────────
    def _sequence(self, accounts: list[dict], indices: list[int], log: LogFn) -> None:
        try:
            self._launch_all(accounts, indices, log)
            time.sleep(1)
            for seq, idx in enumerate(indices):
                if idx >= len(accounts):
                    continue
                acc = accounts[idx]
                self._login_one(seq, idx, acc, log)
            log("Login sequence finished — waiting 60s before starting training", "ok")
            time.sleep(60)
            start_training_all(log)
        except Exception as e:
            log(f"Login sequence error: {e}", "err")
        finally:
            with self._lock:
                self._running = False

    def _launch_all(self, accounts: list[dict], indices: list[int], log: LogFn) -> None:
        for idx in indices:
            if idx >= len(accounts):
                continue
            acc      = accounts[idx]
            username = acc["username"]

            for title in [
                f"[{username}] mBot v1.12b (vSRO 110)",
                f"[{username} - DC] mBot v1.12b (vSRO 110)",
            ]:
                if findwindows.find_elements(class_name="#32770", title=title, visible_only=False):
                    log(f"mBot already open for {username}", "info")
                    time.sleep(0.2)
                    break
            else:
                mbot_path = acc.get("mbot_file_path", "")
                if not mbot_path or not os.path.exists(mbot_path):
                    log(f"mBot path not found for {username}: {mbot_path}", "err")
                    time.sleep(0.2)
                    continue
                folder   = os.path.normpath(os.path.dirname(mbot_path))
                vsro_exe = os.path.join(folder, "mBot_vSRO110.exe")
                if os.path.exists(vsro_exe):
                    ensure_firewall(vsro_exe, username, log)
                subprocess.Popen(mbot_path, cwd=folder)
                log(f"Launched mBot for {username}", "info")
                time.sleep(1)

        log("All mBots launched", "ok")

    def _login_one(self, seq: int, idx: int, acc: dict, log: LogFn) -> None:
        username  = acc["username"]
        character = acc.get("character", username)
        log(f"[{seq + 1}] Starting login for {username} / {character}", "info")

        mbot_hwnd = self._find_mbot_hwnd(username, character, log)
        if mbot_hwnd is None:
            log(f"Could not find mBot window for {username}, skipping", "err")
            return

        sro_hwnd = self._wait_sro_client(username, mbot_hwnd, log)
        if sro_hwnd is None:
            log(f"SRO_Client never appeared for {username}, skipping", "err")
            return

        self._do_login_clicks(sro_hwnd, mbot_hwnd, username, acc, log)

        win32gui.SetWindowPos(sro_hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                              win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        ctypes.windll.user32.ShowWindow(mbot_hwnd, 0)
        ctypes.windll.user32.ShowWindow(sro_hwnd, 0)
        log(f"Login complete for {character}", "ok")
        time.sleep(2)

    def _find_mbot_hwnd(self, username: str, character: str, log: LogFn) -> Optional[int]:
        deadline = time.time() + 30
        while time.time() < deadline:
            try:
                all_mbots = findwindows.find_elements(
                    class_name="#32770",
                    title="mBot v1.12b (vSRO 110)",
                    visible_only=False,
                ) or []
                log(f"[{username}] scanning {len(all_mbots)} mBot window(s) for '{character}'", "info")
                for elem in all_mbots:
                    ctrl = MBotWindow(elem)._find_nth("Character to login", 1)
                    if ctrl and ctrl.name.strip() == character:
                        MBotWindow(elem).start_client()
                        log(f"Start Client sent for {username}", "info")
                        return elem.handle
            except Exception as e:
                log(f"Find mBot error: {e}", "warn")
            time.sleep(1)
        return None

    def _wait_sro_client(self, username: str, mbot_hwnd: int, log: LogFn) -> Optional[int]:
        try:
            import psutil
        except ImportError:
            log("psutil not available", "err")
            return None

        sro_seen  = False
        deadline  = time.time() + 120

        while time.time() < deadline:
            time.sleep(1)
            try:
                _, mbot_pid = win32process.GetWindowThreadProcessId(mbot_hwnd)
                child_pids  = {c.pid for c in psutil.Process(mbot_pid).children(recursive=True)}
            except Exception:
                continue

            if not child_pids and sro_seen:
                log(f"SRO_Client lost for {username}, restarting client", "warn")
                try:
                    elems = findwindows.find_elements(handle=mbot_hwnd)
                    if elems:
                        MBotWindow(elems[0]).start_client()
                except Exception:
                    pass
                sro_seen = False
                continue

            found = self._find_sro_window(child_pids)
            if found:
                log(f"SRO_Client ready for {username}", "info")
                ctypes.windll.user32.ShowWindow(found, 5)
                win32gui.SetWindowPos(found, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                time.sleep(3)
                return found

            if child_pids:
                sro_seen = True

        return None

    @staticmethod
    def _find_sro_window(child_pids: set) -> Optional[int]:
        found = None

        def _enum(hwnd, _):
            nonlocal found
            if found:
                return
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid not in child_pids:
                    return
                if win32gui.GetClassName(hwnd) != "CLIENT":
                    return
                if win32gui.GetWindowText(hwnd) != "SRO_Client":
                    return
                if not win32gui.IsWindowVisible(hwnd):
                    return
                l, t, r, b = win32gui.GetWindowRect(hwnd)
                if (r - l) >= 800 and (b - t) >= 600:
                    found = hwnd
            except Exception:
                pass

        win32gui.EnumWindows(_enum, None)
        return found

    def _do_login_clicks(
        self,
        sro_hwnd: int,
        mbot_hwnd: int,
        username: str,
        acc: dict,
        log: LogFn,
    ) -> None:
        l, t, r, b = win32gui.GetWindowRect(sro_hwnd)
        cx = l + (r - l) // 2
        cy = t + (b - t) // 2

        def _still_alive() -> bool:
            if not win32gui.IsWindow(sro_hwnd):
                log(f"SRO_Client lost during login for {username}, restarting", "warn")
                try:
                    elems = findwindows.find_elements(handle=mbot_hwnd)
                    if elems:
                        MBotWindow(elems[0]).start_client()
                except Exception:
                    pass
                return False
            return True

        if not _still_alive():
            return
        auto.Click(cx, cy)
        time.sleep(0.6)

        if not _still_alive():
            return
        auto.Click(cx, cy - 125)
        time.sleep(0.6)

        if not _still_alive():
            return
        auto.Click(cx - 50, cy + 200)
        time.sleep(0.6)

        if not _still_alive():
            return
        password = base64.b64decode(acc["password"]).decode("utf-8")
        for key in ('{Tab}', username, '{Tab}', password, '{Enter}'):
            auto.SendKeys(key, interval=0.08)
        log(f"Credentials sent for {username}", "ok")
        time.sleep(1)
