"""
core/worker.py — Standalone background process.

Responsibilities:
  - Scan mBot windows every 5 s
  - Poll HP/MP/KPH every 0.5 s
  - Auto BSObj/NetError/Error dialog dismissal every 60 s
  - Poll chat channels every 2 s
  - Execute commands from the `commands` DB table every 0.5 s
  - Run the update sequence (launch SRO client, wait for controls)

Run standalone:
    python -m core.worker
"""

import json
import os
import subprocess
import threading
import time
from datetime import datetime
from typing import Optional

from core import db
from core.login import LoginRunner, hide_mbots_by_path, start_training_all
from core.window import (
    MBotWindow,
    MbotInfo,
    _parse_char_name,
    count_silkroad_controls,
    dismiss_bsobj_dialogs,
    dismiss_neterror_dialogs,
    dismiss_openerror_dialogs,
    init_win32_modules,
    kill_silkroad_processes,
    scan_mbot_windows,
    CHAT_BUTTON_TEXTS,
)

# ---------------------------------------------------------------------------
# Shared mutable state (protected by _state_lock)
# ---------------------------------------------------------------------------
_state_lock   = threading.Lock()
_live_windows: list[MBotWindow] = []
_live_mbots:   list[MbotInfo]   = []

_stop_event = threading.Event()


def _log(msg: str, kind: str = "info") -> None:
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] [{kind.upper()}] {msg}", flush=True)
    db.add_log(msg, kind)


# ---------------------------------------------------------------------------
# Window scan thread — every 5 s
# ---------------------------------------------------------------------------
def _scan_loop() -> None:
    known_names: list[str] = []
    while not _stop_event.is_set():
        try:
            new_windows = scan_mbot_windows()
            new_names   = sorted(w.mbot.name for w in new_windows if w.mbot.name)
            if new_names != known_names:
                known_names = new_names[:]
                mbots: list[MbotInfo] = []
                for i, w in enumerate(new_windows):
                    char, is_dc = _parse_char_name(w.mbot.name)
                    mbots.append(MbotInfo(id=i + 1, window_name=w.mbot.name, char=char, is_dc=is_dc))

                with _state_lock:
                    global _live_windows, _live_mbots
                    _live_windows = new_windows
                    _live_mbots   = mbots

                # Persist to DB
                active_ids = [m.id for m in mbots]
                db.sync_mbot_ids(active_ids)
                for m in mbots:
                    db.upsert_mbot(m.id, m.window_name, m.char, m.is_dc, m.hp, m.mp, m.kph)

                _log(f"Scan: {len(mbots)} mBot(s) found", "info")
        except Exception as e:
            _log(f"Scan error: {e}", "err")

        _stop_event.wait(5.0)


# ---------------------------------------------------------------------------
# HP/MP/KPH poll thread — every 0.5 s
# ---------------------------------------------------------------------------
def _poll_loop() -> None:
    while not _stop_event.is_set():
        try:
            with _state_lock:
                windows = list(_live_windows)
                mbots   = list(_live_mbots)

            for w, m in zip(windows, mbots):
                try:
                    hp  = w.get_hp()
                    mp  = w.get_mp()
                    kph = w.get_kills_per_hour()
                    if hp  is not None: m.hp  = hp
                    if mp  is not None: m.mp  = mp
                    if kph:             m.kph = kph
                    db.upsert_mbot(m.id, m.window_name, m.char, m.is_dc, m.hp, m.mp, m.kph)
                except Exception as e:
                    _log(f"Poll error [{m.char}]: {e}", "err")
        except Exception as e:
            _log(f"Poll loop error: {e}", "err")

        _stop_event.wait(0.5)


# ---------------------------------------------------------------------------
# Chat poll thread — every 2 s
# ---------------------------------------------------------------------------
def _chat_loop() -> None:
    while not _stop_event.is_set():
        try:
            with _state_lock:
                windows = list(_live_windows)
                mbots   = list(_live_mbots)

            for w, m in zip(windows, mbots):
                for channel in CHAT_BUTTON_TEXTS:
                    try:
                        content = w.get_chat_content(channel)
                        if content is not None:
                            db.upsert_chat(m.id, channel, content)
                    except Exception:
                        pass
        except Exception as e:
            _log(f"Chat poll error: {e}", "err")

        _stop_event.wait(2.0)


# ---------------------------------------------------------------------------
# BSObj / dialog auto-check — every 60 s
# ---------------------------------------------------------------------------
def _bsobj_loop(update_runner: "UpdateRunner") -> None:
    while not _stop_event.is_set():
        _stop_event.wait(60.0)
        if _stop_event.is_set():
            break
        try:
            dismiss_neterror_dialogs(_log)
            dismiss_openerror_dialogs(_log)
            n = dismiss_bsobj_dialogs(_log)
            if n:
                _log(f"[BSObj] {n} dialog(s) dismissed — triggering update", "warn")
                update_runner.run_all()
        except Exception as e:
            _log(f"BSObj loop error: {e}", "err")


# ---------------------------------------------------------------------------
# Update sequence
# ---------------------------------------------------------------------------
_UPDATER_FILE      = "updater.json"
_UPDATE_TIMEOUT_S  = 60
_POLL_INTERVAL_S   = 5
_TARGET_CONTROLS   = 94


def _load_updater_paths() -> list[str]:
    if os.path.exists(_UPDATER_FILE):
        try:
            with open(_UPDATER_FILE) as f:
                return json.load(f)
        except Exception:
            pass
    return []


class UpdateRunner:
    """Runs update sequence in its own daemon thread."""

    def __init__(self):
        self._lock    = threading.Lock()
        self._running = False

    def run_all(self) -> bool:
        paths = _load_updater_paths()
        return self._start(list(range(len(paths))), paths)

    def run_indices(self, indices: list[int]) -> bool:
        paths = _load_updater_paths()
        valid = [i for i in indices if i < len(paths)]
        return self._start(valid, paths)

    def _start(self, indices: list[int], paths: list[str]) -> bool:
        with self._lock:
            if self._running or not indices:
                return False
            self._running = True
        t = threading.Thread(target=self._run, args=(indices, paths), daemon=True)
        t.start()
        return True

    def _run(self, indices: list[int], paths: list[str]) -> None:
        _log(f"[Update] Starting sequence for {len(indices)} client(s)", "accent")
        try:
            for seq, idx in enumerate(indices):
                if _stop_event.is_set():
                    break
                path = paths[idx]
                _log(f"[Update] [{seq + 1}/{len(indices)}] Launching: {path}", "info")
                kill_silkroad_processes()
                try:
                    subprocess.Popen(path, cwd=os.path.dirname(path))
                except Exception as e:
                    _log(f"[Update] Failed to launch {path}: {e}", "err")
                    continue

                elapsed = 0
                while elapsed < _UPDATE_TIMEOUT_S and not _stop_event.is_set():
                    time.sleep(_POLL_INTERVAL_S)
                    elapsed += _POLL_INTERVAL_S
                    count = count_silkroad_controls()
                    if count == 0:
                        _log(f"[Update] {os.path.basename(path)} controls=0, waiting ({elapsed}s)", "info")
                    elif count >= _TARGET_CONTROLS:
                        _log(f"[Update] {os.path.basename(path)} reached {count} controls → done", "ok")
                        kill_silkroad_processes()
                        time.sleep(1)
                        break
                    else:
                        _log(f"[Update] {os.path.basename(path)} controls={count} ({elapsed}s)", "info")
                else:
                    _log(f"[Update] Timeout for {os.path.basename(path)} — killing", "warn")
                    kill_silkroad_processes()
                    time.sleep(1)

                time.sleep(2)

            _log("[Update] All clients processed.", "ok")
        finally:
            with self._lock:
                self._running = False


# ---------------------------------------------------------------------------
# Command executor
# ---------------------------------------------------------------------------
_MBOT_ACTIONS = {
    "start_training":   lambda w: w.start_training(),
    "stop_training":    lambda w: w.stop_training(),
    "start_client":     lambda w: w.start_client(),
    "kill_client":      lambda w: w.kill_client(),
    "kill_mbot":        lambda w: w.kill_mbot(),
    "show_hide_mbot":   lambda w: w.show_hide_mbot(),
    "show_hide_client": lambda w: w.show_hide_client(),
    "log_off":          lambda w: w.log_off(),
    "reset":            lambda w: w.reset_mbot(),
    "get_position":     lambda w: (w.get_current_position(), time.sleep(0.1), w.save_settings()),
    "set_delay":        lambda w: (w.set_delay(), time.sleep(0.1), w.save_settings()),
}

_ACCOUNTS_FILE = "accounts.json"


def _load_accounts() -> list[dict]:
    if os.path.exists(_ACCOUNTS_FILE):
        try:
            with open(_ACCOUNTS_FILE) as f:
                return json.load(f)
        except Exception:
            pass
    return []


def _execute_command(cmd: dict, update_runner: UpdateRunner, login_runner: LoginRunner) -> None:
    action    = cmd["action"]
    target_id = cmd.get("target_id")
    params    = json.loads(cmd["params"]) if cmd.get("params") else {}

    if action == "run_update":
        indices = params.get("indices")
        if indices is None:
            update_runner.run_all()
        else:
            update_runner.run_indices(indices)
        return

    if action == "login":
        accounts = _load_accounts()
        indices  = params.get("indices") or list(range(len(accounts)))
        login_runner.run(accounts, indices, _log)
        return

    if action == "hide_mbots":
        accounts = _load_accounts()
        indices  = params.get("indices") or list(range(len(accounts)))
        hide_mbots_by_path(accounts, indices, _log)
        return

    if action == "start_training_all":
        start_training_all(_log)
        return

    if action == "get_inventory":
        inv_type = params.get("inv_type", "Inventory")
        from core.window import INVENTORY_OPTIONS
        inv_idx = INVENTORY_OPTIONS.index(inv_type) if inv_type in INVENTORY_OPTIONS else 3
        with _state_lock:
            windows = list(_live_windows)
            mbots   = list(_live_mbots)
        for w, m in zip(windows, mbots):
            if target_id is not None and m.id != target_id:
                continue
            try:
                w.set_inventory_combo(inv_idx)
                w.refresh_inventory()
                items = w.get_inventory_items()
                db.upsert_inventory(m.id, inv_type, items)
            except Exception as e:
                _log(f"get_inventory error [{m.char}]: {e}", "err")
        return

    if action == "get_mbot_log":
        with _state_lock:
            windows = list(_live_windows)
            mbots   = list(_live_mbots)
        for w, m in zip(windows, mbots):
            if target_id is not None and m.id != target_id:
                continue
            try:
                content = w.get_log() or ""
                db.upsert_mbot_log(m.id, content)
            except Exception as e:
                _log(f"get_mbot_log error [{m.char}]: {e}", "err")
        return

    if action not in _MBOT_ACTIONS:
        _log(f"[CMD] Unknown action: {action}", "warn")
        return

    fn = _MBOT_ACTIONS[action]
    with _state_lock:
        windows = list(_live_windows)
        mbots   = list(_live_mbots)

    targets = [
        (w, m) for w, m in zip(windows, mbots)
        if target_id is None or m.id == target_id
    ]
    if not targets:
        _log(f"[CMD] No target for action={action} target_id={target_id}", "warn")
        return

    for w, m in targets:
        try:
            fn(w)
            _log(f"[CMD] {action} → {m.char}", "ok")
        except Exception as e:
            _log(f"[CMD] {action} error [{m.char}]: {e}", "err")


def _command_loop(update_runner: UpdateRunner, login_runner: LoginRunner) -> None:
    while not _stop_event.is_set():
        try:
            for cmd in db.get_pending_commands():
                db.mark_command(cmd["id"], "running")
                try:
                    _execute_command(cmd, update_runner, login_runner)
                    db.mark_command(cmd["id"], "done")
                except Exception as e:
                    _log(f"[CMD] Unhandled error cmd_id={cmd['id']}: {e}", "err")
                    db.mark_command(cmd["id"], "error")
        except Exception as e:
            _log(f"Command loop error: {e}", "err")

        _stop_event.wait(0.5)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def run() -> None:
    db.init_db()
    _log("Core worker starting", "ok")

    if not init_win32_modules():
        _log("win32/pywinauto not available — running in stub mode", "warn")

    update_runner = UpdateRunner()
    login_runner  = LoginRunner()

    threads = [
        threading.Thread(target=_scan_loop,                                        daemon=True, name="scan"),
        threading.Thread(target=_poll_loop,                                        daemon=True, name="poll"),
        threading.Thread(target=_chat_loop,                                        daemon=True, name="chat"),
        threading.Thread(target=_bsobj_loop,   args=(update_runner,),              daemon=True, name="bsobj"),
        threading.Thread(target=_command_loop, args=(update_runner, login_runner), daemon=True, name="cmd"),
    ]
    for t in threads:
        t.start()

    _log("All worker threads started. Press Ctrl+C to stop.", "ok")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        _log("Shutting down…", "warn")
        _stop_event.set()
        for t in threads:
            t.join(timeout=3)
        _log("Worker stopped.", "ok")


if __name__ == "__main__":
    run()
