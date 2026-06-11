"""
core/db.py — SQLite3 persistence layer.

Tables:
  mbots    — live window state (hp/mp/kph/dc)
  logs     — event log entries
  chat     — per-mbot per-channel chat snapshots
  commands — GUI → worker command queue
"""

import json
import sqlite3
import threading
from datetime import datetime
from pathlib import Path
from typing import Optional

DB_PATH = Path("mbot_state.db")

_lock = threading.Lock()


def _conn() -> sqlite3.Connection:
    c = sqlite3.connect(DB_PATH, check_same_thread=False, timeout=10)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA journal_mode=WAL")
    return c


def init_db() -> None:
    with _lock, _conn() as c:
        c.executescript("""
            CREATE TABLE IF NOT EXISTS mbots (
                id          INTEGER PRIMARY KEY,
                window_name TEXT    NOT NULL,
                char        TEXT    NOT NULL,
                is_dc       INTEGER NOT NULL DEFAULT 0,
                hp          REAL    NOT NULL DEFAULT 0,
                mp          REAL    NOT NULL DEFAULT 0,
                kph         TEXT    NOT NULL DEFAULT '-',
                updated_at  TEXT    NOT NULL
            );

            CREATE TABLE IF NOT EXISTS logs (
                id   INTEGER PRIMARY KEY AUTOINCREMENT,
                ts   TEXT    NOT NULL,
                msg  TEXT    NOT NULL,
                kind TEXT    NOT NULL DEFAULT 'info',
                who  TEXT
            );

            CREATE TABLE IF NOT EXISTS chat (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                mbot_id    INTEGER NOT NULL,
                channel    TEXT    NOT NULL,
                content    TEXT    NOT NULL DEFAULT '',
                updated_at TEXT    NOT NULL,
                UNIQUE(mbot_id, channel)
            );

            CREATE TABLE IF NOT EXISTS commands (
                id        INTEGER PRIMARY KEY AUTOINCREMENT,
                ts        TEXT    NOT NULL,
                target_id INTEGER,
                action    TEXT    NOT NULL,
                params    TEXT,
                status    TEXT    NOT NULL DEFAULT 'pending'
            );

            CREATE TABLE IF NOT EXISTS inventory (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                mbot_id    INTEGER NOT NULL,
                inv_type   TEXT    NOT NULL,
                items_json TEXT    NOT NULL DEFAULT '[]',
                updated_at TEXT    NOT NULL,
                UNIQUE(mbot_id, inv_type)
            );

            CREATE TABLE IF NOT EXISTS mbot_log (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                mbot_id    INTEGER NOT NULL,
                content    TEXT    NOT NULL DEFAULT '',
                updated_at TEXT    NOT NULL,
                UNIQUE(mbot_id)
            );

            CREATE INDEX IF NOT EXISTS idx_commands_status ON commands(status);
            CREATE INDEX IF NOT EXISTS idx_logs_ts         ON logs(ts);
        """)


# ---------------------------------------------------------------------------
# mbots
# ---------------------------------------------------------------------------
def upsert_mbot(
    id: int,
    window_name: str,
    char: str,
    is_dc: bool,
    hp: float,
    mp: float,
    kph: str,
) -> None:
    ts = datetime.now().isoformat()
    with _lock, _conn() as c:
        c.execute(
            """
            INSERT INTO mbots (id, window_name, char, is_dc, hp, mp, kph, updated_at)
            VALUES (?,?,?,?,?,?,?,?)
            ON CONFLICT(id) DO UPDATE SET
                window_name = excluded.window_name,
                char        = excluded.char,
                is_dc       = excluded.is_dc,
                hp          = excluded.hp,
                mp          = excluded.mp,
                kph         = excluded.kph,
                updated_at  = excluded.updated_at
            """,
            (id, window_name, char, int(is_dc), hp, mp, kph, ts),
        )


def sync_mbot_ids(active_ids: list[int]) -> None:
    """Remove rows for mbot IDs that are no longer active."""
    with _lock, _conn() as c:
        if not active_ids:
            c.execute("DELETE FROM mbots")
        else:
            placeholders = ",".join("?" * len(active_ids))
            c.execute(f"DELETE FROM mbots WHERE id NOT IN ({placeholders})", active_ids)


def get_all_mbots() -> list[dict]:
    with _conn() as c:
        return [dict(r) for r in c.execute("SELECT * FROM mbots ORDER BY id")]


# ---------------------------------------------------------------------------
# logs
# ---------------------------------------------------------------------------
def add_log(msg: str, kind: str = "info", who: Optional[str] = None) -> None:
    ts = datetime.now().strftime("%H:%M:%S")
    with _lock, _conn() as c:
        c.execute(
            "INSERT INTO logs (ts, msg, kind, who) VALUES (?,?,?,?)",
            (ts, msg, kind, who),
        )


def get_recent_logs(limit: int = 200) -> list[dict]:
    with _conn() as c:
        return [
            dict(r)
            for r in c.execute(
                "SELECT * FROM logs ORDER BY id DESC LIMIT ?", (limit,)
            )
        ]


# ---------------------------------------------------------------------------
# chat
# ---------------------------------------------------------------------------
def upsert_chat(mbot_id: int, channel: str, content: str) -> None:
    ts = datetime.now().isoformat()
    with _lock, _conn() as c:
        c.execute(
            """
            INSERT INTO chat (mbot_id, channel, content, updated_at)
            VALUES (?,?,?,?)
            ON CONFLICT(mbot_id, channel) DO UPDATE SET
                content    = excluded.content,
                updated_at = excluded.updated_at
            """,
            (mbot_id, channel, content, ts),
        )


def upsert_inventory(mbot_id: int, inv_type: str, items: list) -> None:
    ts = datetime.now().isoformat()
    with _lock, _conn() as c:
        c.execute(
            """
            INSERT INTO inventory (mbot_id, inv_type, items_json, updated_at)
            VALUES (?,?,?,?)
            ON CONFLICT(mbot_id, inv_type) DO UPDATE SET
                items_json = excluded.items_json,
                updated_at = excluded.updated_at
            """,
            (mbot_id, inv_type, json.dumps(items, ensure_ascii=False), ts),
        )


def get_inventory(mbot_id: int, inv_type: str) -> list:
    with _conn() as c:
        row = c.execute(
            "SELECT items_json FROM inventory WHERE mbot_id=? AND inv_type=?",
            (mbot_id, inv_type),
        ).fetchone()
        return json.loads(row["items_json"]) if row else []


def upsert_mbot_log(mbot_id: int, content: str) -> None:
    ts = datetime.now().isoformat()
    with _lock, _conn() as c:
        c.execute(
            """
            INSERT INTO mbot_log (mbot_id, content, updated_at)
            VALUES (?,?,?)
            ON CONFLICT(mbot_id) DO UPDATE SET
                content    = excluded.content,
                updated_at = excluded.updated_at
            """,
            (mbot_id, content, ts),
        )


def get_mbot_log(mbot_id: int) -> str:
    with _conn() as c:
        row = c.execute(
            "SELECT content FROM mbot_log WHERE mbot_id=?", (mbot_id,)
        ).fetchone()
        return row["content"] if row else ""


def get_chat(mbot_id: int, channel: str) -> Optional[str]:
    with _conn() as c:
        row = c.execute(
            "SELECT content FROM chat WHERE mbot_id=? AND channel=?",
            (mbot_id, channel),
        ).fetchone()
        return row["content"] if row else None


# ---------------------------------------------------------------------------
# commands
# ---------------------------------------------------------------------------
def enqueue_command(
    action: str,
    target_id: Optional[int] = None,
    params: Optional[dict] = None,
) -> int:
    ts = datetime.now().isoformat()
    with _lock, _conn() as c:
        cur = c.execute(
            "INSERT INTO commands (ts, target_id, action, params, status) VALUES (?,?,?,?,?)",
            (ts, target_id, action, json.dumps(params) if params else None, "pending"),
        )
        return cur.lastrowid


def get_pending_commands() -> list[dict]:
    with _conn() as c:
        return [
            dict(r)
            for r in c.execute(
                "SELECT * FROM commands WHERE status='pending' ORDER BY id"
            )
        ]


def mark_command(cmd_id: int, status: str) -> None:
    with _lock, _conn() as c:
        c.execute("UPDATE commands SET status=? WHERE id=?", (status, cmd_id))
