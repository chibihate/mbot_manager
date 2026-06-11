"""
gui_web/server.py — FastAPI web GUI for mBot Manager.

Serves:
  GET  /api/mbots         — list all live mbots (from SQLite)
  GET  /api/logs          — last 200 log entries
  GET  /api/chat/{id}/{ch}— chat content for one mbot+channel
  GET  /api/accounts      — account list (passwords stripped)
  POST /api/command       — enqueue a command
  WS   /ws                — push mbot state updates every 500 ms

Run:
    uvicorn gui_web.server:app --host 0.0.0.0 --port 8765 --reload
"""

import asyncio
import json
import os
from typing import Optional

from fastapi import FastAPI, WebSocket, WebSocketDisconnect, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from core import db

app = FastAPI(title="MBot Manager")

_STATIC_DIR = os.path.join(os.path.dirname(__file__), "static")
app.mount("/static", StaticFiles(directory=_STATIC_DIR), name="static")


# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------
class CommandRequest(BaseModel):
    action:    str
    target_id: Optional[int] = None
    params:    Optional[dict] = None


# ---------------------------------------------------------------------------
# REST endpoints
# ---------------------------------------------------------------------------
@app.get("/")
def index():
    return FileResponse(os.path.join(_STATIC_DIR, "index.html"))


@app.get("/api/mbots")
def get_mbots():
    return db.get_all_mbots()


@app.get("/api/logs")
def get_logs(limit: int = 200):
    return db.get_recent_logs(limit)


@app.get("/api/chat/{mbot_id}/{channel}")
def get_chat(mbot_id: int, channel: str):
    content = db.get_chat(mbot_id, channel)
    if content is None:
        raise HTTPException(status_code=404, detail="Not found")
    return {"content": content}


@app.get("/api/accounts")
def get_accounts():
    accounts_file = "accounts.json"
    if not os.path.exists(accounts_file):
        return []
    with open(accounts_file) as f:
        accounts = json.load(f)
    return [
        {k: v for k, v in acc.items() if k != "password"}
        for acc in accounts
    ]


@app.post("/api/command")
def post_command(req: CommandRequest):
    cmd_id = db.enqueue_command(req.action, req.target_id, req.params)
    return {"id": cmd_id, "status": "pending"}


# ---------------------------------------------------------------------------
# WebSocket — push state every 500 ms
# ---------------------------------------------------------------------------
class _ConnectionManager:
    def __init__(self):
        self._clients: list[WebSocket] = []

    async def connect(self, ws: WebSocket):
        await ws.accept()
        self._clients.append(ws)

    def disconnect(self, ws: WebSocket):
        self._clients.remove(ws)

    async def broadcast(self, data: str):
        dead = []
        for ws in self._clients:
            try:
                await ws.send_text(data)
            except Exception:
                dead.append(ws)
        for ws in dead:
            self._clients.remove(ws)


_manager = _ConnectionManager()


@app.websocket("/ws")
async def websocket_endpoint(ws: WebSocket):
    await _manager.connect(ws)
    try:
        while True:
            await asyncio.sleep(0.5)
    except WebSocketDisconnect:
        _manager.disconnect(ws)


@app.on_event("startup")
async def _start_push_task():
    asyncio.create_task(_push_loop())


async def _push_loop():
    while True:
        await asyncio.sleep(0.5)
        if not _manager._clients:
            continue
        try:
            payload = json.dumps({
                "mbots": db.get_all_mbots(),
                "logs":  db.get_recent_logs(50),
            })
            await _manager.broadcast(payload)
        except Exception:
            pass
