"""독서토론대회 AI 활용 중계 서버.

- /team     : 참가팀 화면 (논제 설정, 팀 선택, 프롬프트 실행, 결과 전문)
- /audience : 청중 화면 (프롬프트·결과 요약 카드, 실시간 갱신)
- WebSocket(/ws)으로 모든 화면에 상태를 즉시 브로드캐스트하고,
  WS가 끊긴 클라이언트는 /api/state 폴링으로 폴백한다.

실행:  uvicorn main:app --host 0.0.0.0 --port 8000
"""
import asyncio
import json
import uuid
from datetime import datetime
from pathlib import Path
from typing import List

from fastapi import BackgroundTasks, FastAPI, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, RedirectResponse
from pydantic import BaseModel

import ai

app = FastAPI(title="AI Debate Broadcast")
# 외부 AI 페이지(Gemini 등)에서 브라우저가 직접 /api/ingest를 호출할 수 있도록 허용
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)
BASE = Path(__file__).parent

STATE = {"topic": "", "entries": []}
# entry: {id, team('pro'|'con'), prompt, output, summary(None|dict), source, ts}
# source: 'claude'(서버가 직접 생성) 또는 'gemini' 등 외부 AI 이름

TEAM_LABEL = {"pro": "옹호팀", "con": "비판팀"}


class ConnectionManager:
    def __init__(self):
        self.active: List[WebSocket] = []

    async def connect(self, ws: WebSocket):
        await ws.accept()
        self.active.append(ws)
        await ws.send_text(json.dumps(STATE, ensure_ascii=False))

    def disconnect(self, ws: WebSocket):
        if ws in self.active:
            self.active.remove(ws)

    async def broadcast(self):
        data = json.dumps(STATE, ensure_ascii=False)
        for ws in list(self.active):
            try:
                await ws.send_text(data)
            except Exception:
                self.disconnect(ws)


manager = ConnectionManager()


class TopicRequest(BaseModel):
    topic: str


class RunRequest(BaseModel):
    team: str  # 'pro' | 'con'
    prompt: str


class IngestRequest(BaseModel):
    team: str  # 'pro' | 'con'
    prompt: str
    output: str
    source: str = "gemini"  # 외부 AI 이름 (표시용)


@app.get("/")
async def index():
    return RedirectResponse("/team")


@app.get("/team")
async def team_page():
    return FileResponse(BASE / "templates" / "team.html")


@app.get("/audience")
async def audience_page():
    return FileResponse(BASE / "templates" / "audience.html")


@app.get("/api/state")
async def get_state():
    return STATE


@app.post("/api/topic")
async def set_topic(req: TopicRequest):
    STATE["topic"] = req.topic.strip()
    await manager.broadcast()
    return {"status": "success"}


async def _summarize_entry(entry_id: str):
    """백그라운드 요약 생성 후 브로드캐스트 (원안의 Sub-Agent)."""
    entry = next((e for e in STATE["entries"] if e["id"] == entry_id), None)
    if not entry:
        return
    summary = await asyncio.to_thread(
        ai.summarize, TEAM_LABEL[entry["team"]], entry["prompt"], entry["output"]
    )
    entry["summary"] = summary
    await manager.broadcast()


@app.post("/api/run")
async def run_prompt(req: RunRequest, background_tasks: BackgroundTasks):
    if req.team not in TEAM_LABEL or not req.prompt.strip():
        return {"status": "error", "message": "team 또는 prompt가 올바르지 않습니다."}

    output = await asyncio.to_thread(ai.generate_answer, STATE["topic"], req.prompt)
    entry = {
        "id": str(uuid.uuid4()),
        "team": req.team,
        "prompt": req.prompt,
        "output": output,
        "summary": None,
        "source": "claude",
        "ts": datetime.now().isoformat(timespec="seconds"),
    }
    STATE["entries"].append(entry)
    await manager.broadcast()
    background_tasks.add_task(_summarize_entry, entry["id"])
    return {"status": "success", "entry": entry, "mock_mode": ai.MOCK_MODE}


@app.post("/api/ingest")
async def ingest_entry(req: IngestRequest, background_tasks: BackgroundTasks):
    """외부 AI(Gemini 등)에서 생성된 프롬프트·결과 쌍을 받아 중계에 올린다.

    - Gemini 페이지의 자동 중계 유저스크립트가 새 대화를 감지해 호출하거나,
    - 참가팀 화면의 '외부 결과 입력' 폼에서 수동으로 호출한다.
    """
    if req.team not in TEAM_LABEL or not req.prompt.strip() or not req.output.strip():
        return {"status": "error", "message": "team/prompt/output이 올바르지 않습니다."}

    # 같은 내용이 중복 전송되면 무시 (유저스크립트 재감지 대비)
    for e in STATE["entries"]:
        if e["team"] == req.team and e["prompt"] == req.prompt and e["output"] == req.output:
            return {"status": "success", "entry": e, "duplicate": True}

    entry = {
        "id": str(uuid.uuid4()),
        "team": req.team,
        "prompt": req.prompt.strip(),
        "output": req.output.strip(),
        "summary": None,
        "source": req.source.strip() or "gemini",
        "ts": datetime.now().isoformat(timespec="seconds"),
    }
    STATE["entries"].append(entry)
    await manager.broadcast()
    background_tasks.add_task(_summarize_entry, entry["id"])
    return {"status": "success", "entry": entry}


@app.websocket("/ws")
async def websocket_endpoint(ws: WebSocket):
    await manager.connect(ws)
    try:
        while True:
            await ws.receive_text()  # keepalive용 (클라이언트 ping)
    except WebSocketDisconnect:
        manager.disconnect(ws)
