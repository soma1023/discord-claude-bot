# -*- coding: utf-8 -*-
"""配信ハイライト抽出ツールのWebサーバ。"""

import os
import webbrowser

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from . import audio as audio_mod
from . import cache
from .analyze import Params, analyze, analyze_keyword
from .jobs import manager
from .sources import FetchError

STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")

app = FastAPI(title="配信ハイライト抽出", docs_url=None, redoc_url=None)


class AnalyzeRequest(BaseModel):
    url: str
    refresh: bool = False
    with_audio: bool = False
    params: dict | None = None


class ReanalyzeRequest(BaseModel):
    video_key: str
    with_audio: bool = True
    params: dict | None = None


class KeywordRequest(BaseModel):
    video_key: str
    keyword: str
    params: dict | None = None


def _chat_or_404(video_key):
    loaded = manager.get_chat(video_key)
    if not loaded:
        raise HTTPException(status_code=404,
                            detail="チャットが見つかりません。もう一度解析してください。")
    return loaded


@app.get("/")
def index():
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))


@app.post("/api/analyze")
def start_analyze(req: AnalyzeRequest):
    try:
        job = manager.submit(req.url.strip(), Params.from_dict(req.params),
                             req.refresh, req.with_audio)
    except FetchError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    return {"job_id": job.id}


@app.get("/api/job/{job_id}")
def job_status(job_id: str):
    job = manager.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="ジョブが見つかりません。")
    return job.as_dict()


@app.post("/api/reanalyze")
def reanalyze(req: ReanalyzeRequest):
    """取得済みのチャットを使って、感度などを変えて解析し直す（ダウンロード無し）。"""
    info, messages, loudness = _chat_or_404(req.video_key)
    result = analyze(messages, info, Params.from_dict(req.params),
                     loudness=loudness if req.with_audio else None)
    result["video_key"] = req.video_key
    return result


@app.post("/api/keyword")
def keyword(req: KeywordRequest):
    word = (req.keyword or "").strip()
    if not word:
        raise HTTPException(status_code=400, detail="検索するワードを入れてください。")
    info, messages, _ = _chat_or_404(req.video_key)
    return analyze_keyword(messages, info, word, Params.from_dict(req.params))


@app.get("/api/capabilities")
def capabilities():
    """音声解析が使える環境かどうかを画面に伝える。"""
    return {"ffmpeg": audio_mod.ffmpeg_available()}


@app.get("/api/history")
def history():
    return {"entries": cache.entries()}


@app.delete("/api/history/{platform}/{video_id}")
def delete_history(platform: str, video_id: str):
    return {"removed": cache.remove(platform, video_id)}


app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


def main():
    import argparse
    import uvicorn

    parser = argparse.ArgumentParser(description="配信ハイライト抽出ツールを起動する")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true", help="ブラウザを自動で開かない")
    args = parser.parse_args()

    url = "http://%s:%d/" % ("127.0.0.1" if args.host == "0.0.0.0" else args.host, args.port)
    print("配信ハイライト抽出ツール: %s" % url)
    if not args.no_browser:
        import threading
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    uvicorn.run(app, host=args.host, port=args.port, log_level="warning")


if __name__ == "__main__":
    main()
