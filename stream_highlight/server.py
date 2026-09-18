# -*- coding: utf-8 -*-
"""配信ハイライト抽出ツールのWebサーバ。"""

import os
import webbrowser

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from . import cache, paths
from . import code_version
from .analyze import Params, analyze, analyze_keyword
from .jobs import manager
from .sources import FetchError

STATIC_DIR = paths.static_dir()

app = FastAPI(title="配信ハイライト抽出", docs_url=None, redoc_url=None)


class AnalyzeRequest(BaseModel):
    url: str
    refresh: bool = False
    params: dict | None = None


class ReanalyzeRequest(BaseModel):
    video_key: str
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
        job = manager.submit(req.url.strip(), Params.from_dict(req.params), req.refresh)
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
    info, messages = _chat_or_404(req.video_key)
    result = analyze(messages, info, Params.from_dict(req.params))
    result["video_key"] = req.video_key
    return result


@app.post("/api/keyword")
def keyword(req: KeywordRequest):
    word = (req.keyword or "").strip()
    if not word:
        raise HTTPException(status_code=400, detail="検索するワードを入れてください。")
    info, messages = _chat_or_404(req.video_key)
    return analyze_keyword(messages, info, word, Params.from_dict(req.params))


@app.post("/api/quit")
def quit_app():
    """画面の「終了」ボタンから、サーバを止める。

    コンソールを出さない設定では、ウィンドウを閉じて終わらせることが
    できないため、終了手段を画面側に用意している。
    """
    server = getattr(app.state, "server", None)
    if server is None:
        return {"stopped": False, "reason": "開発用の起動方法では終了できません。"}
    server.should_exit = True
    return {"stopped": True}


@app.get("/api/capabilities")
def capabilities():
    """動作中のコードを画面に伝える。"""
    return {"version": code_version()}


@app.get("/api/history")
def history():
    return {"entries": cache.entries()}


@app.delete("/api/history/{platform}/{video_id}")
def delete_history(platform: str, video_id: str):
    return {"removed": cache.remove(platform, video_id)}


app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


def _already_running(host, port):
    """同じアプリが既に動いていないか確かめる。

    二重起動すると「ポートが使用中」で落ちるだけで分かりにくいので、
    動いていれば新しく立ち上げず、そのブラウザを開くだけにする。
    """
    import json
    import urllib.request

    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    try:
        with opener.open("http://%s:%d/api/capabilities" % (host, port), timeout=1.5) as resp:
            return "version" in json.loads(resp.read().decode("utf-8"))
    except Exception:      # noqa: BLE001（繋がらない＝動いていない）
        return False


def main():
    import argparse
    import uvicorn

    parser = argparse.ArgumentParser(description="配信ハイライト抽出ツールを起動する")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true", help="ブラウザを自動で開かない")
    args = parser.parse_args()

    shown_host = "127.0.0.1" if args.host == "0.0.0.0" else args.host
    url = "http://%s:%d/" % (shown_host, args.port)

    if _already_running(shown_host, args.port):
        print("すでに起動しています: %s" % url, flush=True)
        if not args.no_browser:
            webbrowser.open(url)
        return

    print("配信ハイライト抽出ツール: %s" % url, flush=True)
    print("動作中のコード: %s" % code_version(), flush=True)
    print("保存先: %s" % paths.data_dir(), flush=True)
    if not args.no_browser:
        import threading
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()

    config = uvicorn.Config(app, host=args.host, port=args.port, log_level="warning")
    server = uvicorn.Server(config)
    app.state.server = server      # 「終了」ボタンから止められるようにする
    server.run()
    print("終了しました。", flush=True)


if __name__ == "__main__":
    main()
