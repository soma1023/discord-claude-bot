# -*- coding: utf-8 -*-
"""配信ハイライト抽出ツールのWebサーバ。"""

import os
import webbrowser

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from . import cache, editexport, paths
from . import code_version
from .analyze import Params, analyze, analyze_keyword
from .jobs import manager
from .sources import FetchError

STATIC_DIR = paths.static_dir()

app = FastAPI(title="配信ハイライト抽出", docs_url=None, redoc_url=None)


@app.middleware("http")
async def no_cache(request, call_next):
    """ブラウザに古い画面を使わせない。

    更新したのに画面が変わらない、という取り違えが何度も起きたため、
    常に取り直させる。手元のサーバなので速度への影響はない。
    """
    response = await call_next(request)
    response.headers["Cache-Control"] = "no-store, must-revalidate"
    return response


class AnalyzeRequest(BaseModel):
    url: str
    refresh: bool = False
    params: dict | None = None


class ReanalyzeRequest(BaseModel):
    video_key: str
    params: dict | None = None


class EditExportRequest(BaseModel):
    video_key: str
    peaks: list[float]
    fmt: str = "xml"
    margin_sec: float = editexport.DEFAULT_MARGIN_SEC
    fps: float = 0.0
    media_path: str = ""


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


@app.post("/api/export/edit")
def export_for_editor(req: EditExportRequest):
    """候補の前後を切り出したシーケンスを、編集ソフト向けに書き出す。"""
    info, _ = _chat_or_404(req.video_key)
    fps = req.fps or info.fps or editexport.DEFAULT_FPS
    try:
        filename, content, segments = editexport.export(
            info, req.peaks, fmt=req.fmt, margin_sec=req.margin_sec,
            fps=fps, media_path=req.media_path.strip(),
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    return {"filename": filename, "content": content,
            "segments": segments, "fps": fps}


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

    動いているものが古い版だと、更新したつもりで古い画面を見続けることに
    なるため、見つけた版を返して呼び出し側で知らせられるようにする。
    繋がらなければ None。
    """
    import json
    import urllib.request

    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    try:
        with opener.open("http://%s:%d/api/capabilities" % (host, port), timeout=1.5) as resp:
            found = json.loads(resp.read().decode("utf-8"))
        return found.get("version") or "不明"
    except Exception:      # noqa: BLE001（繋がらない＝動いていない）
        return None


def _port_is_free(host, port):
    import socket

    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
        probe.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            probe.bind((host, port))
            return True
        except OSError:
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

    mine = code_version()
    running = _already_running(shown_host, args.port)

    if running == mine:
        # 同じ版が動いているなら、二重に立ち上げずそれを開く
        print("すでに起動しています: %s" % url, flush=True)
        if not args.no_browser:
            webbrowser.open(url)
        return

    if running:
        # 古い版が動いたままでも、更新した版を使えるようにする。
        # ここで起動をやめると、更新したのに古い画面を見続けることになる。
        print("別の版が動いています（%s）。更新した版（%s）を別のポートで起動します。"
              % (running, mine), flush=True)
        for candidate in range(args.port + 1, args.port + 21):
            if _already_running(shown_host, candidate) is None and \
                    _port_is_free(args.host, candidate):
                args.port = candidate
                url = "http://%s:%d/" % (shown_host, args.port)
                break
        else:
            print("空いているポートが見つかりませんでした。"
                  "動いている方を終了してから起動し直してください。", flush=True)
            return
        print("古い方は不要なら終了してください"
              "（画面右上の「終了」ボタン、またはタスクマネージャー）。", flush=True)

    print("配信ハイライト抽出ツール: %s" % url, flush=True)
    print("動作中のコード: %s" % code_version(), flush=True)
    print("保存先: %s" % paths.data_dir(), flush=True)

    # 音声解析をやめた時に読まれなくなったファイルを片付ける
    removed, freed = cache.cleanup_obsolete()
    if removed:
        print("使わなくなったファイル %d件（%.1f MB）を削除しました。"
              % (removed, freed / 1024 / 1024), flush=True)
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
