# -*- coding: utf-8 -*-
"""配信ハイライト抽出ツールのWebサーバ。"""

import os
import webbrowser

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from . import autoedit, cache, deps, editexport, localedit, paths, prproj, profiles, retake, transcribe
from . import subtitles as captions      # EditExportRequest.subtitles（ON/OFF）と名前がぶつかるため
from . import autocut as cutter
from . import code_version
from .analyze import Params, analyze, analyze_keyword
from .jobs import manager
from .sources import FetchError

STATIC_DIR = paths.static_dir()

# 起動した時点の版を控えておく。
# code_version() はその場でディスクを読むため、起動後に git pull すると
# 「動いている版」ではなく「ディスク上の版」を答えてしまう。
# それでは再起動の要否が分からないので、ここで確定させる。
RUNNING_VERSION = code_version()

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
    width: int = 0
    height: int = 0
    # 以下は自動カット・テロップ。どちらも False なら従来どおりの同期的な書き出し
    autocut: bool = False
    subtitles: bool = False
    retake: bool = False
    engine: str = transcribe.DEFAULT_ENGINE
    gap_sec: float = cutter.DEFAULT_GAP_SEC
    pad_sec: float = cutter.DEFAULT_PAD_SEC
    protect_sec: float = cutter.DEFAULT_PROTECT_SEC
    # テロップの体裁（チャンネルごとの設定で変わる。既定値は従来どおり）
    line_chars: int = captions.LINE_CHARS
    max_lines: int = captions.MAX_LINES
    min_show_sec: float = captions.MIN_SHOW_SEC
    max_show_sec: float = captions.MAX_SHOW_SEC


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
    """動作中のコードと、ディスク上の版を画面に伝える。

    2つが違う場合は、更新が取り込まれているのに再起動していない状態。
    """
    on_disk = code_version()
    return {"version": RUNNING_VERSION, "on_disk": on_disk,
            "restart_needed": on_disk != RUNNING_VERSION}


@app.post("/api/export/edit")
def export_for_editor(req: EditExportRequest):
    """候補の前後を切り出したシーケンスを、編集ソフト向けに書き出す。"""
    info, _ = _chat_or_404(req.video_key)
    fps = req.fps or info.fps or editexport.DEFAULT_FPS
    width = req.width or info.width or editexport.DEFAULT_WIDTH
    height = req.height or info.height or editexport.DEFAULT_HEIGHT
    if req.autocut or req.subtitles:
        return _start_auto_edit(req, info, fps, width, height)
    try:
        filename, content, segments = editexport.export(
            info, req.peaks, fmt=req.fmt, margin_sec=req.margin_sec,
            fps=fps, media_path=req.media_path.strip(),
            width=width, height=height,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    return {"filename": filename, "content": content, "segments": segments,
            "fps": fps, "width": width, "height": height}


def _start_auto_edit(req, info, fps, width, height):
    """自動カット・テロップ付きの書き出しを、バックグラウンドで始める。

    文字起こしは数分かかるので、ジョブにして進捗を問い合わせてもらう。
    画面は返ってきた job_id を /api/job/{id} で追いかける。
    """
    options = autoedit.EditOptions(
        autocut=req.autocut, subtitles=req.subtitles,
        retake=req.autocut and req.retake, engine=req.engine,
        cut=cutter.CutSettings.from_values(req.gap_sec, req.pad_sec, req.protect_sec),
        caption=captions.CaptionStyle.from_values(req.line_chars, req.max_lines,
                                                  req.min_show_sec, req.max_show_sec))
    media_path = req.media_path.strip().strip('"')
    found = autoedit.problems(options, media_path)
    if found:
        raise HTTPException(status_code=400, detail=" ".join(found))
    if not editexport.build_segments(req.peaks, req.margin_sec, info.duration):
        raise HTTPException(status_code=400, detail="書き出す候補がありません。")

    def work(report):
        return autoedit.run(info, req.peaks, options, fmt=req.fmt, margin_sec=req.margin_sec,
                            fps=fps, media_path=media_path, width=width, height=height,
                            report=report)

    job = manager.submit_task(work, "文字起こしの準備中…")
    return {"job_id": job.id}


@app.get("/api/autoedit/capabilities")
def autoedit_capabilities():
    """自動カット・テロップが使えるか。足りないものは日本語で返す。"""
    caps = transcribe.capabilities()
    caps["retake"] = {"missing": retake.available(), "model": retake.model_name()}
    return caps


REALTEST_SHOW_DONE_SEC = 30 * 60     # 終わった実配信テストを画面に出しておく時間
REALTEST_STALL_SEC = 15 * 60         # これだけ更新が無ければ、止まっているかもしれない


def latest_realtest_progress(now=None):
    """いちばん新しい実配信テスト（realtest.py）の進み具合。見せる物が無ければ None。"""
    import glob
    import json
    import time

    now = now or time.time()
    found = glob.glob(os.path.join(paths.data_dir(), "realtest", "*", "progress.json"))
    if not found:
        return None
    try:
        with open(max(found, key=os.path.getmtime), encoding="utf-8") as fh:
            data = json.load(fh)
    except (OSError, ValueError):
        return None
    age = now - float(data.get("updated_at") or 0)
    if data.get("state") != "running" and age > REALTEST_SHOW_DONE_SEC:
        return None
    data["stalled"] = data.get("state") == "running" and age > REALTEST_STALL_SEC
    return data


class ProfileRequest(BaseModel):
    video_key: str
    settings: dict | None = None


@app.get("/api/profile")
def get_profile(video_key: str):
    """この配信のチャンネルに保存した設定（カット・テロップ・区間）。"""
    info, _ = _chat_or_404(video_key)
    return profiles.describe(info)


@app.post("/api/profile")
def save_profile(req: ProfileRequest):
    """今の設定を、この配信のチャンネルの既定にする。"""
    info, _ = _chat_or_404(req.video_key)
    try:
        profiles.save(info, req.settings or {})
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except OSError as exc:
        raise HTTPException(status_code=500, detail="設定を保存できませんでした: %s" % exc)
    return profiles.describe(info)


@app.post("/api/profile/reset")
def reset_profile(req: ProfileRequest):
    """このチャンネルの設定を消して、既定値に戻す。"""
    info, _ = _chat_or_404(req.video_key)
    profiles.remove(info)
    return profiles.describe(info)


class PanelAnalyzeRequest(BaseModel):
    """Premiere のオートカットパネルから送られる、クリップ1つ分の判定の依頼。"""
    media_path: str
    source_in: float               # 素材の中で使っている範囲（秒）
    source_out: float
    profile: str = localedit.DEFAULT_PROFILE
    # パネルで調整した値。None ならプロファイルのまま
    gap_cut_sec: float | None = None
    pad_sec: float | None = None
    fillers: bool | None = None


def _panel_settings(profile, **overrides):
    # プロファイルは名前で選ばせる（HTTP から任意の yaml のパスを読ませない）
    if profile not in localedit.available_profiles():
        raise HTTPException(status_code=400, detail="プロファイル「%s」はありません。" % profile)
    try:
        flat = localedit.flatten_profile(localedit.load_profile(profile))
        flat.update({k: v for k, v in overrides.items() if v is not None})
        return localedit.build_settings(flat)
    except localedit.EditError as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.get("/api/panel/profiles")
def panel_profiles():
    """パネルのプロファイル選択と、調整欄の初期値。"""
    rows = []
    for name in localedit.available_profiles():
        try:
            settings = _panel_settings(name)
        except HTTPException:
            continue
        rows.append({"name": name, "label": settings.profile,
                     "gap_cut_sec": settings.gap_cut_sec, "pad_sec": settings.pad_sec,
                     "fillers": settings.fillers, "engine": settings.engine,
                     "silence_method": settings.silence_method})
    return {"profiles": rows, "default": localedit.DEFAULT_PROFILE}


@app.post("/api/panel/analyze")
def panel_analyze(req: PanelAnalyzeRequest):
    """クリップの素材と使用範囲から、残す区間（素材の秒）を決める。進捗は /api/job/{id}。"""
    settings = _panel_settings(req.profile, gap_cut_sec=req.gap_cut_sec, pad_sec=req.pad_sec,
                               fillers=req.fillers)
    media_path = req.media_path.strip().strip('"')
    if not os.path.isfile(media_path):
        raise HTTPException(status_code=400, detail="素材ファイルが見つかりません: %s" % media_path)
    if req.source_out - req.source_in < 0.1:
        raise HTTPException(status_code=400, detail="クリップが短すぎます。")
    found = deps.problems(settings.engine, settings.silence_method)
    if found:
        raise HTTPException(status_code=400, detail=" ".join(message for _, message in found))

    def work(report):
        try:
            return localedit.analyze_range(media_path, max(0.0, req.source_in), req.source_out,
                                           settings, report=report)
        except (localedit.EditError, deps.MissingDependency) as exc:
            raise ValueError(str(exc))

    job = manager.submit_task(work, "判定の準備中…")
    return {"job_id": job.id}


class PanelClip(BaseModel):
    media_path: str
    source_in: float
    source_out: float
    seq_start: float               # シーケンスでの開始位置（秒）


class PanelCaptionsRequest(BaseModel):
    """Premiere のパネルから送られる、テロップを作る依頼。"""
    clips: list[PanelClip]
    fps: float
    profile: str = localedit.DEFAULT_PROFILE
    project_path: str = ""         # SRT はこのプロジェクトファイルと同じフォルダに書く
    name: str = "テロップ"


@app.post("/api/panel/captions")
def panel_captions(req: PanelCaptionsRequest):
    """タイムラインのクリップの並びから SRT を作り、プロジェクトの横に保存する。進捗は /api/job/{id}。"""
    settings = _panel_settings(req.profile)
    if not req.clips:
        raise HTTPException(status_code=400, detail="テロップを付けるクリップがありません。")
    if not (1.0 <= req.fps <= 240.0):
        raise HTTPException(status_code=400, detail="シーケンスのフレームレートを読めませんでした。")
    folder = os.path.dirname(req.project_path.strip()) if req.project_path.strip() else ""
    if not folder or not os.path.isdir(folder):
        raise HTTPException(status_code=400, detail="プロジェクトを保存してから使ってください（保存先が分かりません）。")
    safe = "".join("_" if ch in '\\/:*?"<>|' else ch for ch in req.name).strip() or "テロップ"
    clips = [c.model_dump() for c in req.clips]
    for clip in clips:
        clip["media_path"] = clip["media_path"].strip().strip('"')

    def work(report):
        try:
            srt, count = localedit.captions_for_clips(clips, req.fps, settings, report=report)
        except (localedit.EditError, deps.MissingDependency) as exc:
            raise ValueError(str(exc))
        path = localedit.unique_path(os.path.join(folder, safe[:80] + "_テロップ.srt"))
        with open(path, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(srt)
        return {"srt_path": path, "cues": count}

    job = manager.submit_task(work, "テロップの準備中…")
    return {"job_id": job.id}


class PanelCaptionLinesRequest(BaseModel):
    """テロップの行数を、保存済みのプロジェクトファイルから読む依頼。"""
    project_path: str


@app.post("/api/panel/caption_lines")
def panel_caption_lines(req: PanelCaptionLinesRequest):
    """グラフィックのテロップごとの {start, end, lines}（ticks）。Premiere がテキストを渡さないための代わり。"""
    path = req.project_path.strip().strip('"')
    if not path.lower().endswith(".prproj") or not os.path.isfile(path):
        raise HTTPException(status_code=400, detail="プロジェクトを保存してから使ってください（プロジェクトファイルが見つかりません）。")
    try:
        found = prproj.caption_lines(path)
    except (OSError, ValueError) as exc:
        raise HTTPException(status_code=400, detail="プロジェクトファイルを読めませんでした: %s" % exc)
    # lines は画面に見える行数（自動の折り返しを含み、枠からはみ出して見えない行は除く）
    return {"captions": [{"start": c["start"], "end": c["end"], "lines": c["lines"],
                          "wrapped_lines": c["wrapped_lines"], "text": c["text"][:40]} for c in found]}


@app.get("/api/realtest/progress")
def realtest_progress():
    """実配信テストの進捗（画面の進捗バー用）。動いていなければ {"active": false}。"""
    data = latest_realtest_progress()
    return dict(data, active=True) if data else {"active": False}


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

    返すのは /api/capabilities の中身そのまま。繋がらなければ None。
    on_disk の項目が無ければ、相手は「動いている版」を答えられない古い版
    （その場でディスクを読んで答えるので、更新後は嘘の版を返す）。
    呼び出し側はその版表示を信用してはいけない。
    """
    import json
    import urllib.request

    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    try:
        with opener.open("http://%s:%d/api/capabilities" % (host, port), timeout=1.5) as resp:
            found = json.loads(resp.read().decode("utf-8"))
    except Exception:      # noqa: BLE001（繋がらない＝動いていない）
        return None
    return found if isinstance(found, dict) else {}


def _stop_running(host, port, wait_sec=10.0, bind_host=None):
    """動いているサーバに終了を頼み、ポートが空くまで待つ。

    古い版が動いたまま別のポートで立ち上げると、ブックマークや
    開いたままのタブから古い画面を見続けることになる。実際にそれで
    「更新したのに新しい機能が出ない」という取り違えが起きたので、
    古い方を終わらせて同じポートを引き継ぐ。

    空いたかどうかは実際に bind して確かめる。HTTPの返事が無いことを
    根拠にすると、応答しないまま居座っているプロセスを「居ない」と
    見なしてしまい、この後の起動が「ポート使用中」で落ちる。

    止められたら True。
    """
    import time
    import urllib.request

    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    req = urllib.request.Request("http://%s:%d/api/quit" % (host, port), method="POST")
    try:
        opener.open(req, timeout=3.0).read()
    except Exception:      # noqa: BLE001（終了要求が通らない版もある）
        pass

    deadline = time.time() + wait_sec
    while time.time() < deadline:
        time.sleep(0.3)
        if _port_is_free(bind_host or host, port):
            return True
    return False


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

    if running is not None:
        trustworthy = "on_disk" in running
        theirs = running.get("version") or "不明"

        if trustworthy and theirs == mine:
            # 同じ版が動いているなら、二重に立ち上げずそれを開く
            print("すでに起動しています: %s" % url, flush=True)
            if not args.no_browser:
                webbrowser.open(url)
            return

        if trustworthy:
            print("古い版（%s）が動いています。終了させて、更新した版（%s）に"
                  "入れ替えます。" % (theirs, mine), flush=True)
        else:
            # 版を答えられない古い版。表示を信用できないので、中身に関わらず
            # 入れ替える（同じ版だったとしても、入れ替えて困ることはない）。
            print("動いている版を確かめられませんでした。終了させて、"
                  "更新した版（%s）に入れ替えます。" % mine, flush=True)

        if not _stop_running(shown_host, args.port, bind_host=args.host):
            # 終了を頼めなかったので、せめて更新した版を使えるようにする
            print("動いている方を終了できませんでした。別のポートで起動します。"
                  "古い画面のタブは閉じてください。", flush=True)
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

    print("配信ハイライト抽出ツール: %s" % url, flush=True)
    print("動作中のコード: %s" % RUNNING_VERSION, flush=True)
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
