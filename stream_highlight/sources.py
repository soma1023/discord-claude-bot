# -*- coding: utf-8 -*-
"""配信URLの解析と、チャットリプレイの取得。

YouTube : yt-dlp の live_chat 字幕としてまとめて取得する（動画本体はDLしない）。
Twitch  : 公式のVODチャットAPIが廃止されているため、Web版が使っているGQL
          エンドポイント（TwitchDownloader等と同じ経路）を叩いて取得する。
"""

import glob
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, asdict

from . import paths


# Windowsでコンソールを出さない設定のアプリから子プロセスを起動すると、
# 既定では子プロセス用の黒い画面が開いてしまう。それを抑える。
# Windows以外にこの定数は無いので、その場合は0（指定なし）になる。
_NO_CONSOLE = getattr(subprocess, "CREATE_NO_WINDOW", 0)


class FetchError(RuntimeError):
    """取得に失敗したときに投げる。メッセージはそのままUIに出す。"""


@dataclass
class ChatMessage:
    offset: float          # 配信開始からの秒数
    author: str
    text: str
    kind: str = "text"     # text / paid / sticker / member / gift
    amount: str = ""       # スパチャ金額の表示文字列

    def to_list(self):
        return [round(self.offset, 3), self.author, self.text, self.kind, self.amount]

    @staticmethod
    def from_list(row):
        return ChatMessage(row[0], row[1], row[2], row[3], row[4])


@dataclass
class StreamInfo:
    platform: str
    video_id: str
    url: str
    title: str = ""
    channel: str = ""
    duration: float = 0.0
    thumbnail: str = ""

    def as_dict(self):
        return asdict(self)

    def time_url(self, seconds):
        """指定秒数から再生を始めるURLを作る。"""
        s = max(0, int(seconds))
        if self.platform == "youtube":
            return "https://www.youtube.com/watch?v=%s&t=%ds" % (self.video_id, s)
        h, rem = divmod(s, 3600)
        m, sec = divmod(rem, 60)
        return "https://www.twitch.tv/videos/%s?t=%dh%dm%ds" % (self.video_id, h, m, sec)


# ---------------------------------------------------------------- URL解析

_YOUTUBE_PATTERNS = [
    r"(?:youtube\.com|youtube-nocookie\.com)/watch\?(?:.*&)?v=([\w-]{11})",
    r"(?:youtube\.com|youtube-nocookie\.com)/(?:live|embed|shorts|v)/([\w-]{11})",
    r"youtu\.be/([\w-]{11})",
]
_TWITCH_PATTERNS = [
    r"twitch\.tv/videos/(\d+)",
    r"twitch\.tv/[\w-]+/(?:video|v)/(\d+)",
]


def parse_url(url):
    """URLから (platform, video_id) を取り出す。対応外なら FetchError。"""
    raw = (url or "").strip()
    if not raw:
        raise FetchError("URLが空です。")
    for pat in _YOUTUBE_PATTERNS:
        m = re.search(pat, raw)
        if m:
            return "youtube", m.group(1)
    for pat in _TWITCH_PATTERNS:
        m = re.search(pat, raw)
        if m:
            return "twitch", m.group(1)
    # URLではなくIDだけ貼られた場合
    if re.fullmatch(r"[\w-]{11}", raw):
        return "youtube", raw
    if re.fullmatch(r"\d{8,}", raw):
        return "twitch", raw
    raise FetchError(
        "YouTubeかTwitchのアーカイブURLを入れてください。\n"
        "例: https://www.youtube.com/watch?v=xxxxxxxxxxx / https://www.twitch.tv/videos/123456789"
    )


# ---------------------------------------------------------------- yt-dlp

def _ytdlp_command():
    """yt-dlp の起動コマンドを返す。

    exe化すると sys.executable は自分自身を指すため、`-m yt_dlp` では
    アプリが再起動してしまう。そこで「--ytdlp を付けて自分を呼ぶと
    yt-dlp として振る舞う」入口を用意し、それを使う。
    """
    if paths.is_frozen():
        return [sys.executable, "--ytdlp"]
    try:
        import yt_dlp  # noqa: F401
        return [sys.executable, "-m", "yt_dlp"]
    except ImportError:
        pass
    exe = shutil.which("yt-dlp") or shutil.which("yt-dlp.exe")
    if exe:
        return [exe]
    raise FetchError("yt-dlp が見つかりません。`pip install -r stream_highlight/requirements.txt` を実行してください。")


def _run_ytdlp(args, timeout=1800, on_line=None):
    cmd = _ytdlp_command() + ["--no-warnings", "--no-playlist", "--newline"] + args
    proc = subprocess.Popen(
        cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, encoding="utf-8", errors="replace",
        creationflags=_NO_CONSOLE,
    )
    lines = []
    deadline = time.time() + timeout
    for line in proc.stdout:
        lines.append(line.rstrip())
        if on_line:
            on_line(line.rstrip())
        if time.time() > deadline:
            proc.kill()
            raise FetchError("yt-dlp がタイムアウトしました。")
    proc.wait()
    return proc.returncode, lines


def _ytdlp_error(lines):
    errors = [l for l in lines if "ERROR" in l]
    detail = errors[-1] if errors else (lines[-1] if lines else "")
    if "Private video" in detail or "members-only" in detail:
        return "限定公開・メンバー限定の動画には対応していません。"
    if "Video unavailable" in detail:
        return "動画が見つかりませんでした。URLを確認してください。"
    return "yt-dlp の実行に失敗しました: %s" % (detail or "原因不明")


def fetch_info(url):
    """動画のタイトル・長さなどのメタ情報を取得する。"""
    code, lines = _run_ytdlp(["--dump-single-json", "--skip-download", url], timeout=180)
    payload = None
    for line in lines:
        if line.startswith("{"):
            try:
                payload = json.loads(line)
            except json.JSONDecodeError:
                continue
    if code != 0 or payload is None:
        raise FetchError(_ytdlp_error(lines))
    platform, video_id = parse_url(url)
    return StreamInfo(
        platform=platform,
        video_id=video_id,
        url=url,
        title=payload.get("title") or "",
        channel=payload.get("uploader") or payload.get("channel") or "",
        duration=float(payload.get("duration") or 0),
        thumbnail=payload.get("thumbnail") or "",
    )


# ---------------------------------------------------------------- YouTube

_RENDERER_KINDS = {
    "liveChatTextMessageRenderer": "text",
    "liveChatPaidMessageRenderer": "paid",
    "liveChatPaidStickerRenderer": "sticker",
    "liveChatMembershipItemRenderer": "member",
    "liveChatSponsorshipsGiftPurchaseAnnouncementRenderer": "gift",
}


def _runs_to_text(message):
    """message.runs を文字列にする。絵文字は :shortcut: 形式で残す。

    実データでは想定した入れ子が null で来ることがあるので、
    期待した型でなければ黙って読み飛ばす。
    """
    if not isinstance(message, dict):
        return ""
    runs = message.get("runs")
    if not isinstance(runs, list):
        return ""
    out = []
    for run in runs:
        if not isinstance(run, dict):
            continue
        text = run.get("text")
        if isinstance(text, str):
            out.append(text)
            continue
        emoji = run.get("emoji")
        if not isinstance(emoji, dict):
            continue
        shortcuts = emoji.get("shortcuts")
        if isinstance(shortcuts, list) and shortcuts:
            out.append(str(shortcuts[0]))
        elif emoji.get("isCustomEmoji"):
            out.append(":%s:" % (emoji.get("emojiId") or "emoji"))
        else:
            out.append(str(emoji.get("emojiId") or ""))
    return "".join(out)


def _simple_text(node):
    if not isinstance(node, dict):
        return ""
    if "simpleText" in node:
        return node["simpleText"]
    return _runs_to_text(node)


def _offset_seconds(renderer, offset_ms, fallback_origin):
    """配信開始からの秒数を求める。取れなければ None。"""
    if offset_ms is not None:
        try:
            return int(offset_ms) / 1000.0
        except (TypeError, ValueError):
            pass
    usec = renderer.get("timestampUsec")
    if usec is None or fallback_origin is None:
        return None
    try:
        return (int(usec) - fallback_origin) / 1_000_000.0
    except (TypeError, ValueError):
        return None


def _parse_live_chat_line(line, fallback_origin):
    """live_chat.json の1行から ChatMessage を作る。対象外の行は None。

    実際のチャットログには、通常の発言のほかに広告・アンケート・削除通知・
    プレースホルダなど多様な要素が混ざる。想定した入れ子が欠けていたり
    null だったりするのは普通なので、形が違えば読み飛ばす。
    """
    try:
        obj = json.loads(line)
    except json.JSONDecodeError:
        return None
    if not isinstance(obj, dict):
        return None

    replay = obj.get("replayChatItemAction")
    if isinstance(replay, dict):
        actions = replay.get("actions")
        offset_ms = replay.get("videoOffsetTimeMsec")
    else:
        actions = obj.get("actions")
        offset_ms = None
    if offset_ms is None:
        offset_ms = obj.get("videoOffsetTimeMsec")
    if not isinstance(actions, list):
        return None

    for action in actions:
        if not isinstance(action, dict):
            continue
        add = action.get("addChatItemAction")
        item = add.get("item") if isinstance(add, dict) else None
        if not isinstance(item, dict):
            continue
        for key, kind in _RENDERER_KINDS.items():
            renderer = item.get(key)
            if not isinstance(renderer, dict):
                continue
            offset = _offset_seconds(renderer, offset_ms, fallback_origin)
            if offset is None:
                return None
            text = _runs_to_text(renderer.get("message"))
            if not text and kind in ("sticker", "member", "gift"):
                text = (_simple_text(renderer.get("headerSubtext"))
                        or _simple_text(renderer.get("headerPrimaryText")))
            author = _simple_text(renderer.get("authorName"))
            if not author:
                # ギフト告知などは、名前が header の中に入っている
                header = renderer.get("header")
                if isinstance(header, dict):
                    for sub in header.values():
                        if isinstance(sub, dict):
                            author = _simple_text(sub.get("authorName"))
                            if author:
                                break
            return ChatMessage(
                offset=offset,
                author=author,
                text=text,
                kind=kind,
                amount=_simple_text(renderer.get("purchaseAmountText")),
            )
    return None


def _first_timestamp_usec(path):
    """videoOffsetTimeMsec が無い形式のために、最初のタイムスタンプを拾う。"""
    with open(path, "r", encoding="utf-8", errors="replace") as fh:
        for line in fh:
            m = re.search(r'"timestampUsec"\s*:\s*"(\d+)"', line)
            if m:
                return int(m.group(1))
    return None


def _tail_line_count(state):
    """ダウンロード中のファイルを追いかけて、取得済み行数を数える。"""
    path = state.get("path")
    if not path:
        matches = sorted(glob.glob(state["pattern"]), key=os.path.getmtime)
        if not matches:
            return state["count"]
        path = state["path"] = matches[-1]
    try:
        with open(path, "rb") as fh:
            fh.seek(state["pos"])
            while True:
                chunk = fh.read(1 << 20)
                if not chunk:
                    break
                state["count"] += chunk.count(b"\n")
                state["pos"] += len(chunk)
    except OSError:
        pass
    return state["count"]


def fetch_youtube_chat(url, progress=None):
    """YouTubeアーカイブのチャットリプレイを取得する。"""
    tmpdir = tempfile.mkdtemp(prefix="stream_highlight_")
    try:
        outtmpl = os.path.join(tmpdir, "chat.%(ext)s")
        state = {"pattern": os.path.join(tmpdir, "*live_chat.json*"), "path": None, "pos": 0, "count": 0}
        stop = threading.Event()

        def watch():
            while not stop.wait(0.7):
                n = _tail_line_count(state)
                if progress and n:
                    progress("チャットを取得中… %s件" % f"{n:,}", None)

        watcher = threading.Thread(target=watch, daemon=True)
        watcher.start()
        try:
            code, lines = _run_ytdlp([
                "--skip-download",
                "--write-subs", "--sub-langs", "live_chat",
                "-o", outtmpl,
                url,
            ])
        finally:
            stop.set()
            watcher.join(timeout=2)

        matches = sorted(glob.glob(os.path.join(tmpdir, "*live_chat.json")), key=os.path.getsize)
        if not matches:
            if code != 0:
                raise FetchError(_ytdlp_error(lines))
            raise FetchError(
                "この動画にはチャットリプレイがありません。\n"
                "配信者がチャットを無効化しているか、通常の投稿動画の可能性があります。"
            )
        path = matches[-1]
        origin = None
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            head = fh.readline()
        if head and '"videoOffsetTimeMsec"' not in head:
            origin = _first_timestamp_usec(path)

        messages = []
        broken = 0
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            for line in fh:
                if not line.strip():
                    continue
                try:
                    msg = _parse_live_chat_line(line, origin)
                except Exception:      # noqa: BLE001（1行の崩れで全体を落とさない）
                    broken += 1
                    continue
                if msg is not None:
                    messages.append(msg)

        if not messages:
            raise FetchError(
                "チャットは取得できましたが、コメントを1件も読み取れませんでした。\n"
                "YouTube側の形式が変わった可能性があります（読み取れなかった行: %d）。" % broken
            )
        messages.sort(key=lambda m: m.offset)
        return messages
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


# ---------------------------------------------------------------- Twitch

_GQL_URL = "https://gql.twitch.tv/gql"

# Twitchは Client-ID によって整合性チェックの扱いが変わる。
# 1つ目は yt-dlp が現在使っている値、2つ目は旧Web版の値。
# 旧IDは弾かれることがあるので、実際に通るものを実行時に選ぶ。
_GQL_CLIENT_IDS = (
    "ue6666qo983tsx6so1t0vnawi233wa",
    "kimne78kx3ncx6brgo4mv6wki5h1ko",
)
_GQL_HASH = "b70a3591ff0f4e0313d126c6a1502d79a1c02baebb288227c582044aa76adf6a"
_NULL_RETRY_WAIT = 1.5      # 空応答の再試行間隔（テストでは0にする）

# 失敗したときに生の応答を残す場所。原因究明の手がかりになる。
DEBUG_PATH = os.path.join(paths.data_dir(), "twitch_debug.json")


def _gql_post(payload, retries=4, client_id=None):
    body = json.dumps(payload).encode("utf-8")
    last = None
    for attempt in range(retries):
        req = urllib.request.Request(
            _GQL_URL, data=body,
            headers={
                "Client-ID": client_id or _GQL_CLIENT_IDS[0],
                # Twitch自身のクライアントと同じ指定にする。
                # application/json にすると扱いが変わることがある。
                "Content-Type": "text/plain;charset=UTF-8",
                "Accept": "*/*",
                "Origin": "https://www.twitch.tv",
                "Referer": "https://www.twitch.tv/",
                "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                               "AppleWebKit/537.36 (KHTML, like Gecko) "
                               "Chrome/126.0.0.0 Safari/537.36"),
            },
        )
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            last = exc
            if exc.code not in (429, 500, 502, 503, 504):
                raise FetchError("Twitchのチャット取得に失敗しました (HTTP %s)。" % exc.code)
        except (urllib.error.URLError, TimeoutError, json.JSONDecodeError) as exc:
            last = exc
        time.sleep(2 ** attempt)
    raise FetchError("Twitchのチャット取得に失敗しました: %s" % last)


def _comment_request(video_id, offset=None, cursor=None):
    variables = {"videoID": str(video_id)}
    if cursor:
        variables["cursor"] = cursor
    else:
        variables["contentOffsetSeconds"] = int(offset or 0)
    return {
        "operationName": "VideoCommentsByOffsetOrCursor",
        "variables": variables,
        "extensions": {"persistedQuery": {"version": 1, "sha256Hash": _GQL_HASH}},
    }


# Twitchのシステム通知。サブスク・ギフト・視聴連続記録・レイドなど。
# 本人が書いたコメントではないので、盛り上がりの判定からは外す。
_TWITCH_NOTICE = re.compile(
    r"subscribed with Prime"
    r"|subscribed at Tier \d"
    r"|(?:T|t)hey've subscribed for \d+ month"
    r"|is continuing the Gift Sub"
    r"|gifted a Tier \d+ Sub"
    r"|is gifting \d+ Tier \d+ Sub"
    r"|watched \d+ consecutive stream"
    r"|sparked a watch streak"
    r"|\d+ raiders? from"
    r"|converted from a Prime Sub"
)


def _split_twitch_notice(text):
    """システム通知と、本人が書いたコメントを切り分ける。

    通知は「○○ subscribed with Prime. They've subscribed for 43 months!」の形で、
    そのあとに本人のひとことが続くことがある（例: 末尾の「ajak0nHi」）。
    その部分は本物のコメントなので残す。

    戻り値は (本人のコメント, 通知だったか)。
    """
    found = _TWITCH_NOTICE.search(text)
    if not found:
        return text, False
    tail = text[found.end():]
    mark = tail.find("!")
    remainder = tail[mark + 1:].strip() if mark >= 0 else ""
    return remainder, True


def _node_to_message(node):
    message = node.get("message") or {}
    fragments = message.get("fragments") or []
    text = "".join(f.get("text") or "" for f in fragments)
    commenter = node.get("commenter") or {}
    body, is_notice = _split_twitch_notice(text)
    return ChatMessage(
        offset=float(node.get("contentOffsetSeconds") or 0),
        author=commenter.get("displayName") or commenter.get("login") or "",
        text=body if is_notice else text,
        # 通知に添えられた本人のひとことは普通のコメントとして扱う
        kind="system" if (is_notice and not body) else "text",
    )


def _extract_comments(payload):
    """GQL応答から comments を取り出す。取れない場合は理由を返す。

    Twitchは HTTP 200 のまま data.video.comments を null で返すことがある
    （混雑・レート制限・サブスク限定VODなど）ので、例外に頼らず形を確かめる。
    """
    if not isinstance(payload, list) or not payload or not isinstance(payload[0], dict):
        return None, "Twitchの応答が想定と違う形式でした"
    first = payload[0]

    errors = first.get("errors")
    if isinstance(errors, list) and errors:
        messages = [e.get("message") for e in errors if isinstance(e, dict)]
        detail = "; ".join(m for m in messages if m)
        return None, "Twitchがエラーを返しました: %s" % (detail or "詳細不明")

    data = first.get("data")
    if not isinstance(data, dict):
        return None, "Twitchの応答にデータが含まれていませんでした"
    video = data.get("video")
    if video is None:
        return None, ("VODが見つかりませんでした。"
                      "削除済み・非公開・サブスク限定の可能性があります。")
    if not isinstance(video, dict):
        return None, "Twitchの応答が想定と違う形式でした"

    comments = video.get("comments")
    if not isinstance(comments, dict):
        return None, ("Twitchがチャットを返しませんでした。"
                      "混雑・アクセス制限のほか、チャットが残っていない可能性があります。")
    return comments, None


def _fetch_twitch_segment(video_id, start, end, seen, lock, collected, on_advance,
                          failures=None, client_id=None):
    """[start, end) 区間のコメントをページングで集める。

    取得できなかった区間があっても、他の区間で取れた分は捨てない。
    全区間が空だったときだけ、記録しておいた理由を使って失敗させる。
    """
    cursor = None
    offset = start
    pages = 0
    while pages < 5000:
        comments = None
        reason = None
        for attempt in range(3):
            payload = _gql_post([_comment_request(video_id, offset=offset, cursor=cursor)],
                                client_id=client_id)
            comments, reason = _extract_comments(payload)
            if comments is not None:
                break
            if attempt < 2:
                time.sleep(_NULL_RETRY_WAIT * (attempt + 1))   # 一時的な混雑なら待てば通る
        if comments is None:
            if failures is not None:
                failures.append(reason or "チャットを取得できませんでした")
            return
        pages += 1
        edges = comments.get("edges") or []
        if not edges:
            return
        last_offset = offset
        for edge in edges:
            node = edge.get("node") or {}
            node_id = node.get("id")
            msg = _node_to_message(node)
            last_offset = msg.offset
            if end is not None and msg.offset >= end:
                on_advance(msg.offset - start)
                return
            with lock:
                if node_id and node_id in seen:
                    continue
                if node_id:
                    seen.add(node_id)
                collected.append(msg)
        on_advance(max(0.0, last_offset - start))
        if not (comments.get("pageInfo") or {}).get("hasNextPage"):
            return
        cursor = edges[-1].get("cursor")
        if not cursor:
            return
        offset = None


def _save_debug(payload):
    """失敗したときの生の応答を残す。次の調査の手がかりにする。"""
    try:
        with open(DEBUG_PATH, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False, indent=2)
        return DEBUG_PATH
    except OSError:
        return None


def probe_twitch(video_id):
    """使える Client-ID を1回のリクエストで見つける。

    いきなり並列で取りに行くと、弾かれたのか混んでいるのか分からなくなる。
    先に単発で確かめてから本番の取得に入る。
    """
    reasons = []
    last_payload = None
    for client_id in _GQL_CLIENT_IDS:
        payload = _gql_post([_comment_request(video_id, offset=0)], client_id=client_id)
        last_payload = payload
        comments, reason = _extract_comments(payload)
        if comments is not None:
            return client_id
        reasons.append("Client-ID %s… → %s" % (client_id[:10], reason))

    saved = _save_debug(last_payload)
    detail = "\n".join(reasons)
    if saved:
        detail += "\n\n応答の内容を %s に保存しました。" % saved
    raise FetchError("Twitchからチャットを取得できませんでした。\n" + detail)


def fetch_twitch_chat(video_id, duration=0, progress=None, workers=3):
    """TwitchのVODコメントを取得する。長い配信は区間分割して並列に取る。"""
    seen = set()
    collected = []
    lock = threading.Lock()
    duration = float(duration or 0)

    if progress:
        progress("Twitchへの接続を確認中…", 0.01)
    client_id = probe_twitch(video_id)

    done = {"sec": 0.0}

    def make_advance(seg_len):
        state = {"prev": 0.0}

        def on_advance(progressed):
            delta = min(progressed, seg_len) - state["prev"]
            if delta <= 0:
                return
            state["prev"] += delta
            with lock:
                done["sec"] += delta
                if progress and duration > 0:
                    progress(
                        "チャットを取得中… %s件" % f"{len(collected):,}",
                        min(0.99, done["sec"] / duration),
                    )
        return on_advance

    # 並列に取りに行くと速いが、増やしすぎるとTwitch側に弾かれる。
    # 30分以上のVODを、控えめな数で分担する。
    failures = []
    if duration > 1800 and workers > 1:
        segments = max(1, min(workers, int(duration // 900)))
    else:
        segments = 1

    if segments == 1:
        _fetch_twitch_segment(video_id, 0, None, seen, lock, collected,
                              make_advance(duration or 1e9), failures, client_id)
    else:
        seg_len = duration / segments
        bounds = [(i * seg_len, (i + 1) * seg_len if i < segments - 1 else None)
                  for i in range(segments)]
        with ThreadPoolExecutor(max_workers=segments) as pool:
            futures = [
                pool.submit(_fetch_twitch_segment, video_id, start, end,
                            seen, lock, collected, make_advance(seg_len),
                            failures, client_id)
                for start, end in bounds
            ]
            for fut in futures:
                try:
                    fut.result()
                except FetchError as exc:
                    failures.append(str(exc))

    if not collected:
        raise FetchError(failures[0] if failures
                         else "このVODからコメントを取得できませんでした。")

    collected.sort(key=lambda m: m.offset)
    return collected


# ---------------------------------------------------------------- 共通入口

def fetch_chat(info, progress=None):
    """プラットフォームに応じてチャットを取得する。"""
    if progress:
        progress("配信情報を確認中…", 0.02)
    if info.platform == "youtube":
        return fetch_youtube_chat(info.url, progress=progress)
    return fetch_twitch_chat(info.video_id, duration=info.duration, progress=progress)
