# -*- coding: utf-8 -*-
"""取得したチャットをディスクに保存しておく。

同じ配信を感度違いで何度も見直すのが主な使い方なので、
2回目以降はダウンロードを丸ごと省略できるようにする。
"""

import gzip
import json
import os
import time

from . import paths
from .sources import ChatMessage, StreamInfo

CACHE_DIR = os.path.join(paths.data_dir(), "cache")
CACHE_VERSION = 1


def _safe(value):
    """ファイル名に使える文字だけ残す（URL由来の値をそのまま繋がない）。"""
    return "".join(c for c in (value or "") if c.isalnum() or c in "-_")


def _path(platform, video_id):
    return os.path.join(CACHE_DIR, "%s_%s.json.gz" % (_safe(platform), _safe(video_id)))


def cleanup_obsolete():
    """使わなくなったファイルを片付ける。

    音声解析をやめた時点で、保存済みの音量ファイル(*.audio.json.gz)は
    読まれなくなった。置いたままでは場所を取るだけなので、起動時に消す。
    戻り値は (消した数, 空いた容量)。
    """
    if not os.path.isdir(CACHE_DIR):
        return 0, 0
    removed = freed = 0
    for name in os.listdir(CACHE_DIR):
        if not name.endswith(".audio.json.gz"):
            continue
        path = os.path.join(CACHE_DIR, name)
        try:
            size = os.path.getsize(path)
            os.remove(path)
        except OSError:
            continue
        removed += 1
        freed += size
    return removed, freed


def _write_gz(path, payload):
    os.makedirs(CACHE_DIR, exist_ok=True)
    tmp = path + ".tmp"
    with gzip.open(tmp, "wt", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False)
    os.replace(tmp, path)
    return path


def _read_gz(path):
    if not os.path.exists(path):
        return None
    try:
        with gzip.open(path, "rt", encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return None


def save(info, messages):
    return _write_gz(_path(info.platform, info.video_id), {
        "version": CACHE_VERSION,
        "saved_at": time.time(),
        "info": info.as_dict(),
        "messages": [m.to_list() for m in messages],
    })


def load(platform, video_id):
    """キャッシュがあれば (StreamInfo, messages) を返す。無ければ None。"""
    path = _path(platform, video_id)
    if not os.path.exists(path):
        return None
    try:
        with gzip.open(path, "rt", encoding="utf-8") as fh:
            payload = json.load(fh)
        if payload.get("version") != CACHE_VERSION:
            return None
        info = StreamInfo(**payload["info"])
        messages = [ChatMessage.from_list(row) for row in payload["messages"]]
        return info, messages
    except (OSError, ValueError, KeyError, TypeError):
        return None


def entries():
    """保存済みの配信を新しい順に返す（履歴表示用）。"""
    if not os.path.isdir(CACHE_DIR):
        return []
    out = []
    for name in os.listdir(CACHE_DIR):
        if not name.endswith(".json.gz") or name.endswith(".audio.json.gz"):
            continue
        path = os.path.join(CACHE_DIR, name)
        try:
            with gzip.open(path, "rt", encoding="utf-8") as fh:
                payload = json.load(fh)
            info = payload["info"]
            out.append({
                "platform": info["platform"],
                "video_id": info["video_id"],
                "url": info["url"],
                "title": info.get("title") or info["video_id"],
                "duration": info.get("duration") or 0,
                "messages": len(payload.get("messages") or []),
                "saved_at": payload.get("saved_at") or os.path.getmtime(path),
            })
        except (OSError, ValueError, KeyError):
            continue
    out.sort(key=lambda e: e["saved_at"], reverse=True)
    return out


def remove(platform, video_id):
    path = _path(platform, video_id)
    if os.path.exists(path):
        os.remove(path)
        return True
    return False
