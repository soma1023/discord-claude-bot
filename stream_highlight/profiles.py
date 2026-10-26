# -*- coding: utf-8 -*-
"""担当チャンネルごとの書き出しの設定（カット感・テロップ量・動画の長さ）。

チャンネルによって編集の好みが違うので、配信のチャンネルから設定を自動で切り替える。
保存先はデータフォルダの channel_profiles.json（手で開いて直せる形）。

設定に無い項目は、それぞれの道具の既定値を使う（画面なら画面の既定値、
realtest.py なら --top 5 / --margin 60）。何も保存していなければ、従来と全く同じ動き。

将来、投稿動画の編集データ（.prproj / XML）から数字を出せるようになったら、
その結果をここに保存する。
"""

import json
import os
import time

from . import paths

FILE_NAME = "channel_profiles.json"
VERSION = 1

# 保存できる項目と、受け付ける範囲（範囲外は端に寄せる）。
# 文字起こしのエンジンだけは文字列なので別扱い。
NUMBER_FIELDS = {
    # 区間の選び方
    "margin_sec": (float, 10.0, 1800.0),      # 見せ場の前後の秒数
    "export_top": (int, 0, 100),              # 書き出す見せ場の数（0 = 表示中の全部）
    # カット
    "gap_sec": (float, 0.5, 5.0),
    "pad_sec": (float, 0.0, 0.5),
    "protect_sec": (float, 0.0, 120.0),
    # テロップ
    "line_chars": (int, 6, 40),
    "max_lines": (int, 1, 4),
    "min_show_sec": (float, 0.2, 5.0),
    "max_show_sec": (float, 1.0, 15.0),
}
BOOL_FIELDS = ("autocut", "subtitles")
TEXT_FIELDS = {"engine": ("qwen", "whisper")}


def _path():
    return os.path.join(paths.data_dir(), FILE_NAME)


def channel_key(info):
    """設定の鍵。ID が取れていればそれを、無ければ表示名を使う。"""
    if getattr(info, "channel_id", ""):
        return "%s:%s" % (info.platform, info.channel_id)
    if info.channel:
        return "%s:name:%s" % (info.platform, info.channel)
    return ""


def clean(settings):
    """保存・適用してよい項目だけを、範囲内に収めて返す。知らない項目は捨てる。"""
    out = {}
    for name, (kind, low, high) in NUMBER_FIELDS.items():
        if settings.get(name) is None:
            continue
        try:
            out[name] = min(high, max(low, kind(settings[name])))
        except (TypeError, ValueError):
            continue
    for name in BOOL_FIELDS:
        if name in settings and settings[name] is not None:
            out[name] = bool(settings[name])
    for name, allowed in TEXT_FIELDS.items():
        if settings.get(name) in allowed:
            out[name] = settings[name]
    if "min_show_sec" in out and "max_show_sec" in out:
        out["max_show_sec"] = max(out["max_show_sec"], out["min_show_sec"])
    return out


def load_all():
    try:
        with open(_path(), encoding="utf-8") as fh:
            data = json.load(fh)
    except (OSError, ValueError):
        return {}
    channels = data.get("channels") if isinstance(data, dict) else None
    return channels if isinstance(channels, dict) else {}


def _save_all(channels):
    path = _path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    temp = path + ".tmp"
    with open(temp, "w", encoding="utf-8") as fh:
        json.dump({"version": VERSION, "channels": channels}, fh, ensure_ascii=False, indent=1)
    os.replace(temp, path)


def find(info):
    """このチャンネルの (鍵, 保存内容)。無ければ (鍵, None)。

    ID で見つからなければ、同じサイトで表示名が同じものを探す（ID を持たない
    古いキャッシュから保存した設定も、新しく取得した配信で使えるように）。
    """
    key = channel_key(info)
    channels = load_all()
    if key in channels:
        return key, channels[key]
    for other_key, entry in channels.items():
        if other_key.startswith(info.platform + ":") and info.channel and \
                entry.get("name") == info.channel:
            return other_key, entry
    return key, None


def settings_for(info):
    """このチャンネルの設定（項目→値）。無ければ空の dict。"""
    _, entry = find(info)
    return clean(entry.get("settings") or {}) if entry else {}


def save(info, settings):
    key, entry = find(info)
    if not key:
        raise ValueError("この配信のチャンネルが分からないため、設定を保存できません。")
    channels = load_all()
    if entry is not None and key != channel_key(info) and channel_key(info) and \
            getattr(info, "channel_id", ""):
        # 表示名で見つけたものは、ID の鍵に移し替える
        channels.pop(key, None)
        key = channel_key(info)
    channels[key] = {"name": info.channel, "platform": info.platform,
                     "settings": clean(settings), "updated_at": time.time()}
    _save_all(channels)
    return key, channels[key]


def remove(info):
    key, entry = find(info)
    if entry is None:
        return False
    channels = load_all()
    channels.pop(key, None)
    _save_all(channels)
    return True


def describe(info):
    """画面に返す形。"""
    key, entry = find(info)
    return {"key": key, "name": info.channel or "", "saved": entry is not None,
            "settings": clean(entry.get("settings") or {}) if entry else {},
            "updated_at": entry.get("updated_at") if entry else None}
