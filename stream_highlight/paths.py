# -*- coding: utf-8 -*-
"""ファイルの置き場所を、通常実行とexe化のどちらでも正しく決める。

PyInstallerで固めると `__file__` は一時展開先を指す。そこに保存したものは
終了時に消えてしまうので、「同梱ファイルを読む場所」と「保存する場所」を
分けて扱う必要がある。
"""

import os
import sys


def is_frozen():
    """exeとして固められた状態で動いているか。"""
    return getattr(sys, "frozen", False)


def bundle_dir():
    """同梱ファイル（画面のHTMLなど）を読む場所。"""
    if is_frozen():
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def static_dir():
    base = bundle_dir()
    for candidate in (os.path.join(base, "static"),
                      os.path.join(base, "stream_highlight", "static")):
        if os.path.isdir(candidate):
            return candidate
    return os.path.join(base, "static")


def data_dir():
    """保存するもの（チャットのキャッシュなど）を置く場所。

    exe化したときは、まず exe と同じ場所の data フォルダを試す。
    フォルダごと持ち運べて、中身を消すのも分かりやすいため。
    書き込めない場所（Program Files など）に置かれていた場合だけ、
    ユーザーのアプリデータ領域に逃がす。
    """
    if not is_frozen():
        return os.path.dirname(os.path.abspath(__file__))

    beside = os.path.join(os.path.dirname(os.path.abspath(sys.executable)), "data")
    try:
        os.makedirs(beside, exist_ok=True)
        probe = os.path.join(beside, ".write-test")
        with open(probe, "w", encoding="utf-8") as fh:
            fh.write("")
        os.remove(probe)
        return beside
    except OSError:
        pass

    home = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    fallback = os.path.join(home, "StreamHighlight")
    os.makedirs(fallback, exist_ok=True)
    return fallback
