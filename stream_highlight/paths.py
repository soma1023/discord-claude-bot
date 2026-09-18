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


def hide_own_console():
    """自分専用のコンソールが付いていたら隠す。

    exe化の設定でコンソールは出さないようにしているが、環境によっては
    付いてしまうことがある。理由を問わず消せるよう、ここで念のため隠す。

    ただし、コマンドプロンプトから起動した場合など、他から引き継いだ
    コンソールは利用者の画面なので触らない。そのコンソールを使っている
    プロセスが自分だけのときだけ隠す。

    戻り値は (コンソールがあったか, 隠したか)。
    """
    if os.name != "nt":
        return False, False
    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32
        user32 = ctypes.windll.user32

        # 戻り値と引数の型は必ず指定する。既定では32ビット整数として
        # 扱われるため、64ビットのウィンドウハンドルが壊れることがある。
        kernel32.GetConsoleWindow.argtypes = []
        kernel32.GetConsoleWindow.restype = ctypes.c_void_p
        kernel32.GetConsoleProcessList.argtypes = [
            ctypes.POINTER(ctypes.c_uint), ctypes.c_uint]
        kernel32.GetConsoleProcessList.restype = ctypes.c_uint
        user32.ShowWindow.argtypes = [ctypes.c_void_p, ctypes.c_int]
        user32.ShowWindow.restype = ctypes.c_bool

        window = kernel32.GetConsoleWindow()
        if not window:
            return False, False

        # このコンソールを使っているプロセスを数える。
        # 0 は取得失敗なので、その場合も触らない。
        buffer = (ctypes.c_uint * 8)()
        count = kernel32.GetConsoleProcessList(buffer, 8)
        if count != 1:
            return True, False        # 自分だけではないので、そのままにする

        user32.ShowWindow(window, 0)   # 0 = SW_HIDE
        return True, True
    except Exception:      # noqa: BLE001（隠せなくても本体は動かす）
        return False, False
