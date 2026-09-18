# -*- coding: utf-8 -*-
"""配信ハイライト抽出ツールの入口。

exe化したときは、ここが2つの役割を兼ねる。

  StreamHighlight.exe            → アプリとして起動する
  StreamHighlight.exe --ytdlp …  → yt-dlp として振る舞う

exe化すると `sys.executable` は自分自身を指すため、`python -m yt_dlp` の
形で外部プロセスを呼べなくなる。そこで自分自身を yt-dlp として
呼び出せるようにしている。
"""

import multiprocessing
import sys
from datetime import datetime


def _redirect_output():
    """コンソールを出さない設定では、標準出力の行き先が無くなる。

    そのまま print すると失敗するうえ、エラーの内容も残らないので、
    保存先フォルダのログファイルに向ける。
    """
    if sys.stdout is not None and sys.stderr is not None:
        return
    import os

    from stream_highlight import paths

    path = os.path.join(paths.data_dir(), "app.log")
    try:
        stream = open(path, "a", encoding="utf-8", errors="replace")
    except OSError:
        stream = open(os.devnull, "w", encoding="utf-8")
    if sys.stdout is None:
        sys.stdout = stream
    if sys.stderr is None:
        sys.stderr = stream


def _hide_console_and_log():
    """コンソールを隠し、起動時の状態をログに残す。

    うまくいかなかったときに何が起きたのかを追えるよう、
    コンソールの有無にかかわらずログファイルに必ず書く。
    """
    import os

    from stream_highlight import code_version, paths

    had_console, hidden = paths.hide_own_console()
    lines = [
        "--- 起動 %s ---" % datetime.now().isoformat(timespec="seconds"),
        "コード      : %s" % code_version(),
        "実行ファイル: %s" % sys.executable,
        "コンソール  : %s" % (
            "隠した" if hidden else
            "あり（他と共有しているため、そのまま）" if had_console else
            "なし"),
    ]
    for line in lines:
        print(line, flush=True)

    # 画面にコンソールが出ている場合、print はそちらに行ってしまう。
    # 後から状況を追えるよう、ログファイルには必ず残す。
    try:
        with open(os.path.join(paths.data_dir(), "app.log"), "a",
                  encoding="utf-8", errors="replace") as fh:
            fh.write("\n".join(lines) + "\n")
    except OSError:
        pass


def main():
    if getattr(sys, "frozen", False) and len(sys.argv) > 1 and sys.argv[1] == "--ytdlp":
        import yt_dlp
        sys.exit(yt_dlp.main(sys.argv[2:]))

    if getattr(sys, "frozen", False):
        _redirect_output()
        _hide_console_and_log()

    from stream_highlight.server import main as serve
    serve()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
