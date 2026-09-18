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


def main():
    if getattr(sys, "frozen", False) and len(sys.argv) > 1 and sys.argv[1] == "--ytdlp":
        import yt_dlp
        sys.exit(yt_dlp.main(sys.argv[2:]))

    if getattr(sys, "frozen", False):
        _redirect_output()

    from stream_highlight.server import main as serve
    serve()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
