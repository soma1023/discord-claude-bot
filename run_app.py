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


def main():
    if getattr(sys, "frozen", False) and len(sys.argv) > 1 and sys.argv[1] == "--ytdlp":
        import yt_dlp
        sys.exit(yt_dlp.main(sys.argv[2:]))

    from stream_highlight.server import main as serve
    serve()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
