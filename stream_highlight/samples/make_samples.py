# -*- coding: utf-8 -*-
"""Premiere での読み込みを確かめるための、小さな XML と SRT を作る。

文字起こしは合成データ（モデルは使わない）。10分の架空の配信に、
ピークを2つ置き、その前後1分ずつを書き出す。

    python -m stream_highlight.samples.make_samples          XML と SRT だけ作る
    python -m stream_highlight.samples.make_samples --video  確認用の動画も作り、XML からつなぐ

--video で作る動画は、画面に経過秒数のカウンタが出る（ffmpeg の testsrc）。
シーケンスで再生したときのカウンタの飛び方が、カットの位置そのもの。
テロップは「12秒 です」のように元の配信の秒を読み上げる体裁にしてあるので、
画面のカウンタとテロップの数字が合っていれば、SRT の時刻も正しい。
"""

import argparse
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(os.path.dirname(HERE)))

from stream_highlight import autocut, editexport, subtitles, transcribe   # noqa: E402
from stream_highlight.sources import StreamInfo                           # noqa: E402
from stream_highlight.transcribe import Word                              # noqa: E402

FPS = 29.97
WIDTH, HEIGHT = 1280, 720
DURATION = 600.0
PEAKS = [200.0, 420.0]
MARGIN = 60.0
VIDEO_NAME = "sample_source.mp4"


def synthetic_words():
    """数秒しゃべっては黙る、を繰り返す架空の発話。語の中身は元の秒数。"""
    words = []
    t = 2.0
    pattern = [(4.0, 0.6), (3.0, 3.5), (5.0, 0.8), (2.5, 6.0), (4.0, 2.2)]
    k = 0
    while t < DURATION - 5:
        talk, pause = pattern[k % len(pattern)]
        end = t + talk
        while t < end:
            words.append(Word("%d秒" % int(t), t, t + 0.45))
            words.append(Word("です。", t + 0.5, t + 0.9))
            t += 1.0
        t = end + pause
        k += 1
    return words


def build(media_path=""):
    info = StreamInfo("youtube", "sample", "https://example.invalid/sample",
                      title="自動カットのサンプル", duration=DURATION,
                      fps=FPS, width=WIDTH, height=HEIGHT)
    words = synthetic_words()
    segments = editexport.build_segments(PEAKS, MARGIN, DURATION)
    cuts = [autocut.plan_cuts(s["start"], s["end"], words, s["peaks"]) for s in segments]
    xml = editexport.to_fcp7_xml(info, segments, FPS, media_path, WIDTH, HEIGHT, cuts)
    timeline = editexport.Timeline(segments, cuts, FPS)
    srt = subtitles.to_srt(subtitles.make_cues(words, timeline))
    return xml, srt, timeline


def make_video(path):
    ffmpeg = transcribe.find_ffmpeg()
    if not ffmpeg:
        raise SystemExit("ffmpeg が見つかりません。")
    args = [ffmpeg, "-y", "-v", "error",
            "-f", "lavfi", "-i", "testsrc=size=%dx%d:rate=30000/1001:duration=%d"
            % (WIDTH, HEIGHT, DURATION),
            "-f", "lavfi", "-i", "sine=frequency=440:sample_rate=48000:duration=%d" % DURATION,
            "-ac", "2", "-c:v", "libx264", "-preset", "veryfast", "-crf", "35",
            "-pix_fmt", "yuv420p", "-c:a", "aac", "-b:a", "64k", "-shortest", path]
    subprocess.run(args, check=True)


def write(name, text):
    path = os.path.join(HERE, name)
    with open(path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(text)
    return path


def main():
    parser = argparse.ArgumentParser(description="Premiere 確認用のサンプルを作る")
    parser.add_argument("--video", action="store_true",
                        help="確認用の動画も作り、XML からつなぐ（数十秒かかる）")
    args = parser.parse_args()
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass

    if args.video:
        video = os.path.join(HERE, VIDEO_NAME)
        print("確認用の動画を作っています… %s" % video)
        make_video(video)
        xml, srt, timeline = build(video)
        print("作りました:")
        print("  %s" % write("sample_linked.xml", xml))
        print("  %s" % write("sample_linked.srt", srt))
    else:
        xml, srt, timeline = build()
        print("作りました:")
        print("  %s" % write("sample_autocut.xml", xml))
        print("  %s" % write("sample_autocut.srt", srt))
    print("区間 %d 個 → クリップ %d 個 / %.1f秒 → %.1f秒"
          % (len(PEAKS), len(timeline.clips), 2 * MARGIN * 2, timeline.total_frames / FPS))


if __name__ == "__main__":
    main()
