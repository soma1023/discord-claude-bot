# -*- coding: utf-8 -*-
"""同じ区間を2つの文字起こしエンジンで処理し、テキストと処理時間を並べる。

実際の配信素材で、どちらのエンジンを使うかを決めるためのもの。

    python -m stream_highlight.compare_engines 動画.mp4 --start 1500 --end 1800

VRAM 8GB では2つのエンジンを同時に載せられないので、1つずつ読み込み、
使い終わったら解放してから次を読み込む。GPUのメモリが足りないときは
CPU に切り替えて続ける（遅くなる）。キャッシュは使わず、毎回処理し直す。
"""

import argparse
import sys
import time

from . import transcribe


def _utf8_console():
    # 日本語Windowsのコンソールは cp932 なので、そのままでは書けない文字がある
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def compare(media_path, start, end, engines, report=print):
    """エンジンごとの結果を [{engine, device, load_sec, run_sec, text, ...}] で返す。"""
    samples = transcribe.extract_audio(media_path, start, end)
    length = len(samples) / float(transcribe.SAMPLE_RATE)
    rows = []
    for name in engines:
        row = {"engine": name, "length": length}
        problems = transcribe.missing(name)
        if problems:
            row["error"] = " ".join(problems)
            rows.append(row)
            continue
        try:
            began = time.time()
            engine = transcribe.holder.get(name, report=report)
            row["load_sec"] = time.time() - began
            row["device"] = engine.device

            began = time.time()
            raw = transcribe.run_engine(engine, samples, report=report)
            row["run_sec"] = time.time() - began
            row["device"] = engine.device       # 途中でCPUに切り替わることがある
            cleaned = transcribe.clean_words(raw)
            row["words"] = len(cleaned)
            row["loops_removed"] = len(raw) - len(cleaned)
            row["text"] = transcribe.join_words(w.text for w in cleaned)
            row["raw_text"] = transcribe.join_words(w.text for w in raw)
        except transcribe.TranscribeError as exc:
            row["error"] = str(exc)
        except Exception as exc:                  # noqa: BLE001（片方が失敗しても、もう片方は見せる）
            row["error"] = "%s: %s" % (type(exc).__name__, exc)
        finally:
            # 次のエンジンを読み込む前に、必ず解放する
            transcribe.holder.release()
        rows.append(row)
    return rows


def format_report(rows, media_path, start, end):
    lines = ["素材: %s（%.1f〜%.1f秒）" % (media_path, start, end), ""]
    lines.append("%-9s %-5s %8s %8s %8s %6s %6s" % (
        "エンジン", "装置", "読込(秒)", "処理(秒)", "実時間比", "語数", "ループ"))
    for row in rows:
        if "error" in row:
            lines.append("%-9s 失敗: %s" % (row["engine"], row["error"]))
            continue
        lines.append("%-9s %-5s %8.1f %8.1f %8.2f %6d %6d" % (
            row["engine"], row["device"], row["load_sec"], row["run_sec"],
            row["run_sec"] / max(row["length"], 0.001), row["words"], row["loops_removed"]))
    for row in rows:
        if "text" not in row:
            continue
        lines += ["", "=== %s ===" % row["engine"], row["text"] or "（何も聞き取れませんでした）"]
        if row["loops_removed"]:
            lines += ["--- ループ除去前 ---", row["raw_text"]]
    lines += ["", "実時間比は「処理にかかった秒 ÷ 音声の秒」。1より小さければ実時間より速い。"]
    return "\n".join(lines)


def main():
    _utf8_console()
    parser = argparse.ArgumentParser(description="文字起こしエンジンを同じ区間で比べる")
    parser.add_argument("media", help="動画または音声ファイル")
    parser.add_argument("--start", type=float, default=0.0, help="開始秒（既定 0）")
    parser.add_argument("--end", type=float, default=None, help="終了秒（既定 開始+300秒）")
    parser.add_argument("--engines", default="qwen,whisper",
                        help="比べるエンジン（カンマ区切り。既定 qwen,whisper）")
    parser.add_argument("--out", default="", help="結果をこのファイルにも保存する")
    args = parser.parse_args()

    end = args.end if args.end is not None else args.start + 300.0
    engines = [e.strip() for e in args.engines.split(",") if e.strip()]
    try:
        rows = compare(args.media, args.start, end, engines,
                       report=lambda m: m and print("  " + m, flush=True))
    except transcribe.TranscribeError as exc:
        print("失敗: %s" % exc)
        raise SystemExit(1)
    text = format_report(rows, args.media, args.start, end)
    print()
    print(text)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as fh:
            fh.write(text + "\n")
        print("\n%s に保存しました。" % args.out)


if __name__ == "__main__":
    main()
