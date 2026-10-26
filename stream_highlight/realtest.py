# -*- coding: utf-8 -*-
"""実際の配信で、自動カットとテロップを Qwen と whisper の両方で試す。

    python -m stream_highlight.realtest <配信URL> [--top 5] [--margin 60] [--out フォルダ]
    python -m stream_highlight.realtest <配信URL> --media 配信全体.mp4   手元の動画を使う

流れ:
  1. ツールと同じ解析（チャットの盛り上がり）で見せ場を出し、上位 --top 件を選ぶ
  2. 見せ場の前後 --margin 秒の区間だけを yt-dlp でダウンロードする（配信全体は落とさない）
  3. Qwen → whisper の順に、自動カット＋テロップで書き出す（モデルは1つずつ読み込む）
       <名前>_qwen.xml / <名前>_qwen.srt / <名前>_whisper.xml / <名前>_whisper.srt
  4. エンジンごとの数字と、怪しいテロップ・誤検出の候補を report.md にまとめる

ダウンロードした区間は parts/ に残し、2回目からは使い回す。文字起こしもキャッシュが
効くので、2回目の処理時間は「最初に起こしたときの秒数」を表示する。
"""

import argparse
import json
import os
import re
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))

from stream_highlight import autoedit, cache, editexport, paths, subtitles, transcribe  # noqa: E402
from stream_highlight import autocut as cutter                                         # noqa: E402
from stream_highlight.analyze import analyze                                            # noqa: E402
from stream_highlight.sources import (FetchError, _run_ytdlp, _ytdlp_error, fetch_chat,  # noqa: E402
                                      fetch_info, parse_url)

ENGINES = ["qwen", "whisper"]
DOWNLOAD_FORMAT = "bv*[height<=1080]+ba/b[height<=1080]/b"
# 配信の締めなどで、音が無いのにエンジンが作りがちな定型文
HALLUCINATIONS = ["ご視聴ありがとうございました", "チャンネル登録", "字幕", "おやすみなさい",
                  "thank you", "subscribe"]


def say(message):
    print(message, flush=True)


# ---------------------------------------------------------------- 1. 解析

def load_stream(url):
    """配信情報とチャット。保存済みのチャットがあれば使う。"""
    platform, video_id = parse_url(url)
    loaded = cache.load(platform, video_id)
    if loaded:
        say("保存済みのチャットを使います（%d件）" % len(loaded[1]))
        return loaded
    say("配信情報とチャットを取得しています…")
    info = fetch_info(url)
    messages = fetch_chat(info, progress=lambda m, f=None: None)
    if not messages:
        raise FetchError("チャットが1件も取得できませんでした。")
    cache.save(info, messages)
    return info, messages


def pick_peaks(info, messages, top):
    moments = analyze(messages, info)["moments"]
    if not moments:
        raise FetchError("見せ場が見つかりませんでした。")
    picked = sorted(moments, key=lambda m: m.get("rank", 0))[:top]
    for m in picked:
        say("  見せ場 %d位  %s  スコア %.1f" % (m["rank"], editexport.timecode(m["peak_sec"], 30),
                                           m["score"]))
    return [m["peak_sec"] for m in picked]


# ---------------------------------------------------------------- 2. ダウンロード

def probe(path):
    """ffmpeg の表示から (長さ秒, fps, 幅, 高さ) を読む。ffprobe が無くても動くように。"""
    out = subprocess.run([transcribe.find_ffmpeg(), "-hide_banner", "-i", path],
                         capture_output=True, text=True, encoding="utf-8", errors="replace",
                         creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)).stderr
    duration = re.search(r"Duration: (\d+):(\d+):([\d.]+)", out)
    seconds = (int(duration.group(1)) * 3600 + int(duration.group(2)) * 60
               + float(duration.group(3))) if duration else 0.0
    video = re.search(r"Video:.*?(\d{3,5})x(\d{3,5}).*?([\d.]+) fps", out)
    if not video:
        return seconds, 0.0, 0, 0
    return seconds, float(video.group(3)), int(video.group(1)), int(video.group(2))


REF_BEFORE_SEC = 10.0       # 照合用の音声は、区間の開始の10秒前から
REF_LENGTH_SEC = 30.0
CLIP_FROM_SEC, CLIP_TO_SEC = 1.0, 11.0     # 区間ファイルの冒頭のこの範囲を照らし合わせる
MIN_MATCH_SCORE = 0.5       # これより似ていなければ、測れなかったとみなす


def _ytdlp_section(url, fmt, start, end, out_template, exact):
    """yt-dlp で start〜end だけを落とす。かかった秒数を返す。"""
    args = ["-f", fmt, "--download-sections", "*%.3f-%.3f" % (start, end),
            "--ffmpeg-location", transcribe.find_ffmpeg(), "-o", out_template, url]
    if exact:
        args[0:0] = ["--force-keyframes-at-cuts"]
    if fmt == DOWNLOAD_FORMAT:
        args[0:0] = ["--merge-output-format", "mp4"]
    began = time.time()
    code, lines = _run_ytdlp(args, timeout=3600)
    if code != 0:
        raise FetchError(_ytdlp_error(lines))
    return time.time() - began


def _download_time(path, elapsed=None):
    """ダウンロードにかかった秒数を、ファイルの横に残す（2回目以降も報告できるように）。"""
    note = path + ".json"
    if elapsed is not None:
        with open(note, "w", encoding="utf-8") as fh:
            json.dump({"download_sec": round(elapsed, 1)}, fh)
        return round(elapsed, 1)
    try:
        with open(note, encoding="utf-8") as fh:
            return json.load(fh).get("download_sec")
    except (OSError, ValueError):
        return None


def _downloaded(template):
    """%(ext)s 付きの出力名から、実際にできたファイルを探す。"""
    folder, head = os.path.split(template.replace(".%(ext)s", ""))
    for name in sorted(os.listdir(folder)):
        if name.startswith(head + ".") and not name.endswith((".part", ".ytdl")):
            return os.path.join(folder, name)
    return ""


def find_lag(reference, clip, rate):
    """clip が reference のどこから始まるか（秒）と、似ている度合い（0〜1）。

    同じ配信の音声どうしなので、波形の相互相関（FFT）のいちばん高い所を取る。
    """
    import numpy as np

    ref = reference.astype("float64") - reference.mean()
    sub = clip.astype("float64") - clip.mean()
    if len(sub) == 0 or len(ref) < len(sub):
        return 0.0, 0.0
    size = 1 << int(np.ceil(np.log2(len(ref) + len(sub))))
    corr = np.fft.irfft(np.fft.rfft(ref, size) * np.conj(np.fft.rfft(sub, size)), size)
    corr = corr[:len(ref) - len(sub) + 1]
    # ずらした位置ごとの reference 側の大きさで割り、似ている度合いにする
    energy = np.concatenate([[0.0], np.cumsum(ref ** 2)])
    window = np.sqrt(np.maximum(energy[len(sub):] - energy[:-len(sub)], 1e-12))
    score = corr / (window * np.sqrt(np.sum(sub ** 2)) + 1e-12)
    best = int(np.argmax(score))
    return best / float(rate), float(score[best])


def measure_start(url, part_path, requested, folder, index):
    """区間ファイルの0秒が、実際には元の配信の何秒か。(秒, 似ている度合い)。

    区間の開始付近の音声だけを別に落とし（音声は約0.02秒単位で切れる）、
    区間ファイルの冒頭の音声がその中のどこにあるかを探す。
    """
    ref_start = max(0.0, requested - REF_BEFORE_SEC)
    template = os.path.join(folder, "ref%02d_%d.%%(ext)s" % (index, requested))
    ref_path = _downloaded(template)
    if not ref_path:
        _ytdlp_section(url, "ba[ext=m4a]/ba", ref_start, ref_start + REF_LENGTH_SEC,
                       template, exact=False)
        ref_path = _downloaded(template)
    reference = transcribe.extract_audio(ref_path, 0.0, REF_LENGTH_SEC)
    clip = transcribe.extract_audio(part_path, CLIP_FROM_SEC, CLIP_TO_SEC)
    lag, score = find_lag(reference, clip, transcribe.SAMPLE_RATE)
    return ref_start + lag - CLIP_FROM_SEC, score


def download_parts(url, segments, folder, check_copy=True):
    """区間ごとに、その範囲だけをダウンロードする。既にあれば使い回す。

    yt-dlp の区間指定はキーフレームや HLS の区切りで切るため、そのままだと
    ファイルの先頭が指定より数秒早くなることがある。そこで二重に対策する。
      1. --force-keyframes-at-cuts で、切れ目を再エンコードして指定の時刻で切る
      2. 落とした後に音声を照らし合わせて実際の先頭を測り、その値で時刻を換算する
    check_copy なら、1つ目の区間だけ再エンコードなしでも落とし、時間とずれを比べる。
    """
    os.makedirs(folder, exist_ok=True)
    media = []
    for index, segment in enumerate(segments, 1):
        start, end = segment["start"], segment["end"]
        path = os.path.join(folder, "part%02d_%d-%d.mp4" % (index, start, end))
        elapsed = None
        if not os.path.isfile(path):
            say("区間 %d/%d をダウンロード中（%s〜%s）…"
                % (index, len(segments), editexport.timecode(start, 30),
                   editexport.timecode(end, 30)))
            elapsed = _ytdlp_section(url, DOWNLOAD_FORMAT, start, end, path, exact=True)
            if not os.path.isfile(path):
                raise FetchError("区間 %d のダウンロード結果が見つかりません。" % index)
        duration, fps, width, height = probe(path)
        measured, score = measure_start(url, path, start, folder, index)
        offset = measured if score >= MIN_MATCH_SCORE else start
        say("  区間 %d: 実際の先頭は指定より %+.3f 秒（一致度 %.2f）%s"
            % (index, measured - start, score,
               "" if score >= MIN_MATCH_SCORE else " → 測れなかったので指定の時刻を使います"))
        item = {"path": path, "offset": offset, "duration": duration,
                "fps": fps, "width": width, "height": height,
                "requested": start, "measured": measured, "score": round(score, 3),
                "download_sec": _download_time(path, elapsed)}
        if check_copy and index == 1:
            item["copy_check"] = check_without_reencode(url, segment, folder)
        media.append(item)
    return media


def check_without_reencode(url, segment, folder):
    """比較用: 再エンコードせずに落としたら、時間とずれがどうなるか（1区間だけ）。"""
    start, end = segment["start"], segment["end"]
    path = os.path.join(folder, "check_copy_%d-%d.mp4" % (start, end))
    elapsed = None
    if not os.path.isfile(path):
        say("  比較のため、区間 1 を再エンコードなしでも落としています…")
        elapsed = _ytdlp_section(url, DOWNLOAD_FORMAT, start, end, path, exact=False)
    measured, score = measure_start(url, path, start, folder, 1)
    say("  再エンコードなしだと、先頭は指定より %+.3f 秒（一致度 %.2f）" % (measured - start, score))
    return {"download_sec": _download_time(path, elapsed),
            "delta": round(measured - start, 3), "score": round(score, 3),
            "duration": round(probe(path)[0], 2)}


# ---------------------------------------------------------------- 3. 書き出し

def export_engine(info, peaks, engine, args, media_path, segment_media, fmt_info, out_dir):
    options = autoedit.EditOptions(autocut=True, subtitles=True, engine=engine,
                                   cut=cutter.CutSettings())
    segments = editexport.build_segments(peaks, args.margin, info.duration)
    cached = all(transcribe.load_cached(*_cache_key(s, m, media_path), engine) is not None
                 for s, m in zip(segments, segment_media or [None] * len(segments)))
    began = time.time()
    result = autoedit.run(info, peaks, options, margin_sec=args.margin, fps=fmt_info["fps"],
                          media_path=media_path, width=fmt_info["width"],
                          height=fmt_info["height"], segment_media=segment_media,
                          report=lambda m, f=None: say("  [%s] %s" % (engine, m)))
    total = time.time() - began
    transcribe.holder.release()          # 次のエンジンを読む前に VRAM を空ける

    base = editexport.safe_name(info)
    xml_path = os.path.join(out_dir, "%s_%s.xml" % (base, engine))
    srt_path = os.path.join(out_dir, "%s_%s.srt" % (base, engine))
    with open(xml_path, "w", encoding="utf-8") as fh:
        fh.write(result["content"])
    with open(srt_path, "w", encoding="utf-8") as fh:
        fh.write(result["srt_content"])
    result.update(total_sec=round(total, 1), cached=cached, xml_path=xml_path, srt_path=srt_path)
    return result


def _cache_key(segment, media, media_path):
    if media:
        return (media["path"], max(0.0, segment["start"] - media["offset"]),
                segment["end"] - media["offset"])
    return (media_path, segment["start"], segment["end"])


def words_and_timeline(info, peaks, engine, args, media_path, segment_media, fps):
    """書き出しと同じ文字起こし（キャッシュから）とカット後の並びを作り直す。"""
    segments = editexport.build_segments(peaks, args.margin, info.duration)
    for segment, media in zip(segments, segment_media or [None] * len(segments)):
        if media:
            segment["media"] = media
    words_per_segment = [autoedit._transcribe(s, media_path, engine, None).words
                         for s in segments]
    cuts = [cutter.plan_cuts(s["start"], s["end"], w, s["peaks"])
            for s, w in zip(segments, words_per_segment)]
    return segments, words_per_segment, editexport.Timeline(segments, cuts, fps)


# ---------------------------------------------------------------- 4. 点検

def _source_time(timeline, seq_frame):
    for clip in timeline.clips:
        length = clip["src_out"] - clip["src_in"]
        if clip["seq_start"] <= seq_frame < clip["seq_start"] + length:
            return (clip["src_in"] + seq_frame - clip["seq_start"]) / timeline.fps
    return None


def suspicious_cues(cues, timeline):
    """中身が明らかにおかしそうなテロップ。(理由, シーケンス秒, 元の秒, 文字列)。"""
    found = []
    for cue in cues:
        text = cue["text"].replace("\n", "")
        plain = re.sub(r"[、。！？!?…\s]", "", text)
        reason = None
        if re.fullmatch(r"[wWｗ草はハあアー笑]+", plain):     # 笑い声は繰り返して当然
            continue
        if re.search(r"(.{2,}?)\1{2,}", plain):
            reason = "同じ語の繰り返し"
        elif re.search(r"(.)\1{7,}", plain):
            reason = "同じ文字の連続"
        elif any(h in text.lower() for h in HALLUCINATIONS):
            reason = "定型文（音の無い所で出やすい）"
        elif plain and len(re.findall(r"[A-Za-z]", plain)) / len(plain) > 0.5 and len(plain) >= 4:
            reason = "英字が多い"
        elif plain and len(re.findall(r"[^\w]", plain)) / len(plain) > 0.3:
            reason = "記号が多い"
        if reason:
            src = _source_time(timeline, cue["start_frame"])
            found.append((reason, cue["start_frame"] / timeline.fps, src, text))
    return found


def parse_srt(srt, fps):
    cues = []
    for block in srt.strip().split("\n\n"):
        lines = block.splitlines()
        if len(lines) < 3:
            continue
        start = lines[1].split(" --> ")[0]
        h, m, rest = start.split(":")
        s, ms = rest.split(",")
        seconds = int(h) * 3600 + int(m) * 60 + int(s) + int(ms) / 1000.0
        cues.append({"start_frame": int(round(seconds * fps)), "text": "\n".join(lines[2:])})
    return cues


def loudness(segment, media, media_path, step=0.5):
    """区間の音量（0.5秒ごとの RMS、dB）。ゲーム音や BGM が大きい所を探すため。"""
    import numpy as np

    samples = transcribe.extract_audio(*_cache_key(segment, media, media_path))
    hop = int(transcribe.SAMPLE_RATE * step)
    out = []
    for i in range(0, len(samples) - hop + 1, hop):
        rms = float(np.sqrt(np.mean(np.square(samples[i:i + hop].astype("float64"))))) + 1e-9
        out.append(20 * np.log10(rms))
    return out


def false_detection_candidates(segments, words_by_engine, segment_media, media_path):
    """発話の誤検出がありそうな所。

    - 片方のエンジンだけが「喋っている」と判定した所（もう片方は前後1秒に語が無い）
    - 音が大きい（区間内の上位20%）のに、前後1.5秒に他の語が無い短い語
    """
    found = []
    for index, segment in enumerate(segments):
        media = segment_media[index] if segment_media else None
        levels = loudness(segment, media, media_path)
        loud = sorted(levels)[int(len(levels) * 0.8)] if levels else None
        for engine, per_segment in words_by_engine.items():
            words = per_segment[index]
            others = [w for e, p in words_by_engine.items() if e != engine for w in p[index]]
            alone = [w for w in words
                     if not any(o.start - 1.0 <= w.end and w.start <= o.end + 1.0 for o in others)]
            for span in _merge_words(alone):
                found.append((engine, "%sだけが拾った" % engine, span[0], span[1], span[2]))
            for i, w in enumerate(words):
                prev_end = words[i - 1].end if i else -1e9
                next_start = words[i + 1].start if i + 1 < len(words) else 1e9
                isolated = w.start - prev_end >= 1.5 and next_start - w.end >= 1.5
                short = len(re.sub(r"[、。！？!?…\s]", "", w.text)) <= 2
                if not (isolated and short and loud is not None):
                    continue
                k = int((w.start - segment["start"]) / 0.5)
                if 0 <= k < len(levels) and levels[k] >= loud:
                    found.append((engine, "音が大きい所の孤立した短い語", w.start, w.end, w.text))
    return found


def _merge_words(words, gap=1.0):
    spans = []
    for w in words:
        if spans and w.start - spans[-1][1] <= gap:
            spans[-1] = (spans[-1][0], w.end, spans[-1][2] + w.text)
        else:
            spans.append((w.start, w.end, w.text))
    return spans


# ---------------------------------------------------------------- まとめ

def tc(seconds):
    return "-" if seconds is None else editexport.timecode(seconds, 30)[:8]


def alignment_lines(segment_media, media_path):
    """区間ダウンロードの位置合わせ（どう対策し、実際どれだけずれていたか）。"""
    lines = ["## 区間ダウンロードの位置合わせ", ""]
    if not segment_media:
        return lines + ["手元の動画（%s）を使ったので、ダウンロードはしていない。" % media_path, ""]
    lines += [
        "対策: 両方を行った。",
        "1. `--force-keyframes-at-cuts` で切れ目を再エンコードし、指定の時刻で切る",
        "2. 落とした後、区間の開始付近の音声だけを別に落として照らし合わせ、"
        "実際の先頭が元の配信の何秒かを測る。一致度が %.1f 以上なら、その値で時刻を換算する"
        % MIN_MATCH_SCORE, "",
        "| 区間 | 指定の開始 | 実際の先頭のずれ | 一致度 | 換算に使った値 | ダウンロード時間 |",
        "| --- | --- | --- | --- | --- | --- |"]
    for index, m in enumerate(segment_media, 1):
        used = "実測" if m["score"] >= MIN_MATCH_SCORE else "指定（測れず）"
        lines.append("| %d | %s（%.2f秒） | %+.3f秒 | %.2f | %s | %s |"
                     % (index, tc(m["requested"]), m["requested"], m["measured"] - m["requested"],
                        m["score"], used,
                        "-" if m["download_sec"] is None else "%.1f秒" % m["download_sec"]))
    check = segment_media[0].get("copy_check")
    if check:
        first = segment_media[0]
        lines += ["", "比較（区間 1 を再エンコードなしで落とした場合）:", "",
                  "| | 再エンコードあり（採用） | 再エンコードなし |", "| --- | --- | --- |",
                  "| ダウンロード時間 | %s | %s |" % (
                      "-" if first["download_sec"] is None else "%.1f秒" % first["download_sec"],
                      "-" if check["download_sec"] is None else "%.1f秒" % check["download_sec"]),
                  "| 実際の先頭のずれ | %+.3f秒 | %+.3f秒 |" % (
                      first["measured"] - first["requested"], check["delta"]),
                  "| ファイルの長さ | %.2f秒 | %.2f秒 |" % (first["duration"], check["duration"])]
    return lines + [""]


def write_report(path, info, peaks, results, checks, detections, alignment=()):
    lines = ["# 実配信テスト: %s" % (info.title or info.video_id), "",
             "- URL: %s" % info.url,
             "- 見せ場: %s" % " / ".join(tc(p) for p in sorted(peaks)), ""]
    lines += list(alignment)
    lines += ["## エンジンごとの結果", "",
              "| | " + " | ".join(results) + " |",
              "| --- | " + " | ".join("---" for _ in results) + " |"]

    def row(label, fn):
        lines.append("| %s | %s |" % (label, " | ".join(fn(r) for r in results.values())))
    row("使った装置", lambda r: r["device"] or "-")
    row("文字起こしの時間", lambda r: "%.1f秒%s" % (r["transcribe_sec"],
                                                 "（キャッシュ）" if r["cached"] else ""))
    row("全体の時間", lambda r: "%.1f秒" % r["total_sec"])
    row("元の長さ → カット後", lambda r: "%s → %s" % (tc(r["original_sec"]), tc(r["edited_sec"])))
    row("カットの数", lambda r: "%d" % (r["clips"] - r["segments"]))
    row("テロップの数", lambda r: "%d" % r["cues"])
    row("おかしなテロップの候補", lambda r: "%d件" % len(checks[r["engine"]]))
    row("誤検出の候補", lambda r: "%d件" % sum(1 for d in detections if d[0] == r["engine"]))
    for engine in results:
        lines += ["", "## %s: おかしなテロップの候補" % engine, ""]
        if not checks[engine]:
            lines.append("なし")
        for reason, seq, src, text in checks[engine]:
            lines.append("- シーケンス %s（元 %s）%s: 「%s」" % (tc(seq), tc(src), reason, text))
        lines += ["", "## %s: 誤検出の候補（元の配信の時刻）" % engine, ""]
        mine = [d for d in detections if d[0] == engine]
        if not mine:
            lines.append("なし")
        for _, reason, start, end, text in mine:
            lines.append("- %s〜%s %s: 「%s」" % (tc(start), tc(end), reason, text))
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")
    return "\n".join(lines)


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        stream.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser(description="実際の配信で、自動カットとテロップを2つのエンジンで試す")
    parser.add_argument("url")
    parser.add_argument("--top", type=int, default=5, help="書き出す見せ場の数（既定5）")
    parser.add_argument("--margin", type=float, default=60.0, help="見せ場の前後の秒数（既定60）")
    parser.add_argument("--media", default="", help="手元にある配信全体の動画。あればダウンロードしない")
    parser.add_argument("--engines", default=",".join(ENGINES))
    parser.add_argument("--out", default="", help="出力先（既定はデータフォルダの realtest/動画ID）")
    parser.add_argument("--no-copy-check", action="store_true",
                        help="区間 1 を再エンコードなしでも落として比べる、のを省く")
    args = parser.parse_args(argv)

    info, messages = load_stream(args.url)
    out_dir = args.out or os.path.join(paths.data_dir(), "realtest",
                                       "%s_%s" % (info.platform, info.video_id))
    os.makedirs(out_dir, exist_ok=True)
    peaks = pick_peaks(info, messages, max(1, args.top))
    segments = editexport.build_segments(peaks, args.margin, info.duration)
    say("区間 %d 個（合計 %.1f 分）" % (len(segments),
                                   sum(s["end"] - s["start"] for s in segments) / 60))

    engines = [e.strip() for e in args.engines.split(",") if e.strip()]
    for engine in engines:
        problems = transcribe.missing(engine)
        if problems:
            raise SystemExit(" ".join(problems))

    if args.media:
        if not os.path.isfile(args.media):
            raise SystemExit("動画ファイルが見つかりません: %s" % args.media)
        segment_media, media_path = None, os.path.abspath(args.media)
        _, fps, width, height = probe(media_path)
    else:
        segment_media = download_parts(args.url, segments, os.path.join(out_dir, "parts"),
                                       check_copy=not args.no_copy_check)
        media_path = ""
        fps, width, height = (segment_media[0]["fps"], segment_media[0]["width"],
                              segment_media[0]["height"])
    fmt_info = {"fps": fps or info.fps or editexport.DEFAULT_FPS,
                "width": width or info.width, "height": height or info.height}

    results, checks, words_by_engine = {}, {}, {}
    for engine in engines:
        say("%s で書き出しています…" % engine)
        result = export_engine(info, peaks, engine, args, media_path, segment_media,
                               fmt_info, out_dir)
        results[engine] = result
        _, words, timeline = words_and_timeline(info, peaks, engine, args, media_path,
                                                segment_media, float(fmt_info["fps"]))
        words_by_engine[engine] = words
        checks[engine] = suspicious_cues(parse_srt(result["srt_content"], timeline.fps), timeline)
        say("  → %s / %s" % (result["xml_path"], result["srt_path"]))
    transcribe.holder.release()

    detections = false_detection_candidates(segments, words_by_engine, segment_media,
                                            media_path) if len(words_by_engine) > 1 else []
    report = write_report(os.path.join(out_dir, "report.md"), info, peaks, results, checks,
                          detections, alignment_lines(segment_media, media_path))
    with open(os.path.join(out_dir, "report.json"), "w", encoding="utf-8") as fh:
        json.dump({"engines": {e: {k: v for k, v in r.items()
                                   if k not in ("content", "srt_content")}
                               for e, r in results.items()},
                   "parts": segment_media or []}, fh, ensure_ascii=False, indent=1)
    say("")
    say(report)
    say("")
    say("出力先: %s" % out_dir)


if __name__ == "__main__":
    try:
        main()
    except FetchError as exc:
        raise SystemExit(str(exc))
