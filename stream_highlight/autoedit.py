# -*- coding: utf-8 -*-
"""自動カットとテロップを付けた書き出しを、区間ごとに組み立てる。

流れ:
  1. 書き出す区間を決める（editexport.build_segments と同じ）
  2. 区間ごとに音声を切り出して文字起こし（キャッシュがあれば使う）
  3. 自動カット: 沈黙（と任意で言い直し）を除いた「残す部分」を決める
  4. XML / EDL を書き、テロップ（SRT）をカット後の時刻で作る

文字起こしは時間がかかるので、jobs.py のバックグラウンド処理から呼ぶ。
"""

from dataclasses import dataclass

from . import editexport, retake, subtitles, transcribe
from . import autocut as cutter
# EditOptions の subtitles（ON/OFF）が同名のモジュールを隠すので、クラスは直接読み込む
from .subtitles import CaptionStyle


@dataclass
class EditOptions:
    autocut: bool = False
    subtitles: bool = False
    retake: bool = False
    engine: str = transcribe.DEFAULT_ENGINE
    cut: cutter.CutSettings = None
    caption: CaptionStyle = None

    @property
    def enabled(self):
        return self.autocut or self.subtitles


def problems(options, media_path):
    """始める前に分かる問題（日本語）。空なら始めてよい。"""
    import os

    if not media_path:
        return ["自動カット・テロップには「動画ファイルの場所」が必要です。"]
    if not os.path.isfile(media_path):
        return ["動画ファイルが見つかりません: %s" % media_path]
    return transcribe.missing(options.engine)


def _transcribe(segment, media_path, engine, report):
    """区間を文字起こしする。区間ごとの動画ファイルがあればそちらを使う。"""
    media = segment.get("media")
    if not media:
        return transcribe.transcribe_segment(
            media_path, segment["start"], segment["end"], engine, report=report)
    # ファイルの中の時刻で起こし、元の配信の時刻に戻す
    offset = media["offset"]
    found = transcribe.transcribe_segment(
        media["path"], max(0.0, segment["start"] - offset), segment["end"] - offset,
        engine, report=report)
    found.words = [transcribe.Word(w.text, w.start + offset, w.end + offset)
                   for w in found.words]
    return found


def run(info, peaks, options, fmt="xml", margin_sec=editexport.DEFAULT_MARGIN_SEC,
        fps=editexport.DEFAULT_FPS, media_path="", width=0, height=0, report=None,
        segment_media=None):
    """書き出し一式を作る。戻り値は画面にそのまま返す dict。

    report(message, frac) で進み具合を伝える。frac は 0〜1。
    segment_media は区間ごとの動画ファイル {"path", "offset", "duration"} の並び
    （build_segments の区間と同じ順）。区間付近だけをダウンロードしたときに渡す。
    """
    fps = float(fps) or editexport.DEFAULT_FPS
    segments = editexport.build_segments(peaks, margin_sec, info.duration)
    if not segments:
        raise ValueError("書き出す候補がありません。")
    if segment_media is not None:
        if len(segment_media) != len(segments):
            raise ValueError("区間の数と動画ファイルの数が合いません。")
        for segment, media in zip(segments, segment_media):
            segment["media"] = media
    settings = options.cut or cutter.CutSettings()

    def say(message, frac=None):
        if report:
            report(message, frac)

    transcripts = []
    for index, segment in enumerate(segments):
        head = "区間 %d/%d を文字起こし中" % (index + 1, len(segments))
        say(head + "…", index / float(len(segments)))

        def detail(message, head=head):
            say("%s%s" % (head, message) if message.startswith("（") else message)
        transcripts.append(_transcribe(segment, media_path, options.engine, detail))

    notes = []
    cuts = None
    removed_retakes = []
    if options.autocut:
        cuts = []
        retake_failed = False
        for index, (segment, transcript) in enumerate(zip(segments, transcripts)):
            extra = []
            if options.retake and not retake_failed:
                say("言い直しを探しています（区間 %d/%d）…" % (index + 1, len(segments)), 0.95)
                try:
                    extra = retake.detect(transcript.words)
                    removed_retakes.extend(retake.describe(transcript.words, extra))
                except retake.RetakeError as exc:
                    # 言い直しの削除だけを諦める。書き出しは続ける
                    retake_failed = True
                    notes.append("言い直しの削除は行いませんでした: %s" % exc)
            cuts.append(cutter.plan_cuts(segment["start"], segment["end"], transcript.words,
                                          segment["peaks"], settings, extra))

    say("書き出しています…", 0.98)
    safe = editexport.safe_name(info)
    if fmt == "edl":
        filename = "%s.edl" % safe
        content = editexport.to_edl(info, segments, fps, media_path, cuts)
    else:
        filename = "%s.xml" % safe
        content = editexport.to_fcp7_xml(info, segments, fps, media_path, width, height, cuts)

    timeline = editexport.Timeline(segments, cuts, fps)
    result = {"filename": filename, "content": content, "segments": len(segments),
              "fps": fps, "width": width, "height": height,
              "clips": len(timeline.clips), "notes": notes,
              "retakes": removed_retakes,
              "engine": options.engine,
              "device": next((t.device for t in transcripts if t.device), ""),
              # キャッシュを使った区間は、最初に起こしたときの秒数
              "transcribe_sec": round(sum(t.elapsed for t in transcripts), 1),
              "original_sec": round(sum(s["end"] - s["start"] for s in segments), 1),
              "edited_sec": round(timeline.total_frames / fps, 1)}
    if options.subtitles:
        words = [w for t in transcripts for w in t.words]
        cues = subtitles.make_cues(words, timeline, options.caption)
        result["srt_filename"] = "%s.srt" % safe
        result["srt_content"] = subtitles.to_srt(cues)
        result["cues"] = len(cues)
    return result
