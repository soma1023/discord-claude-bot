# -*- coding: utf-8 -*-
"""手元の動画（mp4）から、Premiere にそのまま読み込めるカット済みシーケンスとテロップを作る。

  python -m stream_highlight.autoedit input.mp4 --profile nekomiya

流れ:
  1. 書き出す区間を決める（動画全体、またはハイライト抽出の見せ場の前後）
  2. 区間ごとに文字起こし（faster-whisper。結果はキャッシュし、2回目は飛ばす）
  3. カットする所を決める
       沈黙（単語の間 / auto-editor の音量）＋ フィラー → 短すぎる断片は前後とつなぐ
  4. 書き出す（既定は入力と同じ場所の「<名前>」フォルダ）
       <名前>.xml              カット済みシーケンス（1カット＝1クリップ。端を引き出せば戻せる）
       <名前>.srt              テロップ（カット後の時刻）
       <名前>_transcript.json  文字起こしの全文とカット区間（確認用）

XML は auto-editor に書かせず、アプリと同じ editexport で書く。テロップと同じ
フレームの丸め（editexport.Timeline）を使うので、カット点とテロップがずれない。
auto-editor は、音量で沈黙を判定する所（silence_method: audio）だけに使う。

ズームや効果音を足すときは、EditPlan.effects に {"kind", "at", ...} を積み、
write_outputs で書き出す（今は空）。
"""

import argparse
import copy
import json
import os
import re
import subprocess
import sys
import tempfile
import time
from dataclasses import asdict, dataclass, field

from . import deps, editexport, transcribe
from . import autocut as cutter
from .sources import FetchError, StreamInfo
from .subtitles import CaptionStyle, make_cues, to_srt

PROFILE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config", "profiles")
DEFAULT_PROFILE = "default"
ENGINES = ("whisper-large-v3", "whisper", "qwen")
SILENCE_METHODS = ("speech", "audio")
DEFAULT_TOP = 5
DEFAULT_MARGIN_SEC = 60.0
# 語の頭に付いたフィラー（「えーとそれで」）を切り分けるのは、この文字数以上のものだけ。
# 「まあ」「えー」は「まあまあ」「えーマジ」のように普通の語の頭にも来る
MIN_PREFIX_FILLER_CHARS = 3
PREFIX_SAFETY_SEC = 0.05      # 見積もった境目から、後ろの語に食い込まないよう控える秒数
REACTION_MARKS = "!！?？"     # これが付いた「えー！」は驚きの反応なのでフィラーとみなさない


class EditError(RuntimeError):
    """利用者に見せる日本語メッセージを持つ失敗。"""


# ---------------------------------------------------------------- プロファイル

def available_profiles():
    try:
        return sorted(name[:-5] for name in os.listdir(PROFILE_DIR) if name.endswith(".yaml"))
    except OSError:
        return []


def profile_path(name):
    """名前（nekomiya）なら config/profiles/ の yaml、パスならそのまま。"""
    if name.endswith((".yaml", ".yml")) or "/" in name or "\\" in name:
        return name
    return os.path.join(PROFILE_DIR, name + ".yaml")


def _read_yaml(path):
    import yaml

    if not os.path.isfile(path):
        raise EditError("プロファイルが見つかりません: %s\n使えるもの: %s"
                        % (path, ", ".join(available_profiles()) or "なし"))
    try:
        with open(path, encoding="utf-8") as fh:
            data = yaml.safe_load(fh) or {}
    except yaml.YAMLError as exc:
        raise EditError("プロファイルの書き方に誤りがあります（%s）:\n%s" % (path, exc))
    if not isinstance(data, dict):
        raise EditError("プロファイルの中身は「項目: 値」の形で書いてください: %s" % path)
    return data


def _deep_merge(base, extra):
    out = copy.deepcopy(base)
    for key, value in (extra or {}).items():
        if isinstance(value, dict) and isinstance(out.get(key), dict):
            out[key] = _deep_merge(out[key], value)
        else:
            out[key] = copy.deepcopy(value)
    return out


def load_profile(name=DEFAULT_PROFILE):
    """default.yaml に、指定のプロファイルで書いた所だけを上書きしたもの。"""
    data = _read_yaml(profile_path(DEFAULT_PROFILE))
    if name and name != DEFAULT_PROFILE:
        data = _deep_merge(data, _read_yaml(profile_path(name)))
    return data


def flatten_profile(data):
    """yaml の入れ子を、Settings の項目名の平らな dict にする。"""
    t = data.get("transcribe") or {}
    c = data.get("cut") or {}
    audio = c.get("audio") or {}
    fillers = c.get("fillers") or {}
    cap = data.get("caption") or {}
    flat = {
        "profile": data.get("name"),
        "engine": t.get("engine"), "filler_prompt": t.get("filler_prompt"),
        "prompt_text": t.get("prompt_text"),
        "silence_method": c.get("silence_method"), "gap_cut_sec": c.get("gap_cut_sec"),
        "gap_cut": c.get("gap_cut"), "pad_sec": c.get("pad_sec"),
        "min_piece_sec": c.get("min_piece_sec"), "protect_sec": c.get("protect_sec"),
        "audio_threshold": audio.get("threshold"), "audio_margin_sec": audio.get("margin_sec"),
        "audio_track": c.get("audio_track"),
        "fillers": fillers.get("enabled"), "filler_words": fillers.get("words"),
        "line_chars": cap.get("line_chars"), "max_lines": cap.get("max_lines"),
        "min_show_sec": cap.get("min_show_sec"), "max_show_sec": cap.get("max_show_sec"),
        "strip_trailing_punct": cap.get("strip_trailing_punct"),
        "replacements": cap.get("replacements"),
    }
    return {key: value for key, value in flat.items() if value is not None}


# 数の項目: (下限, 上限, 型)
NUMBERS = {
    "gap_cut_sec": (0.1, 10.0, float), "pad_sec": (0.0, 1.0, float),
    "min_piece_sec": (0.0, 5.0, float), "protect_sec": (0.0, 120.0, float),
    "audio_threshold": (0.001, 1.0, float), "audio_margin_sec": (0.0, 5.0, float),
    "audio_track": (0, 31, int),
    "line_chars": (4, 40, int), "max_lines": (1, 4, int),
    "min_show_sec": (0.1, 5.0, float), "max_show_sec": (0.5, 15.0, float),
}
BOOLS = ("filler_prompt", "gap_cut", "fillers", "strip_trailing_punct")
CAPTION_FIELDS = ("line_chars", "max_lines", "min_show_sec", "max_show_sec",
                  "strip_trailing_punct", "replacements")


@dataclass
class Settings:
    profile: str = DEFAULT_PROFILE
    engine: str = "whisper-large-v3"
    filler_prompt: bool = True
    prompt_text: str = ""
    silence_method: str = "speech"
    gap_cut_sec: float = 0.4
    gap_cut: bool = True
    pad_sec: float = 0.15
    min_piece_sec: float = 0.3
    protect_sec: float = 5.0
    audio_threshold: float = 0.04
    audio_margin_sec: float = 0.2
    audio_track: int = 0
    fillers: bool = True
    filler_words: list = field(default_factory=list)
    caption: CaptionStyle = field(default_factory=CaptionStyle)

    @property
    def prompt(self):
        """whisper に渡す initial_prompt。Qwen は使わない。"""
        return self.prompt_text if self.filler_prompt and self.engine != "qwen" else ""

    def cut_settings(self):
        # audio で gap_cut を切ったときは、単語の間では詰めない（無限大の間＝詰めない）
        use_gap = self.silence_method == "speech" or self.gap_cut
        return cutter.CutSettings(gap_sec=self.gap_cut_sec if use_gap else float("inf"),
                                  pad_sec=self.pad_sec, protect_sec=self.protect_sec,
                                  min_piece_sec=self.min_piece_sec)

    def as_dict(self):
        return asdict(self)


def build_settings(flat):
    """平らな dict（プロファイル＋引数の上書き）を確かめて Settings にする。"""
    values = {}
    for key, value in flat.items():
        if key in NUMBERS:
            low, high, kind = NUMBERS[key]
            try:
                number = kind(value)
            except (TypeError, ValueError):
                raise EditError("「%s」は数で書いてください（今の値: %r）。" % (key, value))
            if not low <= number <= high:
                raise EditError("「%s」は %s〜%s の範囲にしてください（今の値: %s）。"
                                % (key, low, high, number))
            values[key] = number
        elif key in BOOLS:
            if not isinstance(value, bool):
                raise EditError("「%s」は true か false で書いてください（今の値: %r）。" % (key, value))
            values[key] = value
        elif key == "engine":
            if value not in ENGINES:
                raise EditError("engine は %s のどれかにしてください（今の値: %s）。"
                                % (" / ".join(ENGINES), value))
            values[key] = value
        elif key == "silence_method":
            if value not in SILENCE_METHODS:
                raise EditError("silence_method は speech か audio にしてください（今の値: %s）。"
                                % value)
            values[key] = value
        elif key == "filler_words":
            if not isinstance(value, list):
                raise EditError("fillers.words は [えー, えっと] のような並びで書いてください。")
            values[key] = [str(w) for w in value if str(w).strip()]
        elif key == "replacements":
            if not isinstance(value, dict):
                raise EditError("replacements は「元の表記: 直した表記」の形で書いてください。")
            values[key] = {str(k): str(v) for k, v in value.items() if str(k)}
        elif key in ("profile", "prompt_text"):
            values[key] = str(value)
    caption = {k: values.pop(k) for k in CAPTION_FIELDS if k in values}
    settings = Settings(**values)
    settings.caption = CaptionStyle(**caption)
    if settings.caption.max_show_sec < settings.caption.min_show_sec:
        raise EditError("max_show_sec は min_show_sec 以上にしてください。")
    return settings


# ---------------------------------------------------------------- 区間

def whole_file(duration):
    return [{"start": 0.0, "end": float(duration), "peaks": []}]


def read_segments_file(path, duration):
    """区間を書いた JSON（[[開始秒, 終了秒], ...] か [{"start", "end"}, ...]）を読む。"""
    try:
        with open(path, encoding="utf-8") as fh:
            rows = json.load(fh)
    except (OSError, ValueError) as exc:
        raise EditError("区間のファイルを読めません: %s（%s）" % (path, exc))
    spans = []
    for row in rows if isinstance(rows, list) else []:
        try:
            if isinstance(row, dict):
                start, end = float(row["start"]), float(row["end"])
                peaks = [float(p) for p in row.get("peaks") or []]
            else:
                start, end, peaks = float(row[0]), float(row[1]), []
        except (KeyError, IndexError, TypeError, ValueError):
            raise EditError("区間の書き方が違います: %r（[開始秒, 終了秒] で書いてください）" % (row,))
        start, end = max(0.0, start), min(float(duration), end)
        if end > start:
            spans.append({"start": start, "end": end, "peaks": peaks})
    if not spans:
        raise EditError("区間のファイルに、動画の長さに収まる区間がありません: %s" % path)
    spans.sort(key=lambda s: s["start"])
    merged = [spans[0]]
    for span in spans[1:]:
        if span["start"] <= merged[-1]["end"]:
            merged[-1]["end"] = max(merged[-1]["end"], span["end"])
            merged[-1]["peaks"] += span["peaks"]
        else:
            merged.append(span)
    return merged


def highlight_segments(source, top, margin_sec, duration):
    """ハイライト抽出の見せ場（上位 top 件）の前後 margin_sec 秒。動画は配信全体である前提。"""
    from . import realtest

    info, messages = realtest.load_stream(source)
    peaks = realtest.pick_peaks(info, messages, top)
    segments = editexport.build_segments(peaks, margin_sec, duration)
    if not segments:
        raise EditError("見せ場が動画の長さに収まりません。配信全体の動画を指定してください。")
    return segments


# ---------------------------------------------------------------- カット

_silence_cache = {}      # 同じ動画・同じ設定なら、音量判定をやり直さない（パネルで何度も呼ぶため）


def audio_silence(media_path, duration, settings):
    """auto-editor の音量判定で「静か」とされた所 [(開始秒, 終了秒)]。"""
    key = (os.path.abspath(media_path), transcribe._media_stamp(media_path),
           settings.audio_threshold, settings.audio_margin_sec, settings.audio_track)
    if key not in _silence_cache:
        _silence_cache.clear()           # 長い配信の結果を溜め込まない。最後の1本だけ持つ
        _silence_cache[key] = _audio_silence(media_path, duration, settings)
    return _silence_cache[key]


def _audio_silence(media_path, duration, settings):
    command = deps.auto_editor_command()
    with tempfile.TemporaryDirectory(prefix="sh_autoeditor_") as folder:
        out = os.path.join(folder, "timeline.v3")
        args = command + [media_path,
                          "--edit", "audio:threshold=%g,stream=%d"
                          % (settings.audio_threshold, settings.audio_track),
                          "--margin", "%gs" % settings.audio_margin_sec,
                          "--export", "v3", "-o", out, "--no-open", "--progress", "none"]
        done = subprocess.run(args, capture_output=True, text=True, encoding="utf-8",
                              errors="replace", creationflags=transcribe._no_window())
        if done.returncode != 0 or not os.path.isfile(out):
            detail = (done.stderr or done.stdout or "").strip().splitlines()
            raise EditError("auto-editor の音量判定に失敗しました: %s"
                            % (detail[-1] if detail else "原因不明"))
        with open(out, encoding="utf-8") as fh:
            timeline = json.load(fh)
    return silence_from_v3(timeline, duration)


def silence_from_v3(timeline, duration):
    """auto-editor の v3 タイムライン（残すクリップの並び）から、消された所を求める。"""
    num, _, den = str(timeline.get("timebase") or "30/1").partition("/")
    base = float(num) / float(den or 1)
    tracks = (timeline.get("v") or []) + (timeline.get("a") or [])
    clips = next((track for track in tracks if track), [])
    kept = [(c["offset"] / base, (c["offset"] + c["dur"]) / base) for c in clips]
    return cutter._subtract([(0.0, float(duration))], kept)


def _normalize(text):
    text = text.replace("〜", "ー").replace("～", "ー")
    text = re.sub(r"[、。，．,.・…\s]", "", text)
    return re.sub("ー+", "ー", text)


def filler_ranges(words, filler_words):
    """フィラーの語の範囲 [(開始秒, 終了秒, 語)]。

    語がフィラーと一致すれば、その語を丸ごと。語の頭にフィラーが付いている
    （「えーとそれで」）ときは、文字数の割合で境目を見積もり、後ろの語に
    食い込まないよう少し手前まで。「えー！」のような驚きの反応は消さない。
    """
    table = sorted({_normalize(w) for w in filler_words if _normalize(w)}, key=len, reverse=True)
    out = []
    for w in words:
        if any(mark in w.text for mark in REACTION_MARKS):
            continue
        norm = _normalize(w.text)
        if not norm:
            continue
        if norm in table:
            out.append((w.start, w.end, w.text))
            continue
        for filler in table:
            if len(filler) >= MIN_PREFIX_FILLER_CHARS and norm.startswith(filler) \
                    and len(norm) > len(filler):
                cut_end = w.start + (w.end - w.start) * len(filler) / len(norm) - PREFIX_SAFETY_SEC
                if cut_end > w.start:
                    out.append((w.start, cut_end, filler))
                break
    return out


@dataclass
class EditPlan:
    """書き出す中身。時刻は元の動画の秒、timeline と cues はカット後。"""
    segments: list
    words: list                 # 全区間の語
    cuts: list                  # 区間ごとの残す部分 [(開始秒, 終了秒)]
    removed: list               # 消した所 [{"start", "end", "reasons"}]
    timeline: object
    cues: list
    effects: list = field(default_factory=list)    # 将来: ズーム・効果音 {"kind", "at", ...}


def _overlaps(lo, hi, ranges):
    return [label for r_lo, r_hi, label in ranges if r_lo < hi and r_hi > lo]


def plan_segment(segment, words, silence, settings):
    """1つの区間で残す部分 [(開始秒, 終了秒)] と、消した所 [{"start", "end", "reasons"}]。

    CLI（build_plan）と Premiere のパネル（analyze_range）の両方がここで判定する。
    """
    cut_settings = settings.cut_settings()
    start, end = segment["start"], segment["end"]
    fillers = filler_ranges(words, settings.filler_words) if settings.fillers else []
    quiet = [(lo, hi) for lo, hi in silence if hi > start and lo < end]
    kept = cutter.plan_cuts(start, end, words, segment["peaks"], cut_settings,
                            quiet + [(lo, hi) for lo, hi, _ in fillers])
    # 確認用に、消した所ごとに理由を添える
    reasons = ([(lo, hi, "単語の間") for lo, hi in
                cutter.silence_ranges(start, end, words, cut_settings)]
               + [(lo, hi, "無音（auto-editor）") for lo, hi in quiet]
               + [(lo, hi, "フィラー: %s" % text) for lo, hi, text in fillers])
    removed = [{"start": round(lo, 3), "end": round(hi, 3),
                "reasons": sorted(set(_overlaps(lo, hi, reasons))) or ["短い断片の整理"]}
               for lo, hi in cutter._subtract([(start, end)], kept)]
    return kept, removed


def build_plan(segments, words_by_segment, silence, settings, fps):
    """区間ごとに残す部分を決め、カット後のタイムラインとテロップを作る。"""
    cuts, removed, all_words = [], [], []
    for segment, words in zip(segments, words_by_segment):
        kept, gone = plan_segment(segment, words, silence, settings)
        cuts.append(kept)
        removed.extend(gone)
        all_words.extend(words)
    timeline = editexport.Timeline(segments, cuts, fps)
    cues = make_cues(all_words, timeline, settings.caption)
    return EditPlan(segments, all_words, cuts, removed, timeline, cues)


def analyze_range(media_path, start, end, settings, report=None, use_cache=True):
    """動画の start〜end 秒（素材の時刻）で、残す部分と消す所を決める。

    Premiere のパネルから、選んだクリップの素材と使用範囲を受け取って呼ぶ。
    report(文言, 割合) で進み具合を伝える。文字起こしはキャッシュを使うので、
    同じ素材の同じ範囲なら2回目はすぐ返る。
    """
    report = report or (lambda message, frac=None: None)
    if not os.path.isfile(media_path):
        raise EditError("素材ファイルが見つかりません: %s" % media_path)
    if end - start < 0.1:
        raise EditError("範囲が短すぎます（%.2f 秒）。" % (end - start))
    deps.require(settings.engine, settings.silence_method)
    length = end - start

    def advance(sec, _total):
        report("文字起こし中 %s / %s" % (_clock(sec), _clock(length)), 0.05 + 0.85 * sec / length)

    def note(message):
        report(message, None)

    report("文字起こしの準備中…", 0.02)
    began = time.time()
    result = transcribe.transcribe_segment(
        media_path, start, end, settings.engine, report=note, use_cache=use_cache,
        prompt=settings.prompt, audio_track=settings.audio_track or None, progress=advance)
    cached = time.time() - began < 1.0 and result.elapsed > 1.0
    silence = []
    if settings.silence_method == "audio":
        report("auto-editor で音量を判定しています…", 0.9)
        from .realtest import probe
        silence = audio_silence(media_path, probe(media_path)[0], settings)
    report("カットする所を決めています…", 0.95)
    kept, removed = plan_segment({"start": start, "end": end, "peaks": []}, result.words,
                                 silence, settings)
    return {"keep": [[round(lo, 4), round(hi, 4)] for lo, hi in kept],
            "removed": removed,
            "words": len(result.words),
            "engine": settings.engine, "device": result.device, "cached": cached,
            "transcribe_sec": round(result.elapsed, 1),
            "original_sec": round(length, 3),
            "kept_sec": round(sum(hi - lo for lo, hi in kept), 3)}


class PlacedTimeline:
    """タイムラインに置かれたクリップの並び（シーケンスの秒）。テロップの時刻合わせに使う。

    subtitles.make_cues は editexport.Timeline と同じ形（fps・clips・locate・to_sequence）を
    求める。ここでは語の時刻をあらかじめシーケンスの秒に直しておくので、素材の位置と
    シーケンスの位置は同じ値になる。クリップの境目ではテロップを分ける（カット点をまたがない）。
    """

    def __init__(self, ranges, fps):
        self.fps = fps
        self.clips = []
        for lo, hi in sorted(ranges):
            a, b = editexport._frames(lo, fps), editexport._frames(hi, fps)
            if b > a:
                self.clips.append({"src_in": a, "src_out": b, "seq_start": a})

    def locate(self, seconds):
        frame = editexport._frames(seconds, self.fps)
        for index, clip in enumerate(self.clips):
            if clip["src_in"] <= frame < clip["src_out"]:
                return index
        return None

    def to_sequence(self, seconds, clip_index=None):
        frame = editexport._frames(seconds, self.fps)
        if clip_index is None:
            clip_index = self.locate(seconds)
            if clip_index is None:
                return None
        clip = self.clips[clip_index]
        return min(max(frame, clip["src_in"]), clip["src_out"])


def _clip_words(clips, settings, report, use_cache=True):
    """クリップごとの語（素材の秒）。保存済みの文字起こしを優先し、無い範囲だけ起こす。"""
    by_media = {}
    for index, clip in enumerate(clips):
        by_media.setdefault(clip["media_path"], []).append(index)
    words = [None] * len(clips)
    done = 0
    for media, indexes in by_media.items():
        if not os.path.isfile(media):
            raise EditError("素材ファイルが見つかりません: %s" % media)
        lo = min(clips[i]["source_in"] for i in indexes)
        hi = max(clips[i]["source_out"] for i in indexes)
        # カットの判定のときに起こした広い範囲が残っていれば、それを切り分けて使う
        whole = transcribe.find_covering(media, lo, hi, settings.engine) if use_cache else None
        for i in indexes:
            clip = clips[i]
            if whole is not None:
                found = whole.words
            else:
                found = transcribe.find_covering(media, clip["source_in"], clip["source_out"],
                                                 settings.engine) if use_cache else None
                if found is None:
                    report("文字起こし中（クリップ %d/%d）…" % (done + 1, len(clips)), None)
                    found = transcribe.transcribe_segment(
                        media, clip["source_in"], clip["source_out"], settings.engine,
                        use_cache=use_cache, prompt=settings.prompt,
                        audio_track=settings.audio_track or None)
                found = found.words
            words[i] = [w for w in found
                        if clip["source_in"] <= (w.start + w.end) / 2.0 < clip["source_out"]]
            done += 1
            report("テロップを組み立てています（%d/%d）" % (done, len(clips)), 0.9 * done / len(clips))
    return words


def captions_for_clips(clips, fps, settings, report=None, use_cache=True):
    """タイムラインのクリップの並びから、シーケンスの時刻のテロップ（SRT の文字列と枚数）を作る。

    clips は [{"media_path", "source_in", "source_out", "seq_start"}]（秒）。
    Premiere のパネルが、カットした後（またはカット前）のタイムラインから送る。
    """
    report = report or (lambda message, frac=None: None)
    if not clips:
        raise EditError("テロップを付けるクリップがありません。")
    words_by_clip = _clip_words(clips, settings, report, use_cache)
    placed, ranges = [], []
    for clip, words in zip(clips, words_by_clip):
        shift = clip["seq_start"] - clip["source_in"]
        ranges.append((clip["seq_start"], clip["seq_start"] + clip["source_out"] - clip["source_in"]))
        placed.extend(transcribe.Word(w.text, w.start + shift, w.end + shift) for w in words)
    placed.sort(key=lambda w: w.start)
    cues = make_cues(placed, PlacedTimeline(ranges, fps), settings.caption)
    return to_srt(cues), len(cues)


def unique_path(path):
    """同じ名前のファイルがあれば「名前 (2).srt」のようにずらす（上書きしない）。"""
    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(path)
    number = 2
    while os.path.exists("%s (%d)%s" % (stem, number, ext)):
        number += 1
    return "%s (%d)%s" % (stem, number, ext)


# ---------------------------------------------------------------- 進み具合

class Progress:
    """「[ 35%] 文字起こし中…」の形で、1%進むか文言が変わったときだけ出す。"""

    def __init__(self, stream=None):
        self.stream = stream or sys.stdout
        self.last = None
        self.frac = 0.0

    def __call__(self, frac, message):
        if frac is not None:
            self.frac = max(self.frac, min(1.0, frac))
        key = (int(self.frac * 100), message)
        if key == self.last:
            return
        self.last = key
        print("[%3d%%] %s" % key, file=self.stream, flush=True)


def _clock(seconds):
    seconds = int(seconds)
    return "%d:%02d:%02d" % (seconds // 3600, seconds % 3600 // 60, seconds % 60) \
        if seconds >= 3600 else "%d:%02d" % (seconds // 60, seconds % 60)


# ---------------------------------------------------------------- 書き出し

def output_names(media_path, out_dir=None):
    stem = os.path.splitext(os.path.basename(media_path))[0]
    folder = out_dir or os.path.join(os.path.dirname(os.path.abspath(media_path)), stem)
    return folder, stem


def write_outputs(plan, media_path, info, settings, folder, stem, extra):
    os.makedirs(folder, exist_ok=True)
    fps = plan.timeline.fps
    xml = editexport.to_fcp7_xml(info, plan.segments, fps, media_path, info.width, info.height,
                                 plan.cuts)
    files = {"xml": os.path.join(folder, stem + ".xml"),
             "srt": os.path.join(folder, stem + ".srt"),
             "json": os.path.join(folder, stem + "_transcript.json")}
    with open(files["xml"], "w", encoding="utf-8", newline="\n") as fh:
        fh.write(xml)
    # Premiere は BOM なしの UTF-8 の SRT を読める
    with open(files["srt"], "w", encoding="utf-8", newline="\n") as fh:
        fh.write(to_srt(plan.cues))
    payload = dict(extra)
    payload.update({
        "input": os.path.abspath(media_path),
        "settings": settings.as_dict(),
        "fps": fps,
        "segments": [{"start": round(s["start"], 3), "end": round(s["end"], 3),
                      "peaks": s["peaks"],
                      "kept": [[round(lo, 3), round(hi, 3)] for lo, hi in kept]}
                     for s, kept in zip(plan.segments, plan.cuts)],
        "removed": plan.removed,
        "text": transcribe.join_words(w.text for w in plan.words),
        "words": [w.to_list() for w in plan.words],
        "cues": len(plan.cues),
        "effects": plan.effects,
    })
    with open(files["json"], "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=1)
    return files


# ---------------------------------------------------------------- 全体

def run(media_path, settings, segments_from=None, out_dir=None, fps=None, use_cache=True,
        progress=None):
    """全体を通して行い、書き出したファイルと集計を返す。

    segments_from は区間の決め方: None（動画全体）/ ("file", パス) /
    ("highlights", URL, 件数, 前後の秒)。
    """
    progress = progress or Progress()
    began_all = time.time()
    if not os.path.isfile(media_path):
        raise EditError("動画ファイルが見つかりません: %s" % media_path)
    deps.require(settings.engine, settings.silence_method)

    from .realtest import probe

    progress(0.0, "動画を調べています…")
    duration, found_fps, width, height = probe(media_path)
    if duration <= 0:
        raise EditError("動画の長さを読めませんでした: %s" % media_path)
    fps = float(fps or found_fps or editexport.DEFAULT_FPS)

    if segments_from is None:
        segments = whole_file(duration)
    elif segments_from[0] == "file":
        segments = read_segments_file(segments_from[1], duration)
    else:
        segments = highlight_segments(segments_from[1], segments_from[2], segments_from[3], duration)
    total = sum(s["end"] - s["start"] for s in segments)
    progress(0.02, "区間 %d 件、計 %s を処理します（%s・%s fps）"
             % (len(segments), _clock(total), settings.engine, round(fps, 3)))

    # 文字起こしが全体の大半（2〜90%）。区間の長さで按分する
    words_by_segment = []
    done_sec = 0.0
    transcribe_sec = 0.0
    device = ""
    cached = 0
    for index, segment in enumerate(segments):
        start, end = segment["start"], segment["end"]
        head = "文字起こし中（区間 %d/%d）" % (index + 1, len(segments))

        def advance(sec, _total, base=done_sec, head=head):
            progress(0.02 + 0.88 * (base + sec) / total,
                     "%s %s / %s" % (head, _clock(base + sec), _clock(total)))

        def note(message, head=head):
            progress(None, message if not message.startswith("（") else head + message)

        progress(0.02 + 0.88 * done_sec / total, head + "…")
        began = time.time()
        result = transcribe.transcribe_segment(
            media_path, start, end, settings.engine, report=note, use_cache=use_cache,
            prompt=settings.prompt, audio_track=settings.audio_track or None, progress=advance)
        if time.time() - began < 1.0 and result.elapsed > 1.0:
            cached += 1
        transcribe_sec += result.elapsed
        device = device or result.device
        words_by_segment.append(result.words)
        done_sec += end - start
    progress(0.90, "文字起こし完了（%d 語%s）" % (sum(len(w) for w in words_by_segment),
                                          "・%d 区間はキャッシュを使用" % cached if cached else ""))

    silence = []
    if settings.silence_method == "audio":
        progress(0.90, "auto-editor で音量を判定しています（音声トラック %d）…" % settings.audio_track)
        silence = audio_silence(media_path, duration, settings)

    progress(0.95, "カット点とテロップを決めています…")
    plan = build_plan(segments, words_by_segment, silence, settings, fps)

    folder, stem = output_names(media_path, out_dir)
    info = StreamInfo("local", stem, os.path.abspath(media_path), title=stem,
                      duration=duration, fps=fps, width=width, height=height)
    edited = plan.timeline.total_frames / fps
    summary = {"engine": settings.engine, "device": device,
               "transcribe_sec": round(transcribe_sec, 1),
               "original_sec": round(total, 1), "edited_sec": round(edited, 1),
               "clips": len(plan.timeline.clips), "cues": len(plan.cues),
               "fillers": sum(1 for r in plan.removed
                              if any(x.startswith("フィラー") for x in r["reasons"]))}
    files = write_outputs(plan, media_path, info, settings, folder, stem, {"summary": summary})
    summary["total_sec"] = round(time.time() - began_all, 1)
    progress(1.0, "書き出しました: %s" % folder)
    return files, summary


# ---------------------------------------------------------------- CLI

# 引数の名前 → Settings の項目名
OVERRIDES = ("engine", "silence_method", "audio_threshold", "audio_margin_sec", "audio_track",
             "gap_cut_sec", "gap_cut", "pad_sec", "min_piece_sec", "protect_sec", "fillers",
             "filler_prompt", "line_chars", "max_lines", "min_show_sec", "max_show_sec",
             "strip_trailing_punct")


def make_parser():
    p = argparse.ArgumentParser(
        prog="python -m stream_highlight.autoedit",
        description="手元の動画から、Premiere 用のカット済みシーケンス（XML）とテロップ（SRT）を作ります。")
    p.add_argument("input", nargs="?", help="動画ファイル（mp4 など）")
    p.add_argument("--profile", default=DEFAULT_PROFILE,
                   help="config/profiles/ のプロファイル名か yaml のパス（既定: default）")
    p.add_argument("--out", help="書き出し先のフォルダ（既定: 動画と同じ場所の「<名前>」フォルダ）")
    p.add_argument("--check", action="store_true", help="必要なものが揃っているかだけを確かめる")
    p.add_argument("--list-profiles", action="store_true", help="使えるプロファイルを一覧する")
    p.add_argument("--no-cache", action="store_true", help="保存済みの文字起こしを使わずにやり直す")
    p.add_argument("--fps", type=float, help="シーケンスのフレームレート（既定: 動画から読む）")

    g = p.add_argument_group("区間（指定しなければ動画全体）")
    g.add_argument("--segments", help="区間を書いた JSON（[[開始秒, 終了秒], ...]）")
    g.add_argument("--highlights", metavar="URL",
                   help="ハイライト抽出の見せ場を使う。配信の URL（動画は配信全体であること）")
    g.add_argument("--top", type=int, default=DEFAULT_TOP, help="見せ場の上位何件か（既定: 5）")
    g.add_argument("--margin", type=float, default=DEFAULT_MARGIN_SEC,
                   help="見せ場の前後に取る秒数（既定: 60）")

    o = p.add_argument_group("プロファイルの値の上書き")
    bool_flag = argparse.BooleanOptionalAction
    o.add_argument("--engine", choices=ENGINES)
    o.add_argument("--silence-method", dest="silence_method", choices=SILENCE_METHODS)
    o.add_argument("--silence-threshold", dest="audio_threshold", type=float,
                   help="auto-editor の音量のしきい値（silence_method: audio のとき）")
    o.add_argument("--audio-margin", dest="audio_margin_sec", type=float)
    o.add_argument("--audio-track", dest="audio_track", type=int,
                   help="判定と文字起こしに使う音声トラック（0 = 1本目）")
    o.add_argument("--gap-cut-sec", dest="gap_cut_sec", type=float)
    o.add_argument("--gap-cut", dest="gap_cut", action=bool_flag)
    o.add_argument("--pad-sec", dest="pad_sec", type=float)
    o.add_argument("--min-piece-sec", dest="min_piece_sec", type=float)
    o.add_argument("--protect-sec", dest="protect_sec", type=float)
    o.add_argument("--fillers", dest="fillers", action=bool_flag)
    o.add_argument("--filler-prompt", dest="filler_prompt", action=bool_flag)
    o.add_argument("--line-chars", dest="line_chars", type=int)
    o.add_argument("--max-lines", dest="max_lines", type=int)
    o.add_argument("--min-show", dest="min_show_sec", type=float)
    o.add_argument("--max-show", dest="max_show_sec", type=float)
    o.add_argument("--strip-punct", dest="strip_trailing_punct", action=bool_flag)
    return p


def settings_from_args(args):
    flat = flatten_profile(load_profile(args.profile))
    flat.update({key: getattr(args, key) for key in OVERRIDES if getattr(args, key) is not None})
    return build_settings(flat)


def main(argv=None):
    args = make_parser().parse_args(argv)
    if args.check:
        print("\n".join(deps.report()))
        return 0
    if args.list_profiles:
        for name in available_profiles():
            print(name)
        return 0
    if not args.input:
        make_parser().print_usage()
        print("動画ファイルを指定してください。", file=sys.stderr)
        return 2
    try:
        settings = settings_from_args(args)
        segments_from = None
        if args.segments:
            segments_from = ("file", args.segments)
        elif args.highlights:
            segments_from = ("highlights", args.highlights, args.top, args.margin)
        files, summary = run(args.input, settings, segments_from, args.out, args.fps,
                             use_cache=not args.no_cache)
    except (EditError, deps.MissingDependency, transcribe.TranscribeError, FetchError) as exc:
        print("エラー: %s" % exc, file=sys.stderr)
        return 1
    print("")
    print("プロファイル: %s / エンジン: %s（%s）" % (settings.profile, summary["engine"],
                                             summary["device"] or "キャッシュ"))
    print("長さ: %s → %s（%d カット、テロップ %d 枚、フィラー %d 箇所）"
          % (_clock(summary["original_sec"]), _clock(summary["edited_sec"]),
             summary["clips"], summary["cues"], summary["fillers"]))
    print("時間: 文字起こし %.1f 秒 / 全体 %.1f 秒" % (summary["transcribe_sec"], summary["total_sec"]))
    for key in ("xml", "srt", "json"):
        print("  %s" % files[key])
    return 0
