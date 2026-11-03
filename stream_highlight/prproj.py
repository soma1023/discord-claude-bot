# -*- coding: utf-8 -*-
"""Premiere のプロジェクトファイル（.prproj）から、グラフィックのテロップの文字を読む。

Premiere 26 の UXP は「ソーステキスト」の値をパネルに渡さない（null が返る）ので、
テロップの行数はこちらで保存済みのプロジェクトファイルから読む。
.prproj は gzip で圧縮した XML。オブジェクトは ObjectID / ObjectRef でつながっている:
VideoClipTrackItem → VideoComponentChain → 部品（VideoFilterComponent）→ パラメーター。
ソーステキストの値は base64 のバイナリ（FlatBuffers）で、その中に UTF-8 の文字がそのまま入っている。
"""

import base64
import functools
import glob
import math
import os
import re
import struct
import gzip

TICKS_PER_SECOND = 254016000000
LINE_BREAKS = re.compile("\r\n|\r|\n| | |\x03")
# バイナリの中の、文字として読める並び（UTF-8 の2〜3バイト文字と ASCII）
_TEXT_RUN = re.compile(rb"(?:[\xe0-\xef][\x80-\xbf]{2}|[\xc2-\xdf][\x80-\xbf]|[\x20-\x7e\r\n])+")
# 同じ値は1回だけ中身を書き、ほかは BinaryHash だけの空要素になる
_BINARY = re.compile(r'BinaryHash="([^"]+)">([A-Za-z0-9+/=\s]+)<')
_OBJECT = re.compile(r'<(\w+) ObjectID="(\d+)"[^>]*>(.*?)</\1>', re.S)
TEXT_PARAM = "<Name>ソーステキスト</Name>"
POSITION_PARAM = "<Name>位置</Name>"


def read_xml(path):
    with open(path, "rb") as fh:
        data = fh.read()
    if data[:2] == b"\x1f\x8b":
        data = gzip.decompress(data)
    return data.decode("utf-8")


def text_of(blob):
    """ソーステキストの値（バイナリ）から、テロップの文字を取り出す。一番長い、非 ASCII を含む並び。"""
    runs = [r.decode("utf-8", "ignore") for r in _TEXT_RUN.findall(blob) if re.search(rb"[\xc2-\xef]", r)]
    return max(runs, key=len) if runs else None


def line_count(text):
    return max(1, len([s for s in LINE_BREAKS.split(text or "") if s.strip()]))


# ---- 画面に見える行数（折り返しと、枠からはみ出して見えない行を考える） ----
# ソーステキストの値のうち、文字より前の決まった場所に、テキストの枠の高さと幅（px）が入っている。
# 文字より後ろの最後の整数値がフォントサイズ。フォントは PostScript 名（例: NewRodinPro-EB）で入っている。
BOX_OFFSET = 192
LINE_SPACING = 1.2            # 行送り（フォントサイズに対する倍率。Premiere の自動の行送り）
# 枠の幅をフォントサイズの半分ほど超えるまでは折り返さなかった（実機で 1495px は1行、1539px は折り返し）
WRAP_ALLOWANCE = 0.5
FONT_DIRS = [os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts"),
             os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")]


def text_style(blob, text):
    """(枠の幅, 枠の高さ, フォントサイズ, PostScript 名)。読めないものは None。"""
    raw = text.encode("utf-8")
    end = blob.find(raw)
    if end < 0:
        return None, None, None, None
    end += len(raw)
    width = height = None
    if len(blob) >= BOX_OFFSET + 8:
        height, width = struct.unpack_from("<ff", blob, BOX_OFFSET)
        if not (100 <= width <= 8000 and 20 <= height <= 4000):
            width = height = None
    size = None
    for i in range(end + (-end) % 4, len(blob) - 3, 4):
        value = struct.unpack_from("<f", blob, i)[0]
        if 4 <= value <= 1000 and abs(value - round(value)) < 1e-6:
            size = value
    names = re.findall(rb"[A-Za-z0-9]+-[A-Za-z0-9]+", blob[:end - len(raw)])
    return width, height, size, names[-1].decode("ascii") if names else None


@functools.lru_cache(maxsize=16)
def _font(ps_name, size):
    try:
        from PIL import ImageFont
    except ImportError:
        return None
    for folder in FONT_DIRS:
        for path in glob.glob(os.path.join(folder, "*")):
            stem = os.path.splitext(os.path.basename(path))[0]
            if stem == ps_name or stem.endswith("-" + ps_name) or stem.endswith(ps_name):
                try:
                    return ImageFont.truetype(path, size)
                except OSError:
                    continue
    return None


def visible_lines(text, width, height, size, ps_name):
    """(見える行数, 折り返し後の行数)。フォントや枠が分からなければ改行の数だけで数える。"""
    explicit = line_count(text)
    font = _font(ps_name, int(round(size))) if (ps_name and size) else None
    if not (font and width and height):
        return explicit, explicit
    limit = width + size * WRAP_ALLOWANCE
    wrapped = sum(max(1, math.ceil(font.getlength(line) / limit))
                  for line in LINE_BREAKS.split(text) if line.strip()) or 1
    fits = max(1, int(height // (size * LINE_SPACING)))
    return min(wrapped, fits), wrapped


def captions(xml):
    """テキストを1つ持つグラフィックのクリップごとに {start, end, text, lines, position} を返す。

    start / end はシーケンス上の ticks（UXP の TickTime.ticks と同じ単位）。
    テキストが2つ以上あるグラフィックや、文字を読めないものは入れない。
    """
    binaries = {m.group(1): m.group(2) for m in _BINARY.finditer(xml)}
    objects = {m.group(2): (m.group(1), m.group(3)) for m in _OBJECT.finditer(xml)}
    found = []
    for tag, body in objects.values():
        if tag != "VideoClipTrackItem":
            continue
        chain_ref = re.search(r'<Components ObjectRef="(\d+)"', body)
        start = re.search(r"<Start>(-?\d+)</Start>", body)
        end = re.search(r"<End>(-?\d+)</End>", body)
        if not (chain_ref and start and end):
            continue
        layers = []
        chain = objects.get(chain_ref.group(1), ("", ""))[1]
        for comp_ref in re.findall(r'<Component Index="\d+" ObjectRef="(\d+)"', chain):
            component = objects.get(comp_ref, ("", ""))[1]
            text = None
            for param_ref in re.findall(r'<Param Index="\d+" ObjectRef="(\d+)"', component):
                param = objects.get(param_ref, ("", ""))[1]
                if TEXT_PARAM in param:
                    blob = base64.b64decode(_text_value(param, binaries))
                    text = text_of(blob)
                    layers.append({"text": text, "position": None,
                                   "style": text_style(blob, text) if text else (None,) * 4})
                elif text is not None and POSITION_PARAM in param:
                    point = re.search(r"<StartKeyframe>[^,]*,([-\d.e]+):([-\d.e]+)", param)
                    if point:
                        layers[-1]["position"] = [float(point.group(1)), float(point.group(2))]
                    text = None
        if len(layers) != 1 or layers[0]["text"] is None:
            continue
        found.append({
            "start": int(start.group(1)), "end": int(end.group(1)),
            "text": layers[0]["text"], "style": layers[0]["style"],
            "position": layers[0]["position"],
        })
    # フォント名が読めなかったものは、一番多いフォントとみなす
    names = [c["style"][3] for c in found if c["style"][3]]
    common = max(set(names), key=names.count) if names else None
    for c in found:
        width, height, size, name = c.pop("style")
        c["explicit_lines"] = line_count(c["text"])
        c["lines"], c["wrapped_lines"] = visible_lines(c["text"], width, height, size, name or common)
    return found


def _text_value(param, binaries):
    value = re.search(r"<StartKeyframeValue([^>]*?)(/>|>([^<]*)<)", param)
    if not value:
        return ""
    if value.group(3) and value.group(3).strip():
        return value.group(3)
    digest = re.search(r'BinaryHash="([^"]+)"', value.group(1))
    return binaries.get(digest.group(1), "") if digest else ""


def caption_lines(path):
    return captions(read_xml(path))
