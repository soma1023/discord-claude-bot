# -*- coding: utf-8 -*-
"""見せ場の候補を、編集ソフトに読み込める形で書き出す。

長い配信をそのまま編集ソフトに乗せると、結局スクラブして探すことになる。
候補の前後だけを切り出したシーケンスを作っておけば、編集はその中だけで済む。

Premiere は FCP7 XML（.xml）と EDL を読み込める。XML は動画ファイルの場所と
解像度まで書けるので自動でつながるが、形式が厳密。EDL は素材を手動でつなぐ
代わりにどの編集ソフトでも通る。両方出せるようにしてある。

解像度を書き忘れると、Premiere はシーケンスを自前の既定値で作ってしまい、
1080pの配信を読み込んだのに別の解像度のシーケンスができる。
"""

import posixpath
import urllib.parse
import xml.etree.ElementTree as ET
from xml.dom import minidom

DEFAULT_MARGIN_SEC = 300.0      # ピークの前後5分
DEFAULT_FPS = 30.0
DEFAULT_WIDTH = 1920
DEFAULT_HEIGHT = 1080


def _file_name(media_path, fallback):
    """パスからファイル名を取り出す。Windowsの区切りでも動くようにする。"""
    if not media_path:
        return fallback
    return posixpath.basename(media_path.replace("\\", "/")) or fallback


def _frames(seconds, fps):
    return int(round(max(0.0, seconds) * fps))


def timecode(seconds, fps):
    """HH:MM:SS:FF 形式にする。"""
    total = _frames(seconds, fps)
    rate = int(round(fps))
    frames = total % rate
    total //= rate
    return "%02d:%02d:%02d:%02d" % (total // 3600, total % 3600 // 60,
                                    total % 60, frames)


def build_segments(peaks, margin_sec=DEFAULT_MARGIN_SEC, duration=0.0):
    """各ピークの前後に余白を付け、重なったものをまとめる。

    まとめないと、近いピークどうしで同じ映像が何度も並んでしまう。
    戻り値は開始順の [{"start", "end", "peaks"}]。
    """
    spans = []
    for peak in sorted(float(p) for p in peaks):
        start = max(0.0, peak - margin_sec)
        end = peak + margin_sec
        if duration:
            end = min(end, float(duration))
        if end <= start:
            continue
        if spans and start <= spans[-1]["end"]:
            spans[-1]["end"] = max(spans[-1]["end"], end)
            spans[-1]["peaks"].append(peak)
        else:
            spans.append({"start": start, "end": end, "peaks": [peak]})
    return spans


def _path_url(media_path):
    """Premiere が読む file:// 形式にする。"""
    if not media_path:
        return ""
    path = media_path.replace("\\", "/")
    if len(path) > 1 and path[1] == ":":        # C:/... のとき
        path = "/" + path
    # ドライブレターのコロンは残す。Premiere は file://localhost/C:/... を期待する
    return "file://localhost" + urllib.parse.quote(path, safe="/:")


def _rate(parent, fps):
    rate = ET.SubElement(parent, "rate")
    ET.SubElement(rate, "timebase").text = str(int(round(fps)))
    # 29.97 や 59.94 は「整数timebase + NTSC」で表す
    ET.SubElement(rate, "ntsc").text = "TRUE" if abs(fps - round(fps)) > 0.001 else "FALSE"
    return rate


def _sample_characteristics(parent, fps, width, height):
    """映像の素性（解像度・フレームレート）を書く。

    これを省くと、Premiere はシーケンスの解像度を自前の既定値で作ってしまう。
    1080pの配信を読み込んでも別の解像度になり、書き出した動画がおかしくなる。
    """
    sc = ET.SubElement(parent, "samplecharacteristics")
    _rate(sc, fps)
    ET.SubElement(sc, "width").text = str(int(width))
    ET.SubElement(sc, "height").text = str(int(height))
    ET.SubElement(sc, "anamorphic").text = "FALSE"
    ET.SubElement(sc, "pixelaspectratio").text = "square"
    ET.SubElement(sc, "fielddominance").text = "none"
    return sc


def to_fcp7_xml(info, segments, fps=DEFAULT_FPS, media_path="",
                width=0, height=0):
    """Premiere が読み込める FCP7 XML を組み立てる。"""
    width = int(width or info.width or DEFAULT_WIDTH)
    height = int(height or info.height or DEFAULT_HEIGHT)
    name = (info.title or info.video_id or "highlights").strip()
    file_name = _file_name(media_path, "%s.mp4" % info.video_id)
    source_frames = _frames(info.duration or 0, fps) or 1
    total_frames = sum(_frames(s["end"] - s["start"], fps) for s in segments)

    root = ET.Element("xmeml", {"version": "4"})
    sequence = ET.SubElement(root, "sequence")
    ET.SubElement(sequence, "name").text = "%s 見せ場" % name
    ET.SubElement(sequence, "duration").text = str(total_frames)
    _rate(sequence, fps)
    media = ET.SubElement(sequence, "media")

    video = ET.SubElement(media, "video")
    # シーケンスの解像度は format で決まる。track より前に置く必要がある。
    _sample_characteristics(ET.SubElement(video, "format"), fps, width, height)
    video_track = ET.SubElement(video, "track")

    audio = ET.SubElement(media, "audio")
    audio_tracks = [ET.SubElement(audio, "track") for _ in range(2)]

    first_file = True
    position = 0
    for index, segment in enumerate(segments, 1):
        length = _frames(segment["end"] - segment["start"], fps)
        src_in = _frames(segment["start"], fps)
        label = "%02d %s" % (index, timecode(segment["start"], fps))

        for kind, track, channel in ([("video", video_track, 0)]
                                     + [("audio", audio_tracks[i], i + 1) for i in range(2)]):
            item = ET.SubElement(track, "clipitem",
                                 {"id": "%s-%d" % (kind if kind == "video" else
                                                   "audio%d" % channel, index)})
            ET.SubElement(item, "name").text = label
            ET.SubElement(item, "duration").text = str(source_frames)
            _rate(item, fps)
            ET.SubElement(item, "start").text = str(position)
            ET.SubElement(item, "end").text = str(position + length)
            ET.SubElement(item, "in").text = str(src_in)
            ET.SubElement(item, "out").text = str(src_in + length)

            if first_file:
                file_el = ET.SubElement(item, "file", {"id": "media-1"})
                ET.SubElement(file_el, "name").text = file_name
                if media_path:
                    ET.SubElement(file_el, "pathurl").text = _path_url(media_path)
                _rate(file_el, fps)
                ET.SubElement(file_el, "duration").text = str(source_frames)
                file_media = ET.SubElement(file_el, "media")
                file_video = ET.SubElement(file_media, "video")
                ET.SubElement(file_video, "duration").text = str(source_frames)
                _sample_characteristics(file_video, fps, width, height)
                file_audio = ET.SubElement(file_media, "audio")
                ET.SubElement(file_audio, "channelcount").text = "2"
                first_file = False
            else:
                ET.SubElement(item, "file", {"id": "media-1"})

            if kind == "audio":
                source = ET.SubElement(item, "sourcetrack")
                ET.SubElement(source, "mediatype").text = "audio"
                ET.SubElement(source, "trackindex").text = str(channel)

        # 区間のどこが候補の時刻なのかを、マーカーで残す
        for peak in segment["peaks"]:
            marker = ET.SubElement(sequence, "marker")
            ET.SubElement(marker, "name").text = "ピーク %s" % timecode(peak, fps)
            ET.SubElement(marker, "in").text = str(
                position + _frames(peak - segment["start"], fps))
            ET.SubElement(marker, "out").text = "-1"

        position += length

    body = minidom.parseString(ET.tostring(root, encoding="utf-8")).toprettyxml(indent="  ")
    body = body.split("\n", 1)[1]        # minidom の宣言を自前のものに置き換える
    return ('<?xml version="1.0" encoding="UTF-8"?>\n'
            '<!DOCTYPE xmeml>\n' + body)


def to_edl(info, segments, fps=DEFAULT_FPS, media_path=""):
    """どの編集ソフトでも読める EDL を組み立てる。"""
    name = (info.title or info.video_id or "highlights").strip()
    file_name = _file_name(media_path, "%s.mp4" % info.video_id)

    lines = ["TITLE: %s 見せ場" % name, "FCM: NON-DROP FRAME", ""]
    position = 0.0
    for index, segment in enumerate(segments, 1):
        length = segment["end"] - segment["start"]
        src_in = timecode(segment["start"], fps)
        src_out = timecode(segment["end"], fps)
        rec_in = timecode(position, fps)
        rec_out = timecode(position + length, fps)
        for channel in ("V", "AA"):
            lines.append("%03d  AX       %-5s C        %s %s %s %s"
                         % (index, channel, src_in, src_out, rec_in, rec_out))
        lines.append("* FROM CLIP NAME: %s" % file_name)
        lines.append("* 候補: %s" % " / ".join(timecode(p, fps) for p in segment["peaks"]))
        lines.append("")
        position += length
    return "\n".join(lines)


def export(info, peaks, fmt="xml", margin_sec=DEFAULT_MARGIN_SEC,
           fps=DEFAULT_FPS, media_path="", width=0, height=0):
    """書き出し一式。(ファイル名, 中身, 区間の数) を返す。"""
    fps = float(fps) or DEFAULT_FPS
    segments = build_segments(peaks, margin_sec, info.duration)
    if not segments:
        raise ValueError("書き出す候補がありません。")

    safe = "".join(c for c in (info.title or info.video_id)
                   if c not in '\\/:*?"<>|').strip()[:50] or "highlights"
    if fmt == "edl":
        return "%s.edl" % safe, to_edl(info, segments, fps, media_path), len(segments)
    return ("%s.xml" % safe,
            to_fcp7_xml(info, segments, fps, media_path, width, height),
            len(segments))
