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


class Timeline:
    """自動カットした後の並び。元の配信の位置とシーケンス上の位置を対応づける。

    カットすると1区間が複数のクリップになる。フレームへの丸めをここ1か所で
    行い、XML・EDL・テロップ（SRT）の全てがこの値を使う。別々に丸めると、
    1フレームずつずれて隙間や重なりができ、テロップも映像とずれる。

    clips は [{"segment", "piece", "src_in", "src_out", "seq_start"}]（フレーム）。
    隣り合うクリップの seq は必ず隙間なく続く。
    """

    def __init__(self, segments, cuts, fps):
        self.fps = fps
        self.clips = []
        position = 0
        for seg_index, segment in enumerate(segments):
            if not cuts or not cuts[seg_index]:
                # カットしない区間は、従来の書き出しと同じ丸め（長さを先に丸める）にする。
                # テロップだけを付けるときに、従来どおりの XML とずれないように。
                src_in = _frames(segment["start"], fps)
                length = _frames(segment["end"] - segment["start"], fps)
                self.clips.append({"segment": seg_index, "piece": 0, "src_in": src_in,
                                   "src_out": src_in + length, "seq_start": position})
                position += length
                continue
            for piece_index, (lo, hi) in enumerate(cuts[seg_index]):
                src_in, src_out = _frames(lo, fps), _frames(hi, fps)
                if src_out <= src_in:
                    continue
                self.clips.append({"segment": seg_index, "piece": piece_index,
                                   "src_in": src_in, "src_out": src_out,
                                   "seq_start": position})
                position += src_out - src_in
        self.total_frames = position

    def locate(self, seconds):
        """元の配信の秒が、どのクリップに入るか。カットされた所なら None。"""
        frame = _frames(seconds, self.fps)
        for index, clip in enumerate(self.clips):
            if clip["src_in"] <= frame < clip["src_out"]:
                return index
        return None

    def to_sequence(self, seconds, clip_index=None):
        """元の配信の秒を、シーケンス上のフレームにする。カットされた所なら None。

        clip_index を渡すと、そのクリップの範囲に収めて返す（テロップの端が
        カット点をまたぐときに使う）。
        """
        frame = _frames(seconds, self.fps)
        if clip_index is None:
            clip_index = self.locate(seconds)
            if clip_index is None:
                return None
        clip = self.clips[clip_index]
        frame = min(max(frame, clip["src_in"]), clip["src_out"])
        return clip["seq_start"] + frame - clip["src_in"]

    def nearest_sequence(self, seconds, segment):
        """カットされた所でも、同じ区間の一番近いクリップに寄せて位置を返す。"""
        found = self.to_sequence(seconds)
        if found is not None:
            return found
        frame = _frames(seconds, self.fps)
        best = None
        for index, clip in enumerate(self.clips):
            if clip["segment"] != segment:
                continue
            dist = min(abs(frame - clip["src_in"]), abs(frame - clip["src_out"]))
            if best is None or dist < best[0]:
                best = (dist, index)
        if best is None:
            return 0
        return self.to_sequence(seconds, best[1])


def to_fcp7_xml(info, segments, fps=DEFAULT_FPS, media_path="",
                width=0, height=0, cuts=None):
    """Premiere が読み込める FCP7 XML を組み立てる。

    cuts は区間ごとの「残す部分」[(開始秒, 終了秒)]。None なら区間を丸ごと
    1クリップにする（自動カットを使わない、従来どおりの出力）。
    """
    if cuts is None:
        clips, markers, position = [], [], 0
        for index, segment in enumerate(segments, 1):
            length = _frames(segment["end"] - segment["start"], fps)
            clips.append({"id": str(index), "position": position, "length": length,
                          "src_in": _frames(segment["start"], fps),
                          "label": "%02d %s" % (index, timecode(segment["start"], fps)),
                          "media": segment.get("media"),
                          # 区間のどこが候補の時刻なのかを、マーカーで残す
                          "markers": [(peak, position + _frames(peak - segment["start"], fps))
                                      for peak in segment["peaks"]]})
            position += length
        return _xml_document(info, clips, position, fps, media_path, width, height)

    timeline = Timeline(segments, cuts, fps)
    clips = []
    for clip in timeline.clips:
        seg_index, piece = clip["segment"] + 1, clip["piece"] + 1
        # クリップ名は従来の「区間番号 元配信のタイムコード」に、区間内の連番を添える
        clips.append({"id": "%d-%d" % (seg_index, piece), "position": clip["seq_start"],
                      "length": clip["src_out"] - clip["src_in"], "src_in": clip["src_in"],
                      "label": "%02d-%02d %s" % (seg_index, piece,
                                                 timecode(clip["src_in"] / fps, fps)),
                      "media": segments[clip["segment"]].get("media"),
                      "markers": []})
    # ピークはカット後の位置に付け直す。区間の最後のクリップの後ろに並べる
    for seg_index, segment in enumerate(segments):
        owner = [c for c, t in zip(clips, timeline.clips) if t["segment"] == seg_index]
        if not owner:
            continue
        for peak in segment["peaks"]:
            owner[-1]["markers"].append((peak, timeline.nearest_sequence(peak, seg_index)))
    return _xml_document(info, clips, timeline.total_frames, fps, media_path, width, height)


def _media_frames(media, fps):
    """区間ごとの動画ファイルの (先頭のフレーム, 長さのフレーム)。

    配信全体ではなく区間付近だけをダウンロードしたときに使う。segment["media"] は
    {"path", "offset", "duration"} で、offset はファイルの0秒が元の配信の何秒か。
    """
    return _frames(media["offset"], fps), _frames(media.get("duration") or 0, fps) or 1


def _xml_document(info, clips, total_frames, fps, media_path, width, height):
    """クリップの並び（フレーム）から xmeml を組み立てる。

    クリップが "media"（区間ごとの動画ファイル）を持つときは、そのファイルを
    参照し、in/out をファイル先頭からのフレームに直す。
    """
    width = int(width or info.width or DEFAULT_WIDTH)
    height = int(height or info.height or DEFAULT_HEIGHT)
    name = (info.title or info.video_id or "highlights").strip()
    file_name = _file_name(media_path, "%s.mp4" % info.video_id)
    source_frames = _frames(info.duration or 0, fps) or 1
    file_ids = {}

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

    for clip in clips:
        position, length, src_in = clip["position"], clip["length"], clip["src_in"]
        media = clip.get("media")
        if media:
            offset, frames = _media_frames(media, fps)
            src_in = max(0, src_in - offset)
            key, path = media["path"], media["path"]
        else:
            frames, key, path = source_frames, None, media_path
        for kind, track, channel in ([("video", video_track, 0)]
                                     + [("audio", audio_tracks[i], i + 1) for i in range(2)]):
            item = ET.SubElement(track, "clipitem",
                                 {"id": "%s-%s" % (kind if kind == "video" else
                                                   "audio%d" % channel, clip["id"])})
            ET.SubElement(item, "name").text = clip["label"]
            ET.SubElement(item, "duration").text = str(frames)
            _rate(item, fps)
            ET.SubElement(item, "start").text = str(position)
            ET.SubElement(item, "end").text = str(position + length)
            ET.SubElement(item, "in").text = str(src_in)
            ET.SubElement(item, "out").text = str(src_in + length)

            if key not in file_ids:
                # ファイルの中身は最初に出てきた所で1回だけ書き、以降は id で参照する
                file_ids[key] = "media-%d" % (len(file_ids) + 1)
                file_el = ET.SubElement(item, "file", {"id": file_ids[key]})
                ET.SubElement(file_el, "name").text = (
                    _file_name(path, "") if media else file_name)
                if path:
                    ET.SubElement(file_el, "pathurl").text = _path_url(path)
                _rate(file_el, fps)
                ET.SubElement(file_el, "duration").text = str(frames)
                file_media = ET.SubElement(file_el, "media")
                file_video = ET.SubElement(file_media, "video")
                ET.SubElement(file_video, "duration").text = str(frames)
                _sample_characteristics(file_video, fps, width, height)
                file_audio = ET.SubElement(file_media, "audio")
                ET.SubElement(file_audio, "channelcount").text = "2"
            else:
                ET.SubElement(item, "file", {"id": file_ids[key]})

            if kind == "audio":
                source = ET.SubElement(item, "sourcetrack")
                ET.SubElement(source, "mediatype").text = "audio"
                ET.SubElement(source, "trackindex").text = str(channel)

        for peak, frame in clip["markers"]:
            marker = ET.SubElement(sequence, "marker")
            ET.SubElement(marker, "name").text = "ピーク %s" % timecode(peak, fps)
            ET.SubElement(marker, "in").text = str(frame)
            ET.SubElement(marker, "out").text = "-1"

    body = minidom.parseString(ET.tostring(root, encoding="utf-8")).toprettyxml(indent="  ")
    body = body.split("\n", 1)[1]        # minidom の宣言を自前のものに置き換える
    return ('<?xml version="1.0" encoding="UTF-8"?>\n'
            '<!DOCTYPE xmeml>\n' + body)


def to_edl(info, segments, fps=DEFAULT_FPS, media_path="", cuts=None):
    """どの編集ソフトでも読める EDL を組み立てる。"""
    name = (info.title or info.video_id or "highlights").strip()
    file_name = _file_name(media_path, "%s.mp4" % info.video_id)

    lines = ["TITLE: %s 見せ場" % name, "FCM: NON-DROP FRAME", ""]
    if cuts is not None:
        timeline = Timeline(segments, cuts, fps)
        for event, clip in enumerate(timeline.clips, 1):
            length = clip["src_out"] - clip["src_in"]
            media = segments[clip["segment"]].get("media")
            # 区間ごとの動画ファイルなら、ソースの時刻はそのファイルの先頭から数える
            offset = _media_frames(media, fps)[0] if media else 0
            src_in = timecode(max(0, clip["src_in"] - offset) / fps, fps)
            src_out = timecode(max(0, clip["src_out"] - offset) / fps, fps)
            rec_in = timecode(clip["seq_start"] / fps, fps)
            rec_out = timecode((clip["seq_start"] + length) / fps, fps)
            for channel in ("V", "AA"):
                lines.append("%03d  AX       %-5s C        %s %s %s %s"
                             % (event, channel, src_in, src_out, rec_in, rec_out))
            lines.append("* FROM CLIP NAME: %s"
                         % (_file_name(media["path"], file_name) if media else file_name))
            lines.append("* 区間 %02d-%02d" % (clip["segment"] + 1, clip["piece"] + 1))
            lines.append("")
        return "\n".join(lines)

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


def safe_name(info):
    """書き出すファイル名の本体（拡張子なし）。"""
    return "".join(c for c in (info.title or info.video_id)
                   if c not in '\\/:*?"<>|').strip()[:50] or "highlights"


def export(info, peaks, fmt="xml", margin_sec=DEFAULT_MARGIN_SEC,
           fps=DEFAULT_FPS, media_path="", width=0, height=0, cuts=None):
    """書き出し一式。(ファイル名, 中身, 区間の数) を返す。

    cuts は build_segments が返す区間と同じ並びの「残す部分」。
    None なら従来どおり、区間を丸ごと1クリップにする。
    """
    fps = float(fps) or DEFAULT_FPS
    segments = build_segments(peaks, margin_sec, info.duration)
    if not segments:
        raise ValueError("書き出す候補がありません。")

    safe = safe_name(info)
    if fmt == "edl":
        return ("%s.edl" % safe, to_edl(info, segments, fps, media_path, cuts),
                len(segments))
    return ("%s.xml" % safe,
            to_fcp7_xml(info, segments, fps, media_path, width, height, cuts),
            len(segments))
