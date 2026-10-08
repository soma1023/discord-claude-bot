# -*- coding: utf-8 -*-
"""編集ソフト向けの書き出しを確かめる。"""

import os
import sys
import unittest
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from stream_highlight.editexport import (build_segments, export, timecode,
                                         to_edl, to_fcp7_xml)
from stream_highlight.sources import StreamInfo

INFO = StreamInfo("twitch", "123456", "https://www.twitch.tv/videos/123456",
                  title="テスト配信", duration=9000.0)


def parse(xml_text):
    """宣言とDOCTYPEを外してから読む。"""
    return ET.fromstring(xml_text.split("\n", 2)[2])


class TestSegments(unittest.TestCase):
    def test_margin_applied(self):
        spans = build_segments([1800], margin_sec=300, duration=9000)
        self.assertEqual(len(spans), 1)
        self.assertEqual(spans[0]["start"], 1500.0)
        self.assertEqual(spans[0]["end"], 2100.0)

    def test_overlapping_peaks_merge(self):
        """近い候補の区間をまとめ、同じ場面が何度も並ばないこと。"""
        spans = build_segments([1800, 1900, 2000], margin_sec=300, duration=9000)
        self.assertEqual(len(spans), 1)
        self.assertEqual(spans[0]["start"], 1500.0)
        self.assertEqual(spans[0]["end"], 2300.0)
        self.assertEqual(len(spans[0]["peaks"]), 3)

    def test_distant_peaks_stay_separate(self):
        spans = build_segments([1800, 5000], margin_sec=300, duration=9000)
        self.assertEqual(len(spans), 2)

    def test_clamped_to_stream(self):
        """配信の外にはみ出さないこと。"""
        spans = build_segments([60, 8990], margin_sec=300, duration=9000)
        self.assertEqual(spans[0]["start"], 0.0)
        self.assertEqual(spans[-1]["end"], 9000.0)

    def test_unordered_input(self):
        spans = build_segments([5000, 1800, 1900], margin_sec=300, duration=9000)
        self.assertEqual([round(s["start"]) for s in spans], [1500, 4700])

    def test_peak_outside_stream_is_dropped(self):
        self.assertEqual(build_segments([99999], margin_sec=300, duration=9000), [])


class TestTimecode(unittest.TestCase):
    def test_formats(self):
        self.assertEqual(timecode(0, 30), "00:00:00:00")
        self.assertEqual(timecode(3661.5, 30), "01:01:01:15")
        self.assertEqual(timecode(1, 60), "00:00:01:00")

    def test_frame_rate_changes_frames(self):
        self.assertEqual(timecode(0.5, 60), "00:00:00:30")
        self.assertEqual(timecode(0.5, 30), "00:00:00:15")


class TestXml(unittest.TestCase):
    def setUp(self):
        self.spans = build_segments([1800, 5000], margin_sec=300, duration=9000)
        self.xml = to_fcp7_xml(INFO, self.spans, fps=60,
                               media_path=r"C:\Users\me\Videos\配信 本編.mp4")
        self.root = parse(self.xml)

    def test_well_formed_and_declared(self):
        self.assertTrue(self.xml.startswith('<?xml version="1.0" encoding="UTF-8"?>'))
        self.assertIn("<!DOCTYPE xmeml>", self.xml)
        self.assertEqual(self.root.tag, "xmeml")

    def test_clips_are_laid_end_to_end(self):
        """並びに隙間や重なりが無いこと。"""
        items = self.root.findall("sequence/media/video/track/clipitem")
        self.assertEqual(len(items), 2)
        position = 0
        for item in items:
            self.assertEqual(int(item.find("start").text), position)
            position = int(item.find("end").text)
        self.assertEqual(int(self.root.find("sequence/duration").text), position)

    def test_source_points_match_the_segment(self):
        item = self.root.find("sequence/media/video/track/clipitem")
        self.assertEqual(int(item.find("in").text), int(1500 * 60))
        self.assertEqual(int(item.find("out").text), int(2100 * 60))

    def test_audio_tracks_reference_the_same_file(self):
        tracks = self.root.findall("sequence/media/audio/track")
        self.assertEqual(len(tracks), 2)
        for index, track in enumerate(tracks, 1):
            for item in track.findall("clipitem"):
                self.assertEqual(item.find("file").get("id"), "media-1")
                self.assertEqual(item.find("sourcetrack/trackindex").text, str(index))

    def test_file_defined_once(self):
        """素材の定義は1回だけで、以降は参照にすること。"""
        defined = [f for f in self.root.iter("file") if f.find("name") is not None]
        self.assertEqual(len(defined), 1)
        self.assertEqual(defined[0].find("name").text, "配信 本編.mp4")

    def test_windows_path_becomes_file_url(self):
        url = self.root.find(".//file/pathurl").text
        self.assertTrue(url.startswith("file://localhost/C:/"), url)
        self.assertNotIn("C%3A", url)      # ドライブのコロンは残す

    def test_markers_point_at_peaks(self):
        markers = self.root.findall("sequence/marker")
        self.assertEqual(len(markers), 2)
        # 1件目は区間先頭(1500秒)から300秒後
        self.assertEqual(int(markers[0].find("in").text), int(300 * 60))

    def test_frame_rate_is_written(self):
        self.assertEqual(self.root.find("sequence/rate/timebase").text, "60")
        self.assertEqual(self.root.find("sequence/rate/ntsc").text, "FALSE")

    def test_ntsc_rate(self):
        root = parse(to_fcp7_xml(INFO, self.spans, fps=29.97))
        self.assertEqual(root.find("sequence/rate/timebase").text, "30")
        self.assertEqual(root.find("sequence/rate/ntsc").text, "TRUE")

    def test_without_media_path(self):
        root = parse(to_fcp7_xml(INFO, self.spans, fps=30))
        self.assertIsNone(root.find(".//file/pathurl"))
        self.assertEqual(root.find(".//file/name").text, "123456.mp4")


class TestEdl(unittest.TestCase):
    def test_structure(self):
        spans = build_segments([1800, 5000], margin_sec=300, duration=9000)
        edl = to_edl(INFO, spans, fps=30)
        self.assertTrue(edl.startswith("TITLE:"))
        self.assertIn("FCM: NON-DROP FRAME", edl)
        video = [l for l in edl.split("\n") if "  V " in l]
        audio = [l for l in edl.split("\n") if "  AA " in l]
        self.assertEqual(len(video), 2)
        self.assertEqual(len(audio), 2)
        self.assertIn("00:25:00:00 00:35:00:00 00:00:00:00 00:10:00:00", video[0])


class TestExport(unittest.TestCase):
    def test_filename_is_safe(self):
        info = StreamInfo("twitch", "1", "u", title='ひどい/名前:の*配信?', duration=9000)
        name, _, _ = export(info, [1800])
        for bad in '\\/:*?"<>|':
            self.assertNotIn(bad, name)

    def test_empty_peaks_rejected(self):
        with self.assertRaises(ValueError):
            export(INFO, [])

    def test_format_selection(self):
        xml_name, xml_body, _ = export(INFO, [1800], fmt="xml")
        edl_name, edl_body, _ = export(INFO, [1800], fmt="edl")
        self.assertTrue(xml_name.endswith(".xml"))
        self.assertTrue(edl_name.endswith(".edl"))
        self.assertIn("xmeml", xml_body)
        self.assertIn("TITLE:", edl_body)

    def test_segment_count_reported(self):
        _, _, count = export(INFO, [1800, 1900, 5000], margin_sec=300)
        self.assertEqual(count, 2)


if __name__ == "__main__":
    unittest.main(verbosity=2)
