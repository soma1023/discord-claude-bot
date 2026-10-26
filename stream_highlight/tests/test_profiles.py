# -*- coding: utf-8 -*-
"""チャンネルごとの設定（profiles.py）と、設定で変わるテロップの体裁を確かめる。"""

import os
import shutil
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from stream_highlight import cache, editexport, profiles, subtitles
from stream_highlight.sources import StreamInfo
from stream_highlight.transcribe import Word


def info(channel_id="UCabc", channel="結川ユイナ", video_id="vid00000001"):
    return StreamInfo("youtube", video_id, "https://www.youtube.com/watch?v=" + video_id,
                      title="配信", channel=channel, duration=3600.0, channel_id=channel_id)


class TempDataDir(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.mkdtemp(prefix="sh_prof_")
        patcher = mock.patch.object(profiles.paths, "data_dir", return_value=self.folder)
        patcher.start()
        self.addCleanup(patcher.stop)
        self.addCleanup(shutil.rmtree, self.folder, True)


class TestProfiles(TempDataDir):
    def test_nothing_saved_means_defaults(self):
        self.assertEqual(profiles.settings_for(info()), {})
        self.assertFalse(profiles.describe(info())["saved"])

    def test_save_and_load_by_channel(self):
        profiles.save(info(), {"gap_sec": 0.8, "line_chars": 12, "engine": "qwen",
                               "autocut": True, "unknown": 1})
        got = profiles.settings_for(info(video_id="vid00000002"))     # 同じチャンネルの別の配信
        self.assertEqual(got, {"gap_sec": 0.8, "line_chars": 12, "engine": "qwen",
                               "autocut": True})
        self.assertEqual(profiles.settings_for(info(channel_id="UCother", channel="別の人")), {})

    def test_values_are_clamped(self):
        got = profiles.clean({"gap_sec": 99, "pad_sec": -1, "max_lines": "x", "engine": "gpt",
                              "min_show_sec": 3.0, "max_show_sec": 2.0})
        self.assertEqual(got, {"gap_sec": 5.0, "pad_sec": 0.0, "min_show_sec": 3.0,
                               "max_show_sec": 3.0})

    def test_old_cache_without_id_finds_by_name(self):
        """チャンネル ID を持たない古いキャッシュの配信でも、表示名で同じ設定が出る。"""
        profiles.save(info(), {"gap_sec": 2.0})
        self.assertEqual(profiles.settings_for(info(channel_id="")), {"gap_sec": 2.0})

    def test_name_key_moves_to_id_key(self):
        profiles.save(info(channel_id=""), {"gap_sec": 2.0})
        self.assertIn("youtube:name:結川ユイナ", profiles.load_all())
        profiles.save(info(), {"gap_sec": 2.5})
        self.assertEqual(list(profiles.load_all()), ["youtube:UCabc"])
        self.assertEqual(profiles.settings_for(info()), {"gap_sec": 2.5})

    def test_remove(self):
        profiles.save(info(), {"gap_sec": 2.0})
        self.assertTrue(profiles.remove(info()))
        self.assertEqual(profiles.settings_for(info()), {})
        self.assertFalse(profiles.remove(info()))

    def test_channel_unknown_cannot_save(self):
        with self.assertRaises(ValueError):
            profiles.save(info(channel_id="", channel=""), {"gap_sec": 2.0})

    def test_broken_file_is_ignored(self):
        with open(os.path.join(self.folder, profiles.FILE_NAME), "w", encoding="utf-8") as fh:
            fh.write("{壊れている")
        self.assertEqual(profiles.settings_for(info()), {})


class TestCaptionStyle(unittest.TestCase):
    def setUp(self):
        segments = [{"start": 0.0, "end": 60.0, "peaks": [30.0]}]
        self.timeline = editexport.Timeline(segments, [[(0.0, 60.0)]], 30.0)
        self.words = []
        t = 1.0
        for _ in range(30):
            self.words.append(Word("ながめの言葉", t, t + 0.4))
            t += 0.45

    def test_default_style_matches_previous_constants(self):
        self.assertEqual(subtitles.make_cues(self.words, self.timeline),
                         subtitles.make_cues(self.words, self.timeline, subtitles.CaptionStyle()))

    def test_narrow_single_line_style(self):
        style = subtitles.CaptionStyle(line_chars=8, max_lines=1, min_show_sec=0.5,
                                       max_show_sec=2.0)
        cues = subtitles.make_cues(self.words, self.timeline, style)
        for cue in cues:
            self.assertEqual(len(cue["lines"]), 1)
            self.assertLessEqual(len(cue["lines"][0]), 8)
            self.assertLessEqual(cue["end"] - cue["start"], 2.0 + 1e-6)
        self.assertGreater(len(cues), len(subtitles.make_cues(self.words, self.timeline)))

    def test_from_values_clamps(self):
        style = subtitles.CaptionStyle.from_values("x", 99, 0.0, 0.5)
        self.assertEqual((style.line_chars, style.max_lines), (subtitles.LINE_CHARS, 4))
        self.assertEqual(style.min_show_sec, 0.2)
        self.assertEqual(style.max_show_sec, 1.0)


class TestProfileApi(TempDataDir):
    def setUp(self):
        super().setUp()
        from fastapi.testclient import TestClient

        from stream_highlight import server
        from stream_highlight.jobs import manager

        original = cache.CACHE_DIR
        cache.CACHE_DIR = self.folder
        self.addCleanup(setattr, cache, "CACHE_DIR", original)
        cache.save(info(), [])
        manager._loaded.clear()
        self.client = TestClient(server.app)
        self.key = "youtube:vid00000001"

    def test_get_save_reset(self):
        res = self.client.get("/api/profile", params={"video_key": self.key})
        self.assertEqual(res.status_code, 200, res.text)
        self.assertFalse(res.json()["saved"])
        self.assertEqual(res.json()["name"], "結川ユイナ")

        res = self.client.post("/api/profile", json={"video_key": self.key,
                                                     "settings": {"gap_sec": 0.9, "export_top": 8}})
        self.assertEqual(res.status_code, 200, res.text)
        self.assertTrue(res.json()["saved"])
        self.assertEqual(res.json()["settings"], {"gap_sec": 0.9, "export_top": 8})

        res = self.client.post("/api/profile/reset", json={"video_key": self.key})
        self.assertFalse(res.json()["saved"])

    def test_unknown_video_is_404(self):
        res = self.client.get("/api/profile", params={"video_key": "youtube:nothing0000"})
        self.assertEqual(res.status_code, 404)


if __name__ == "__main__":
    unittest.main()
