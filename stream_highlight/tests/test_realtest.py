# -*- coding: utf-8 -*-
"""実配信テストの補助（区間ダウンロードの位置合わせ・点検）を確かめる。

ネットワークもモデルも使わない。音声は乱数で作る。
"""

import os
import sys
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

try:
    import numpy as np
except ImportError:          # numpy が無い環境では位置合わせのテストを飛ばす
    np = None

from stream_highlight import editexport, realtest
from stream_highlight.transcribe import Word

RATE = 16000


@unittest.skipIf(np is None, "numpy が入っていない")
class TestAlignment(unittest.TestCase):
    def setUp(self):
        rng = np.random.default_rng(3)
        self.stream = rng.standard_normal(RATE * 40).astype("float32")   # 40秒の「配信の音」

    def test_lag_is_found(self):
        reference = self.stream[:RATE * 30]
        clip = self.stream[int(RATE * 12.345):int(RATE * 22.345)]
        lag, score = realtest.find_lag(reference, clip, RATE)
        self.assertAlmostEqual(lag, 12.345, places=3)
        self.assertGreater(score, 0.99)

    def test_unrelated_audio_scores_low(self):
        other = np.random.default_rng(9).standard_normal(RATE * 10).astype("float32")
        _, score = realtest.find_lag(self.stream[:RATE * 30], other, RATE)
        self.assertLess(score, realtest.MIN_MATCH_SCORE)

    def measure(self, file_start, requested=100.0):
        """配信の file_start 秒から始まる区間ファイルを、requested 秒として落とした想定。"""
        stream = np.concatenate([np.zeros(RATE * 80, "float32"), self.stream])  # 80秒〜が音

        def audio(path, start, end):
            base = (requested - realtest.REF_BEFORE_SEC) if path == "ref" else file_start
            lo = int(round((base + start) * RATE)) - RATE * 80 + RATE * 80
            return stream[lo:lo + int(round((end - start) * RATE))]

        with mock.patch.object(realtest, "_downloaded", return_value="ref"), \
                mock.patch.object(realtest.transcribe, "extract_audio", side_effect=audio):
            return realtest.measure_start("url", "part", requested, "folder", 1)

    def test_measured_start_matches_actual(self):
        for actual in (100.0, 97.63, 100.4):
            with self.subTest(actual=actual):
                measured, score = self.measure(actual)
                self.assertAlmostEqual(measured, actual, places=2)
                self.assertGreater(score, 0.99)


class TestChecks(unittest.TestCase):
    def setUp(self):
        segments = [{"start": 0.0, "end": 60.0, "peaks": [30.0]}]
        self.timeline = editexport.Timeline(segments, [[(0.0, 60.0)]], 30.0)

    def cue(self, text):
        return {"start_frame": 300, "text": text}

    def test_suspicious_cues(self):
        found = realtest.suspicious_cues([
            self.cue("今日はいい天気ですね"),
            self.cue("ありがとうありがとうありがとう"),
            self.cue("ご視聴ありがとうございました"),
            self.cue("wwwwwwwwwwww"),
            self.cue("Game over now"),
        ], self.timeline)
        reasons = [f[0] for f in found]
        self.assertEqual(reasons, ["同じ語の繰り返し", "定型文（音の無い所で出やすい）", "英字が多い"])
        self.assertEqual(found[0][2], 10.0)          # 元の配信の秒

    def test_parse_srt(self):
        cues = realtest.parse_srt("1\n00:00:10,000 --> 00:00:11,000\nあい\nうえ\n", 30.0)
        self.assertEqual(cues, [{"start_frame": 300, "text": "あい\nうえ"}])

    def test_one_engine_only_speech_is_listed(self):
        segments = [{"start": 0.0, "end": 60.0, "peaks": [30.0]}]
        words = {"qwen": [[Word("こんにちは", 5.0, 6.0), Word("ドーン", 40.0, 40.5)]],
                 "whisper": [[Word("こんにちは", 5.1, 6.0)]]}
        with mock.patch.object(realtest, "loudness", return_value=[]):
            found = realtest.false_detection_candidates(segments, words, None, "x.mp4")
        self.assertEqual([(f[0], f[4]) for f in found], [("qwen", "ドーン")])


if __name__ == "__main__":
    unittest.main()
