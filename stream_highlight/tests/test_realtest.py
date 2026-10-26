# -*- coding: utf-8 -*-
"""実配信テストの補助（区間ダウンロードの位置合わせ・点検）を確かめる。

ネットワークもモデルも使わない。音声は乱数で作る。
"""

import json
import os
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

try:
    import numpy as np
except ImportError:          # numpy が無い環境では位置合わせのテストを飛ばす
    np = None

from stream_highlight import editexport, notify, realtest
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

        ref = ("ref", requested - realtest.REF_BEFORE_SEC, "DASH")
        with mock.patch.object(realtest, "reference_audio", return_value=ref), \
                mock.patch.object(realtest.transcribe, "extract_audio", side_effect=audio):
            return realtest.measure_start("url", "part", requested, "folder", 1)

    def test_measured_start_matches_actual(self):
        for actual in (100.0, 97.63, 100.4):
            with self.subTest(actual=actual):
                measured, score, kind = self.measure(actual)
                self.assertAlmostEqual(measured, actual, places=2)
                self.assertGreater(score, 0.99)
                self.assertEqual(kind, "DASH")


class TestSidx(unittest.TestCase):
    def test_fragments_have_time_and_bytes(self):
        import struct
        # version 0, timescale 1000, 最初の断片は 2.5 秒、目次の直後から 100 バイト空けて始まる
        body = struct.pack(">B3xIIIIHH", 0, 1, 1000, 2500, 100, 0, 3)
        for size, duration in ((4000, 10000), (5000, 9500), (3000, 10000)):
            body += struct.pack(">III", size, duration, 0x90000000)
        fragments = realtest.parse_sidx(body, base=800)
        self.assertEqual(fragments, [(2.5, 900, 4000), (12.5, 4900, 5000), (22.0, 9900, 3000)])


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


class TestProgress(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.mkdtemp()
        self.path = os.path.join(self.folder, "progress.json")
        self.sent = []
        self.now = [1000.0]
        stages = realtest.plan_stages(5, ["qwen", "whisper"], download=True, copy_check=True)
        self.progress = realtest.Progress(self.path, "配信", "https://x", stages,
                                          notifier=self.sent.append,
                                          clock=lambda: self.now[0])

    def read(self):
        with open(self.path, encoding="utf-8") as fh:
            return json.load(fh)

    def test_percent_follows_stage_weights(self):
        total = self.progress.total
        download = 5 * realtest.WEIGHT_DOWNLOAD + realtest.WEIGHT_COPY_CHECK
        self.progress.update("download", 1.0, "済")
        self.assertAlmostEqual(self.progress.percent, 100 * download / total)
        self.progress.update("qwen", 0.5, "途中")
        qwen = realtest.WEIGHT_ENGINE_LOAD["qwen"] + 5 * realtest.WEIGHT_ENGINE["qwen"]
        self.assertAlmostEqual(self.progress.percent, 100 * (download + qwen / 2) / total)
        data = self.read()
        self.assertEqual(data["state"], "running")
        self.assertEqual(data["message"], "途中")

    def test_percent_never_goes_back(self):
        self.progress.update("qwen", 0.5, "a")
        before = self.progress.percent
        self.progress.update("download", 0.2, "b")
        self.assertEqual(self.progress.percent, before)

    def test_eta_from_elapsed(self):
        self.progress.update("download", 0.0, "a")
        self.assertIsNone(self.progress.eta_sec())          # 始めのうちは出さない
        self.now[0] += 100
        self.progress.percent = 25.0
        self.assertAlmostEqual(self.progress.eta_sec(), 300.0)

    def test_notify_only_when_asked(self):
        self.progress.update("download", 0.1, "静か")
        self.assertEqual(self.sent, [])
        self.progress.update("download", 0.2, "区間 1/5 をダウンロードしました", notify=True)
        self.assertEqual(len(self.sent), 1)
        self.assertIn("配信", self.sent[0])
        self.assertIn("全体", self.sent[0])

    def test_finish_and_fail(self):
        self.progress.finish("完了", "まとめ")
        self.assertEqual(self.read()["state"], "done")
        self.assertEqual(self.read()["percent"], 100.0)
        self.assertIn("まとめ", self.sent[-1])
        self.progress.fail("回線が切れた")
        self.assertEqual(self.read()["state"], "error")
        self.assertIn("回線が切れた", self.sent[-1])

    def test_engine_reporter_maps_segments(self):
        report = realtest.engine_reporter("qwen", 5, self.progress)
        with mock.patch.object(realtest, "say"):
            report("区間 1/5 を文字起こし中…", 0.0)
            start = self.progress.percent
            report("（2/3 分割）")                           # frac が無い表示は進めない
            self.assertEqual(self.progress.percent, start)
            report("区間 3/5 を文字起こし中…", 0.4)
        base = 5 * realtest.WEIGHT_DOWNLOAD + realtest.WEIGHT_COPY_CHECK
        done = realtest.WEIGHT_ENGINE_LOAD["qwen"] + 2 * realtest.WEIGHT_ENGINE["qwen"]
        self.assertAlmostEqual(self.progress.percent, 100 * (base + done) / self.progress.total)
        self.assertEqual(len(self.sent), 2)                 # 区間の始まりだけ通知


class TestNotify(unittest.TestCase):
    HOOK = "https://discord.com/api/webhooks/1/abc"

    def test_only_discord_webhooks_are_used(self):
        for value, expected in ((self.HOOK, self.HOOK), ("https://example.com/hook", ""),
                                ("", "")):
            with self.subTest(value=value), \
                    mock.patch.object(notify.retake, "setting", return_value=value):
                self.assertEqual(notify.webhook_url(), expected)

    def test_post_sends_json_and_truncates(self):
        seen = []

        class Response:
            def __enter__(self):
                return self

            def __exit__(self, *a):
                return False

        def opener(request, timeout):
            seen.append(request)
            return Response()

        self.assertTrue(notify.post("あ" * 3000, url=self.HOOK, opener=opener))
        body = json.loads(seen[0].data.decode("utf-8"))
        self.assertEqual(len(body["content"]), notify.MAX_CHARS)
        self.assertIn("DiscordBot", seen[0].get_header("User-agent"))

    def test_failure_does_not_raise(self):
        def opener(request, timeout):
            raise OSError("offline")
        self.assertFalse(notify.post("x", url=self.HOOK, opener=opener))

    def test_no_webhook_means_no_post(self):
        with mock.patch.object(notify.retake, "setting", return_value=""):
            self.assertFalse(notify.post("x", opener=lambda *a, **k: self.fail("送った")))


class TestRealtestProgressApi(unittest.TestCase):
    def write(self, folder, **data):
        os.makedirs(folder, exist_ok=True)
        with open(os.path.join(folder, "progress.json"), "w", encoding="utf-8") as fh:
            json.dump(data, fh)

    def test_latest_running_done_and_stalled(self):
        from stream_highlight import server

        base = tempfile.mkdtemp()
        with mock.patch.object(server.paths, "data_dir", return_value=base):
            self.assertIsNone(server.latest_realtest_progress(now=5000))
            self.write(os.path.join(base, "realtest", "a"), state="running", percent=40,
                       updated_at=4990)
            data = server.latest_realtest_progress(now=5000)
            self.assertEqual(data["percent"], 40)
            self.assertFalse(data["stalled"])
            self.assertTrue(server.latest_realtest_progress(now=4990 + 16 * 60)["stalled"])
            self.write(os.path.join(base, "realtest", "a"), state="done", updated_at=5000)
            self.assertIsNotNone(server.latest_realtest_progress(now=5000 + 60))
            self.assertIsNone(server.latest_realtest_progress(now=5000 + 31 * 60))


if __name__ == "__main__":
    unittest.main()
