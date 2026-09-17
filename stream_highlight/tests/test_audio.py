# -*- coding: utf-8 -*-
"""音声解析と、チャット照合による誤検知フィルタの検証。

狙いは「ゲームのSEで音量だけ跳ねた箇所」を落とし、
「配信者が声を上げてチャットも反応した箇所」を残すこと。
"""

import os
import random
import subprocess
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from stream_highlight import audio as audio_mod
from stream_highlight.analyze import Params, analyze
from stream_highlight.audio import Loudness, SILENCE_FLOOR
from stream_highlight.sources import ChatMessage
from stream_highlight.tests.test_analyze import DURATION, INFO, SPIKES, build_chat

# 配信者が声を上げた箇所。チャットのピークより少し前に起きる（チャットは遅れて反応する）
VOICE_EVENTS = [s - 3.0 for s in SPIKES]
# ゲームのSEなど、音量だけ跳ねてチャットが動かない箇所
SFX_EVENTS = [1500.0, 2000.0, 4200.0, 6000.0]
# 配信者は叫んだが、チャットの盛り上がりは控えめだった箇所
QUIET_REACTION = 4800.0


def build_loudness(events, duration=DURATION, seed=3):
    """平常 -30 LUFS に、指定時刻の大音量を差し込んだ列を作る。"""
    rng = random.Random(seed)
    n = int(duration / 0.1)
    values = [-30.0 + rng.uniform(-1.5, 1.5) for _ in range(n)]
    for start, length, db in events:
        for i in range(int(start / 0.1), int((start + length) / 0.1)):
            if 0 <= i < n:
                values[i] = db + rng.uniform(-0.5, 0.5)
    return Loudness(step=0.1, values=values)


def scenario():
    """チャットと音声がそろった、現実に近い配信を組み立てる。"""
    messages = build_chat()
    rng = random.Random(11)
    # 控えめな反応: 見せ場のしきい値(chat_z 3.0)には届かないが、増加ははっきり出る規模。
    # 実測では 14件 で chat_z 2.1 / chat_rise 2.4 になり、音声側でだけ拾える帯に入る。
    for _ in range(14):
        messages.append(ChatMessage(QUIET_REACTION + 2 + rng.uniform(0, 5),
                                    "viewer%d" % rng.randint(1, 500), "おお"))
    messages.sort(key=lambda m: m.offset)

    events = [(t, 2.5, -12.0) for t in VOICE_EVENTS]
    events += [(t, 1.0, -11.0) for t in SFX_EVENTS]
    events += [(QUIET_REACTION, 2.0, -11.5)]
    return messages, build_loudness(events)


class TestLoudness(unittest.TestCase):
    def test_bin_max_keeps_short_bursts(self):
        """5秒ビンに落としても、1秒だけの叫びが均されて消えないこと。"""
        lo = build_loudness([(100.0, 1.0, -10.0)])
        bins = lo.bin_max(5.0, int(DURATION // 5))
        self.assertGreater(bins[20], -15.0)
        self.assertLess(bins[10], -25.0)

    def test_padding_beyond_audio_end(self):
        lo = Loudness(step=0.1, values=[-20.0] * 100)
        bins = lo.bin_max(5.0, 10)
        self.assertEqual(len(bins), 10)
        self.assertEqual(bins[0], -20.0)
        self.assertEqual(bins[-1], SILENCE_FLOOR)

    def test_roundtrip(self):
        lo = build_loudness([(10.0, 1.0, -10.0)], duration=60)
        again = Loudness.from_dict(lo.as_dict())
        self.assertEqual(len(again.values), len(lo.values))
        self.assertAlmostEqual(again.values[105], lo.values[105], places=1)


class TestCrossValidation(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.messages, cls.loudness = scenario()
        cls.result = analyze(cls.messages, INFO, Params(), loudness=cls.loudness)

    def test_audio_available(self):
        self.assertTrue(self.result["audio"]["available"])
        self.assertGreater(self.result["audio"]["spread_db"], 0)

    def test_chat_moments_get_audio_annotation(self):
        """チャットの見せ場に、対応する音声の跳ねが紐づくこと。"""
        for moment in self.result["moments"][:5]:
            if not any(abs(moment["peak_sec"] - s) <= 20 for s in SPIKES):
                continue
            self.assertIsNotNone(moment["audio"], moment["peak_sec"])
            self.assertGreater(moment["audio"]["excess_db"], 10.0,
                               "声の跳ねを拾えていない: %s" % moment["peak_sec"])
            # 音声イベントはチャットのピークより前にあるはず
            self.assertLessEqual(moment["audio"]["at"], moment["peak_sec"] + 5)

    def test_game_sfx_is_rejected(self):
        """チャットが反応していない音量ピークが候補に残らないこと。"""
        found = [m["peak_sec"] for m in self.result["audio_moments"]]
        for sfx in SFX_EVENTS:
            self.assertFalse(
                any(abs(f - sfx) <= 20 for f in found),
                "SEを見せ場として拾ってしまった: %.0f秒 (検出: %s)" % (sfx, found),
            )

    def test_filter_actually_rejects(self):
        stats = self.result["audio"]["stats"]
        self.assertGreaterEqual(stats["rejected"], len(SFX_EVENTS),
                                "SEが弾かれていない: %s" % stats)
        self.assertEqual(stats["raw_peaks"], stats["confirmed"] + stats["rejected"])

    def test_finds_moment_chat_alone_missed(self):
        """叫んだがチャットは控えめ、という箇所を音声側で拾えること。"""
        chat_only = [m["peak_sec"] for m in self.result["moments"]]
        self.assertFalse(any(abs(t - QUIET_REACTION) <= 20 for t in chat_only),
                         "前提が崩れている: チャット単独で拾えてしまっている")
        found = [m["peak_sec"] for m in self.result["audio_moments"]]
        self.assertTrue(any(abs(t - QUIET_REACTION) <= 25 for t in found),
                        "音声側でも拾えていない (検出: %s)" % found)

    def test_audio_moments_are_not_duplicates(self):
        """すでにチャット側で出ている箇所は、音声タブに重複させない。"""
        chat_peaks = [m["peak_sec"] for m in self.result["moments"]]
        for moment in self.result["audio_moments"]:
            self.assertFalse(any(abs(moment["peak_sec"] - c) <= 45 for c in chat_peaks),
                             "重複: %s" % moment["peak_sec"])
            self.assertEqual(moment["source"], "audio")

    def test_without_audio_nothing_breaks(self):
        plain = analyze(self.messages, INFO, Params())
        self.assertFalse(plain["audio"]["available"])
        self.assertEqual(plain["audio_moments"], [])
        self.assertIsNone(plain["moments"][0].get("audio"))
        self.assertEqual(plain["moments"][0]["source"], "chat")

    def test_support_threshold_changes_strictness(self):
        loose = analyze(self.messages, INFO, Params(chat_support_z=0.0),
                        loudness=self.loudness)
        strict = analyze(self.messages, INFO, Params(chat_support_z=6.0),
                         loudness=self.loudness)
        self.assertGreater(loose["audio"]["stats"]["confirmed"],
                           strict["audio"]["stats"]["confirmed"])
        # 照合を外すとSEが混ざる = フィルタが実際に効いている証拠
        loose_found = [m["peak_sec"] for m in loose["audio_moments"]]
        self.assertTrue(any(any(abs(f - sfx) <= 20 for f in loose_found)
                            for sfx in SFX_EVENTS),
                        "照合を外してもSEが出ない＝テストが機能していない")


class TestFfmpegIntegration(unittest.TestCase):
    """実際の ffmpeg を通して、ラウドネスを取り出せるか確かめる。"""

    @classmethod
    def setUpClass(cls):
        if not audio_mod.ffmpeg_available():
            raise unittest.SkipTest("ffmpeg が無い環境なのでスキップ")

    def test_extract_from_generated_wav(self):
        import math
        import struct
        import tempfile
        import wave

        path = os.path.join(tempfile.mkdtemp(), "tone.wav")
        sr = 16000
        with wave.open(path, "w") as fh:
            fh.setnchannels(1)
            fh.setsampwidth(2)
            fh.setframerate(sr)
            frames = bytearray()
            for i in range(sr * 20):
                t = i / sr
                amp = 20000 if 8.0 <= t < 10.0 else 1500   # 8〜10秒だけ大音量
                frames += struct.pack("<h", int(amp * math.sin(2 * math.pi * 300 * t)))
            fh.writeframes(bytes(frames))

        loudness = audio_mod.extract_loudness(path, duration=20)
        self.assertGreater(len(loudness.values), 150)
        bins = loudness.bin_max(1.0, 20)
        self.assertGreater(bins[9] - bins[3], 15.0,
                           "大音量区間が検出できていない: %s" % bins[:12])


if __name__ == "__main__":
    unittest.main(verbosity=2)
