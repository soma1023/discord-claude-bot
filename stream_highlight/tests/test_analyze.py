# -*- coding: utf-8 -*-
"""合成チャットで、仕込んだ盛り上がりを検出できるか確かめる。"""

import math
import random
import sys
import os
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from stream_highlight.analyze import Params, analyze, analyze_keyword, rolling_median
from stream_highlight.sources import ChatMessage, StreamInfo
from stream_highlight.patterns import normalize, classify

DURATION = 7200.0   # 2時間
SPIKES = [600.0, 2400.0, 3000.0, 5400.0, 6900.0]
FILLER = ["こんばんは", "おつ", "そうだね", "なるほど", "はい", "ここすき", "見てる"]


def build_chat(seed=0):
    """視聴者数が増減する平常チャット + 仕込んだ爆発。"""
    rng = random.Random(seed)
    messages = []

    # 平常運転: 配信中盤に向けて人が増え、後半に減る（毎分12〜60コメント）
    t = 0.0
    while t < DURATION:
        ramp = 1.0 + 3.0 * math.sin(math.pi * t / DURATION)
        rate_per_sec = 0.2 * ramp
        t += rng.expovariate(rate_per_sec)
        if t < DURATION:
            messages.append(ChatMessage(t, "viewer%d" % rng.randint(1, 500),
                                        rng.choice(FILLER)))

    # 仕込み: 各スパイクで8秒間に120コメントの草が飛ぶ
    for peak in SPIKES:
        for _ in range(120):
            offset = peak + rng.uniform(0.0, 8.0)
            messages.append(ChatMessage(offset, "viewer%d" % rng.randint(1, 500),
                                        rng.choice(["wwwww", "ｗｗｗ", "草", "www"])))

    messages.sort(key=lambda m: m.offset)
    return messages


INFO = StreamInfo("youtube", "testvideo00", "https://youtu.be/testvideo00",
                  title="テスト配信", duration=DURATION)


class TestAnalyze(unittest.TestCase):
    def setUp(self):
        self.messages = build_chat()
        self.result = analyze(self.messages, INFO, Params())

    def test_finds_every_planted_spike(self):
        """仕込んだ5箇所すべてが上位に入ること。"""
        found = [m["peak_sec"] for m in self.result["moments"]]
        for spike in SPIKES:
            near = [f for f in found if abs(f - spike) <= 20]
            self.assertTrue(near, "%.0f秒 の盛り上がりを検出できていない (検出: %s)" % (spike, found))

    def test_spikes_rank_at_top(self):
        """仕込みが上位5件を占めること（誤検出が混ざっていない）。"""
        top5 = self.result["moments"][:5]
        for moment in top5:
            self.assertTrue(
                any(abs(moment["peak_sec"] - s) <= 20 for s in SPIKES),
                "上位に想定外の候補: %.0f秒" % moment["peak_sec"],
            )

    def test_baseline_follows_traffic_changes(self):
        """平常値が視聴者数の増減に追従し、賑やかな時間帯に偏らないこと。"""
        moments = self.result["moments"][:5]
        # 中盤(2400,3000)と序盤(600)はコメント総数が3倍近く違うが、
        # どちらもスコアが出ていれば平常値が時間帯ごとに効いている
        early = [m for m in moments if abs(m["peak_sec"] - 600) <= 20]
        mid = [m for m in moments if abs(m["peak_sec"] - 3000) <= 20]
        self.assertTrue(early and mid)
        self.assertGreater(early[0]["rate"], 0)
        self.assertLess(early[0]["baseline_rate"], mid[0]["baseline_rate"],
                        "平常値が時間帯で変化していない")

    def test_moment_payload(self):
        m = self.result["moments"][0]
        self.assertLess(m["clip_start"], m["peak_sec"])
        self.assertGreater(m["clip_end"], m["peak_sec"])
        self.assertIn("t=", m["url"])
        self.assertTrue(m["top_comments"])
        self.assertTrue(any(t["id"] == "laugh" for t in m["tags"]), m["tags"])
        self.assertGreater(m["messages"], 50)

    def test_category_ranking(self):
        laugh = self.result["category_moments"].get("laugh")
        self.assertTrue(laugh, "草ランキングが空")
        top = laugh[0]["peak_sec"]
        self.assertTrue(any(abs(top - s) <= 20 for s in SPIKES))

    def test_series_shape(self):
        series = self.result["series"]
        self.assertEqual(len(series["rate"]), len(series["baseline"]))
        self.assertLessEqual(len(series["rate"]), 1800)
        self.assertGreater(max(series["rate"]), max(series["baseline"]))

    def test_keyword_search(self):
        found = analyze_keyword(self.messages, INFO, "www", Params())
        self.assertGreater(found["hits"], 100)
        self.assertTrue(found["moments"])
        self.assertTrue(any(abs(found["moments"][0]["peak_sec"] - s) <= 20 for s in SPIKES))

        empty = analyze_keyword(self.messages, INFO, "存在しない語", Params())
        self.assertEqual(empty["hits"], 0)
        self.assertEqual(empty["moments"], [])

    def test_quiet_chat_does_not_produce_junk(self):
        """過疎チャットで数コメントが1位になったりしないこと。"""
        rng = random.Random(1)
        sparse = [
            ChatMessage(rng.uniform(0, DURATION), "v", "ふむ")
            for _ in range(60)
        ]
        sparse.sort(key=lambda m: m.offset)
        result = analyze(sparse, INFO, Params())
        self.assertEqual(result["moments"], [])
        self.assertIn("少ない", result["notice"])

    def test_system_notices_do_not_become_moments(self):
        """サブスク告知が集中しても、見せ場として拾わないこと。"""
        messages = list(self.messages)
        for i in range(40):
            messages.append(ChatMessage(4000 + i * 0.4, "sub%d" % i, "", kind="system"))
        messages.sort(key=lambda m: m.offset)
        result = analyze(messages, INFO, Params())
        self.assertEqual(result["stats"]["system_notices"], 40)
        self.assertFalse(any(abs(m["peak_sec"] - 4000) <= 30 for m in result["moments"]),
                         "システム通知を見せ場にしている")
        # 件数にも数えない
        self.assertEqual(result["stats"]["messages"],
                         len([m for m in messages if m.kind != "system"]))

    def test_no_messages(self):
        result = analyze([], INFO, Params())
        self.assertEqual(result["moments"], [])
        self.assertEqual(result["stats"]["messages"], 0)

    def test_sensitivity_changes_result_count(self):
        loose = analyze(self.messages, INFO, Params(min_z=1.5))
        strict = analyze(self.messages, INFO, Params(min_z=6.0))
        self.assertGreaterEqual(len(loose["moments"]), len(strict["moments"]))

    def test_params_clamped(self):
        p = Params.from_dict({"bin_sec": 999, "min_z": -5, "top_n": "20", "window_sec": None})
        self.assertEqual(p.bin_sec, 60)
        self.assertEqual(p.min_z, 0.5)
        self.assertEqual(p.top_n, 20)
        # 刻みより短い窓は意味がないので、bin_sec まで引き上げられる
        self.assertEqual(p.window_sec, 60)

    def test_rolling_median_ignores_spikes(self):
        values = [10.0] * 200
        values[100] = 5000.0
        med = rolling_median(values, 50)
        self.assertLess(med[100], 20.0)


class TestPatterns(unittest.TestCase):
    def test_laugh_variants(self):
        for text in ["wwwww", "ｗｗｗｗ", "草", "大草原", "まじかwww", "LUL", "lmao"]:
            self.assertIn("laugh", classify(normalize(text)), text)

    def test_twitch_channel_emotes(self):
        """Twitchのチャンネル絵文字（○○Www 形式）も草として扱うこと。"""
        for text in ["ajak0nWww", "kawaiiWWW", "hogeLUL", "fugaKEKW", "pepeLaugh"]:
            self.assertIn("laugh", classify(normalize(text)), text)

    def test_emote_without_laugh_marker(self):
        """笑いと無関係な絵文字まで草にしないこと。"""
        for text in ["ajak0nTe3", "ajak0nHi", "PogChamp", "catJAM"]:
            self.assertNotIn("laugh", classify(normalize(text)), text)

    def test_url_is_not_laugh(self):
        self.assertNotIn("laugh", classify(normalize("http://www.example.com")))

    def test_normalize_collapses_runs(self):
        self.assertEqual(normalize("ｗｗｗｗｗｗｗｗ"), "www")
        self.assertEqual(normalize("  Hello   World "), "hello world")


if __name__ == "__main__":
    unittest.main(verbosity=2)
