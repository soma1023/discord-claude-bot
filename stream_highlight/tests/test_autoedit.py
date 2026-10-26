# -*- coding: utf-8 -*-
"""自動カットとテロップを確かめる。

文字起こしは合成データで代用する。モデルもネットワークも使わない。
"""

import os
import random
import re
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from stream_highlight import autocut, autoedit, cache, editexport, retake, subtitles, transcribe
from stream_highlight.autocut import CutSettings, plan_cuts
from stream_highlight.editexport import Timeline, build_segments, export, to_fcp7_xml
from stream_highlight.sources import StreamInfo
from stream_highlight.transcribe import Transcript, Word

DATA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
INFO = StreamInfo("twitch", "123456", "https://www.twitch.tv/videos/123456",
                  title="テスト配信", duration=9000.0, fps=59.94, width=2560, height=1440)


def parse(xml_text):
    return ET.fromstring(xml_text.split("\n", 2)[2])


def talk(start, phrases, word_sec=0.4, gap=0.1):
    """(テキスト, その前の沈黙秒) の並びから、単語の列を作る。"""
    words = []
    t = start
    for text, silence in phrases:
        t += silence
        for part in text.split():
            words.append(Word(part, t, t + word_sec))
            t += word_sec + gap
    return words


class TestCutPlanning(unittest.TestCase):
    def test_long_silence_is_removed_with_padding(self):
        words = [Word("あ", 10.0, 11.0), Word("い", 14.0, 15.0)]
        kept = plan_cuts(10.0, 15.0, words, peaks=[], settings=CutSettings(protect_sec=0))
        self.assertEqual(kept, [(10.0, 11.15), (13.85, 15.0)])

    def test_short_pause_is_kept(self):
        words = [Word("あ", 10.0, 11.0), Word("い", 12.4, 13.0)]
        kept = plan_cuts(10.0, 13.0, words, peaks=[], settings=CutSettings(protect_sec=0))
        self.assertEqual(kept, [(10.0, 13.0)])

    def test_leading_and_trailing_silence(self):
        words = [Word("あ", 20.0, 21.0)]
        kept = plan_cuts(10.0, 30.0, words, peaks=[], settings=CutSettings(protect_sec=0))
        self.assertEqual(kept, [(19.85, 21.15)])

    def test_peak_window_is_never_cut(self):
        """ピークの前後は、喋っていなくても残す。"""
        words = [Word("あ", 0.0, 1.0), Word("い", 40.0, 41.0)]
        settings = CutSettings(protect_sec=5.0)
        kept = plan_cuts(0.0, 41.0, words, peaks=[20.0], settings=settings)
        covered = lambda t: any(lo <= t <= hi for lo, hi in kept)
        for t in [15.0, 17.5, 20.0, 22.5, 25.0]:
            self.assertTrue(covered(t), (t, kept))
        # 保護範囲の外の沈黙は詰めている
        self.assertFalse(covered(10.0))
        self.assertFalse(covered(30.0))

    def test_peak_window_beats_retake_removal(self):
        words = talk(0.0, [("あ い う え お か き く け こ", 0)])
        kept = plan_cuts(0.0, 5.0, words, peaks=[2.0], settings=CutSettings(protect_sec=1.0),
                         extra_removals=[(1.5, 2.5)])
        self.assertEqual(kept, [(0.0, 5.0)])

    def test_tiny_piece_is_merged(self):
        """0.5秒未満の断片は、消さずに隣とまとめる。"""
        words = [Word("あ", 0.0, 1.0), Word("い", 3.0, 3.1), Word("う", 6.0, 7.0)]
        settings = CutSettings(protect_sec=0, pad_sec=0.05, min_piece_sec=0.5)
        kept = plan_cuts(0.0, 7.0, words, peaks=[], settings=settings)
        for lo, hi in kept:
            self.assertGreaterEqual(hi - lo, 0.5, kept)
        # 「い」の発話は残っている
        self.assertTrue(any(lo <= 3.0 and 3.1 <= hi for lo, hi in kept), kept)

    def test_no_words_means_no_cut(self):
        self.assertEqual(plan_cuts(0.0, 60.0, [], peaks=[]), [(0.0, 60.0)])

    def test_settings_are_clamped(self):
        s = CutSettings.from_values("abc", -1, 999)
        self.assertEqual(s.gap_sec, autocut.DEFAULT_GAP_SEC)
        self.assertEqual(s.pad_sec, 0.0)
        self.assertEqual(s.protect_sec, 120.0)


class TestTimelineAndXml(unittest.TestCase):
    """カット後の in/out とシーケンス時刻の整合。"""

    def setUp(self):
        self.segments = build_segments([1800.0, 5000.0], margin_sec=60, duration=9000)
        rng = random.Random(7)
        self.cuts = []
        for seg in self.segments:
            words, t = [], seg["start"]
            while t < seg["end"] - 1:
                length = rng.uniform(0.2, 1.2)
                words.append(Word("語", t, min(seg["end"], t + length)))
                t += length + rng.choice([0.1, 0.3, 2.5, 4.0])
            self.cuts.append(plan_cuts(seg["start"], seg["end"], words, seg["peaks"]))

    def check_contiguous(self, fps):
        root = parse(to_fcp7_xml(INFO, self.segments, fps=fps, cuts=self.cuts))
        for path in ("sequence/media/video/track", "sequence/media/audio/track"):
            for track in root.findall(path):
                position = 0
                last_out = None
                for item in track.findall("clipitem"):
                    start, end = int(item.find("start").text), int(item.find("end").text)
                    src_in, src_out = int(item.find("in").text), int(item.find("out").text)
                    self.assertEqual(start, position, "隙間か重なりがある")
                    self.assertEqual(end - start, src_out - src_in)
                    self.assertGreater(end, start)
                    if last_out is not None:
                        self.assertGreaterEqual(src_in, last_out, "元の配信で逆戻りしている")
                    position, last_out = end, src_out
                self.assertEqual(int(root.find("sequence/duration").text), position)
        return root

    def test_contiguous_at_common_rates(self):
        for fps in (30, 60, 29.97, 59.94):
            with self.subTest(fps=fps):
                self.check_contiguous(fps)

    def test_one_segment_becomes_many_clips(self):
        root = self.check_contiguous(60)
        items = root.findall("sequence/media/video/track/clipitem")
        self.assertGreater(len(items), len(self.segments))

    def test_clip_names_follow_the_rule(self):
        root = self.check_contiguous(59.94)
        names = [i.find("name").text for i in root.findall("sequence/media/video/track/clipitem")]
        self.assertTrue(all(re.match(r"^\d{2}-\d{2} \d{2}:\d{2}:\d{2}:\d{2}$", n) for n in names),
                        names[:5])
        self.assertTrue(names[0].startswith("01-01 "))

    def test_clip_ids_are_unique(self):
        root = self.check_contiguous(60)
        ids = [i.get("id") for i in root.iter("clipitem")]
        self.assertEqual(len(ids), len(set(ids)))

    def test_markers_follow_the_cut(self):
        fps = 59.94
        root = parse(to_fcp7_xml(INFO, self.segments, fps=fps, cuts=self.cuts))
        timeline = Timeline(self.segments, self.cuts, fps)
        markers = root.findall("sequence/marker")
        self.assertEqual(len(markers), 2)
        for marker, seg in zip(markers, self.segments):
            peak = seg["peaks"][0]
            frame = int(marker.find("in").text)
            self.assertEqual(frame, timeline.to_sequence(peak))
            # ピークの位置の映像は、元の配信のピークそのもの
            clip = timeline.clips[timeline.locate(peak)]
            self.assertEqual(clip["src_in"] + frame - clip["seq_start"],
                             editexport._frames(peak, fps))

    def test_edl_with_cuts_is_contiguous(self):
        _, content, _ = export(INFO, [1800.0, 5000.0], fmt="edl", margin_sec=60, fps=30,
                               cuts=self.cuts)
        events = [l for l in content.splitlines() if re.match(r"^\d{3}  AX       V ", l)]
        self.assertEqual(len(events), len(Timeline(self.segments, self.cuts, 30).clips))
        previous_out = "00:00:00:00"
        for line in events:
            rec_in, rec_out = line.split()[-2:]
            self.assertEqual(rec_in, previous_out)
            previous_out = rec_out


class TestUnchangedWhenOff(unittest.TestCase):
    """機能を使わないときは、変更前と1文字も変わらないこと。

    tests/data/legacy_* は変更前のコードで書き出したもの。
    """

    CASES = {
        "legacy_60.xml": dict(fmt="xml", fps=60, media_path=r"C:\Users\me\Videos\配信 本編.mp4"),
        "legacy_2997.xml": dict(fmt="xml", fps=29.97),
        "legacy_5994.xml": dict(fmt="xml", fps=59.94, width=1920, height=1080,
                                media_path="D:/v.mp4"),
        "legacy_2997.edl": dict(fmt="edl", fps=29.97, media_path=r"C:\v.mp4"),
    }

    def test_identical_to_previous_output(self):
        for name, kwargs in self.CASES.items():
            with self.subTest(name=name):
                with open(os.path.join(DATA, name), encoding="utf-8", newline="") as fh:
                    expected = fh.read()
                _, content, _ = export(INFO, [1800.5, 1900.25, 5000.1, 8990],
                                       margin_sec=300, **kwargs)
                self.assertEqual(content, expected)

    def test_subtitle_only_timeline_matches_legacy_xml(self):
        """テロップだけのときは、従来の XML とフレーム単位で同じ並びになること。"""
        segments = build_segments([1800.5, 5000.1], 300, 9000)
        root = parse(to_fcp7_xml(INFO, segments, fps=29.97))
        timeline = Timeline(segments, None, 29.97)
        items = root.findall("sequence/media/video/track/clipitem")
        self.assertEqual(len(items), len(timeline.clips))
        for item, clip in zip(items, timeline.clips):
            self.assertEqual(int(item.find("start").text), clip["seq_start"])
            self.assertEqual(int(item.find("in").text), clip["src_in"])
            self.assertEqual(int(item.find("out").text), clip["src_out"])


class TestLoopRemoval(unittest.TestCase):
    def texts(self, words):
        return [w.text for w in words]

    def test_repeated_phrase_collapses(self):
        words = talk(0, [("ありがとう ございます " * 5, 0)])
        out = transcribe.remove_loops(words)
        self.assertEqual(self.texts(out), ["ありがとう", "ございます"])

    def test_single_repeated_word(self):
        words = talk(0, [("今日は すごい すごい すごい すごい 日", 0)])
        self.assertEqual(self.texts(transcribe.remove_loops(words)), ["今日は", "すごい", "日"])

    def test_two_repeats_are_kept(self):
        words = talk(0, [("やばい やばい 来た", 0)])
        self.assertEqual(self.texts(transcribe.remove_loops(words)), ["やばい", "やばい", "来た"])

    def test_laughter_is_kept(self):
        """1文字の繰り返し（笑い）は本物の発話なので消さない。"""
        words = talk(0, [("は は は は 草", 0)])
        self.assertEqual(len(transcribe.remove_loops(words)), 5)

    def test_punctuation_differences_still_match(self):
        words = [Word("おはよう。", 0, 1), Word("おはよう", 1, 2), Word("おはよう、", 2, 3),
                 Word("おはよう", 3, 4)]
        self.assertEqual(self.texts(transcribe.remove_loops(words)), ["おはよう。"])

    def test_clean_words_fixes_order_and_blanks(self):
        words = [Word(" ", 0, 1), Word("後", 5, 4), Word("前", 1, 2)]
        out = transcribe.clean_words(words)
        self.assertEqual(self.texts(out), ["前", "後"])
        self.assertEqual((out[1].start, out[1].end), (4, 5))


class TestJapaneseGrouping(unittest.TestCase):
    """Qwen の1文字ずつ・句読点なしの出力を、テロップに使える単位に直す。"""

    def chars(self, text, start=0.0, step=0.1, pauses=()):
        out, t = [], start
        for i, ch in enumerate(text):
            if i in pauses:
                t += 1.0
            out.append(Word(ch, t, t + step * 0.8))
            t += step
        return out

    def test_punctuation_is_restored(self):
        words = transcribe.attach_punctuation("こんにちは、今日は。", self.chars("こんにちは今日は"))
        self.assertEqual([w.text for w in words][4], "は、")
        self.assertEqual(words[-1].text, "は。")

    def test_unmatched_text_is_left_alone(self):
        words = transcribe.attach_punctuation("全然違う", self.chars("あい"))
        self.assertEqual([w.text for w in words], ["あ", "い"])

    def test_grouped_like_bunsetsu(self):
        text = "今日はホラーゲームをやっていきます。"
        words = transcribe.group_japanese(
            transcribe.attach_punctuation(text, self.chars(text.rstrip("。"))))
        self.assertEqual([w.text for w in words], ["今日は", "ホラーゲームを", "やっていきます。"])
        self.assertAlmostEqual(words[0].start, 0.0)

    def test_pause_splits_a_word(self):
        """語の中に沈黙を隠さない（隠すと詰められなくなる）。"""
        words = transcribe.group_japanese(self.chars("えーとちょっと", pauses=(3,)))
        self.assertEqual([w.text for w in words], ["えーと", "ちょっと"])

    def test_sentence_end_splits(self):
        words = transcribe.group_japanese([Word("す。", 0, 0.1), Word("よ", 0.1, 0.2)])
        self.assertEqual(len(words), 2)


class TestChunking(unittest.TestCase):
    """5分を超える区間を分けて処理したときの、時刻のつなぎ直し。"""

    def test_split_points_respect_limit_and_pick_quiet_spot(self):
        frame = 0.05
        total = 800.0
        energies = [1.0] * int(total / frame)
        energies[int(280.0 / frame)] = 0.0    # 語の中の一瞬の途切れ（選ばれてはいけない）
        for k in range(int(269.5 / frame), int(270.5 / frame)):
            energies[k] = 0.05                # 上限290秒の手前の、しばらく静かな所
        points = transcribe.choose_split_points(energies, frame, total, max_chunk=290)
        self.assertAlmostEqual(points[0], 270.0, delta=0.3)
        bounds = [0.0] + points + [total]
        for lo, hi in zip(bounds, bounds[1:]):
            self.assertLessEqual(hi - lo, 290.0 + 1e-9)
            self.assertGreater(hi - lo, 0)

    def test_chunk_length_follows_free_vram(self):
        """VRAM の空きが少ないほど短く区切る（溢れると2倍近く遅くなるため）。"""
        self.assertEqual(transcribe.chunk_sec_for_vram(8.0), transcribe.QWEN_MAX_CHUNK_SEC)
        self.assertEqual(transcribe.chunk_sec_for_vram(0.1), transcribe.QWEN_MIN_CHUNK_SEC)
        middle = transcribe.chunk_sec_for_vram(1.5)
        self.assertTrue(60 < middle < 80, middle)

    def test_whisper_word_ends_are_extended_but_not_overlapping(self):
        words = [Word("あ", 0.0, 1.0), Word("い", 1.05, 2.0), Word("う", 5.0, 6.0)]
        out = transcribe.extend_word_ends(words, 0.15)
        self.assertEqual([round(w.end, 2) for w in out], [1.05, 2.15, 6.15])

    def test_short_audio_is_not_split(self):
        self.assertEqual(transcribe.choose_split_points([1.0] * 100, 0.05, 150.0), [])

    def test_merge_restores_offsets(self):
        chunks = [(0.0, 270.0, [Word("あ", 1.0, 2.0), Word("い", 269.5, 271.0)]),
                  (270.0, 280.0, [Word("う", 0.2, 0.8), Word("え", -0.5, 0.1)])]
        out = transcribe.merge_chunks(chunks)
        self.assertEqual([(w.text, w.start, w.end) for w in out],
                         [("あ", 1.0, 2.0), ("い", 269.5, 270.0), ("え", 270.0, 270.1),
                          ("う", 270.2, 270.8)])

    def test_run_engine_splits_long_audio_and_stitches(self):
        """分割処理を通しても、元の時刻で単語が返ること。"""
        import numpy as np

        rate = transcribe.SAMPLE_RATE
        total = 700.0
        samples = np.ones(int(total * rate), dtype=np.float32) * 0.5
        for quiet_at in (250.0, 520.0):          # 静かな所を作っておく
            samples[int(quiet_at * rate):int((quiet_at + 0.5) * rate)] = 0.0

        seen = []

        class FakeEngine(transcribe.Engine):
            name = "fake"
            max_chunk_sec = 290.0
            device = "CPU"

            def run(self, piece):
                length = len(piece) / float(rate)
                seen.append(length)
                self.assert_ok = length <= 290.0
                # 各断片の先頭から10秒ごとに1語
                return [Word("w", t, t + 0.5) for t in np.arange(0.0, length - 0.5, 10.0)]

        words = transcribe.run_engine(FakeEngine(), samples)
        self.assertTrue(all(length <= 290.0 for length in seen), seen)
        self.assertEqual(len(seen), 3)
        self.assertAlmostEqual(sum(seen), total, places=2)
        # 時刻は単調で、全体の範囲に収まっている
        starts = [w.start for w in words]
        self.assertEqual(starts, sorted(starts))
        self.assertLessEqual(words[-1].end, total)
        # 2つ目の断片の最初の語は、1つ目の断片の長さだけずれた位置にある
        second_start = seen[0]
        self.assertTrue(any(abs(w.start - second_start) < 1e-6 for w in words))


class TestSubtitles(unittest.TestCase):
    def test_wrap_limits(self):
        parts = ["今日は", "ホラーゲームを", "やっていきます、", "よろしく", "お願いします"]
        lines = subtitles.wrap(parts)
        self.assertTrue(all(len(l) <= 16 for l in lines), lines)
        self.assertEqual("".join(lines), "".join(parts))

    def test_wrap_prefers_punctuation(self):
        lines = subtitles.wrap(["いやー", "怖かった、", "本当に", "怖かったです"])
        self.assertEqual(lines[0], "いやー怖かった、")

    def test_long_word_is_split(self):
        lines = subtitles.wrap(["あ" * 40])
        self.assertEqual([len(l) for l in lines], [16, 16, 8])

    def test_srt_time_format(self):
        self.assertEqual(subtitles._srt_time(3725.5), "01:02:05,500")

    def build(self, fps=29.97):
        segments = build_segments([100.0], margin_sec=20, duration=9000)   # 80〜120秒
        words = (talk(81.0, [("こんにちは 今日は", 0)])
                 + talk(90.0, [("ホラーゲームを やります", 0)])
                 + talk(100.5, [("うわっ びっくりした", 0)])
                 + talk(110.0, [("怖かった です", 0)]))
        cuts = [plan_cuts(80.0, 120.0, words, [100.0], CutSettings(protect_sec=2.0))]
        timeline = Timeline(segments, cuts, fps)
        return words, timeline, subtitles.make_cues(words, timeline)

    def test_cue_times_match_the_sequence(self):
        """SRT の時刻が、カット後のシーケンス上の位置と一致すること。"""
        fps = 29.97
        words, timeline, cues = self.build(fps)
        self.assertEqual(len(cues), 4)
        for cue, first in zip(cues, [words[0], words[2], words[4], words[6]]):
            expected = timeline.to_sequence(first.start) / fps
            self.assertAlmostEqual(cue["start"], expected, places=6)
        # 元の配信の時刻（81秒など）ではない
        self.assertLess(cues[0]["start"], 2.0)
        srt = subtitles.to_srt(cues)
        first_line = srt.splitlines()[1]
        self.assertEqual(first_line.split(" --> ")[0],
                         subtitles._srt_time(timeline.to_sequence(words[0].start) / fps))

    def test_cues_do_not_overlap_and_respect_duration(self):
        _, _, cues = self.build()
        for a, b in zip(cues, cues[1:]):
            self.assertLessEqual(a["end"], b["start"] + 1e-9)
        for cue in cues:
            self.assertLessEqual(cue["end"] - cue["start"], subtitles.MAX_SHOW_SEC + 1e-9)
            self.assertLessEqual(len(cue["lines"]), 2)

    def test_short_cue_is_extended(self):
        segments = build_segments([100.0], margin_sec=20, duration=9000)
        timeline = Timeline(segments, None, 30)
        cues = subtitles.make_cues([Word("あ", 90.0, 90.1)], timeline)
        self.assertAlmostEqual(cues[0]["end"] - cues[0]["start"], subtitles.MIN_SHOW_SEC)

    def test_cue_is_split_at_cut_point(self):
        """発話がカット点をまたぐときは、そこでテロップを分ける。"""
        segments = build_segments([100.0], margin_sec=20, duration=9000)
        # 「残す部分」を2つに分け、その境目をまたいで喋らせる
        cuts = [[(80.0, 95.0), (97.0, 120.0)]]
        timeline = Timeline(segments, cuts, 30)
        words = [Word("前半", 94.0, 94.6), Word("後半", 97.1, 97.6)]
        cues = subtitles.make_cues(words, timeline)
        self.assertEqual([c["lines"] for c in cues], [["前半"], ["後半"]])
        self.assertLessEqual(cues[0]["end"], cues[1]["start"])

    def test_cut_words_are_dropped(self):
        segments = build_segments([100.0], margin_sec=20, duration=9000)
        timeline = Timeline(segments, [[(80.0, 95.0), (97.0, 120.0)]], 30)
        cues = subtitles.make_cues([Word("消えた", 95.5, 96.5)], timeline)
        self.assertEqual(cues, [])


class TestRetake(unittest.TestCase):
    def setUp(self):
        self.words = talk(0, [("今日は ホラゲ 今日は ホラーゲームを やります "
                               "最後まで 見て いって ください ね", 0)])

    def test_transcript_has_indices(self):
        text = retake.format_transcript(self.words)
        self.assertIn("[0]今日は", text)
        self.assertIn("[4]やります", text)

    def test_answer_becomes_time_ranges(self):
        ranges = retake.ranges_from_answer(self.words, {"remove": [{"first": 0, "last": 1}]})
        self.assertEqual(ranges, [(self.words[0].start, self.words[1].end)])

    def test_bad_indices_are_ignored(self):
        self.assertEqual(retake.ranges_from_answer(self.words, {"remove": [{"first": 9, "last": 12}]}), [])

    def test_too_many_removals_rejected(self):
        with self.assertRaises(retake.RetakeError):
            retake.ranges_from_answer(self.words, {"remove": [{"first": 0, "last": 5}]})

    def test_api_failure_does_not_stop_export(self):
        """言い直しの判定に失敗しても、書き出しは続けること。"""
        transcript = Transcript("qwen", talk(1790.0, [("あ い う", 0), ("え お", 5)]))
        options = autoedit.EditOptions(autocut=True, retake=True, engine="qwen")
        with mock.patch.object(transcribe, "transcribe_segment", return_value=transcript), \
                mock.patch.object(retake, "detect", side_effect=retake.RetakeError("テスト")):
            result = autoedit.run(INFO, [1800.0], options, margin_sec=30, fps=30,
                                  media_path="dummy.mp4")
        self.assertIn("xmeml", result["content"])
        self.assertTrue(any("言い直しの削除は行いませんでした" in n for n in result["notes"]))

    def test_env_file_is_read_but_not_partner_file(self):
        tmp = tempfile.mkdtemp(prefix="sh_env_")
        env = os.path.join(tmp, ".env")
        with open(env, "w", encoding="utf-8") as fh:
            fh.write("# comment\nSTREAM_HIGHLIGHT_RETAKE_MODEL='claude-test'\n")
        with open(os.path.join(tmp, ".env.for-partner"), "w", encoding="utf-8") as fh:
            fh.write("ANTHROPIC_API_KEY=partner\n")
        with mock.patch.object(retake, "_env_files", return_value=[env]), \
                mock.patch.dict(os.environ, {}, clear=False):
            os.environ.pop("STREAM_HIGHLIGHT_RETAKE_MODEL", None)
            os.environ.pop("ANTHROPIC_API_KEY", None)
            self.assertEqual(retake.model_name(), "claude-test")
            self.assertEqual(retake.setting("ANTHROPIC_API_KEY"), "")


class TestAutoEditRun(unittest.TestCase):
    """区間ごとの文字起こし→カット→書き出しの通し（文字起こしは合成）。"""

    def fake_transcribe(self, media_path, start, end, engine, report=None, use_cache=True):
        words = []
        t = start + 1.0
        while t < end - 1.0:
            words.append(Word("話す", t, t + 0.8))
            t += 0.9 if int(t) % 20 else 6.0        # ときどき長い沈黙を入れる
        return Transcript(engine, words, device="CPU")

    def run_auto(self, **kw):
        options = autoedit.EditOptions(**kw)
        with mock.patch.object(transcribe, "transcribe_segment", side_effect=self.fake_transcribe):
            return autoedit.run(INFO, [1800.0, 5000.0], options, margin_sec=120, fps=59.94,
                                media_path="dummy.mp4")

    def test_autocut_shortens_and_writes_srt(self):
        result = self.run_auto(autocut=True, subtitles=True)
        self.assertLess(result["edited_sec"], result["original_sec"])
        self.assertTrue(result["srt_filename"].endswith(".srt"))
        self.assertEqual(os.path.splitext(result["srt_filename"])[0],
                         os.path.splitext(result["filename"])[0])
        self.assertIn(" --> ", result["srt_content"])

    def test_subtitles_only_keeps_legacy_xml(self):
        result = self.run_auto(subtitles=True)
        _, legacy, _ = export(INFO, [1800.0, 5000.0], margin_sec=120, fps=59.94,
                              media_path="dummy.mp4")
        self.assertEqual(result["content"], legacy)

    def test_progress_is_reported_per_segment(self):
        messages = []
        options = autoedit.EditOptions(subtitles=True)
        with mock.patch.object(transcribe, "transcribe_segment", side_effect=self.fake_transcribe):
            autoedit.run(INFO, [1800.0, 5000.0], options, margin_sec=120, fps=30,
                         media_path="dummy.mp4", report=lambda m, f=None: messages.append(m))
        self.assertIn("区間 1/2 を文字起こし中…", messages)
        self.assertIn("区間 2/2 を文字起こし中…", messages)


class TestSegmentMedia(unittest.TestCase):
    """区間付近だけをダウンロードし、区間ごとに別の動画ファイルを使うとき。"""

    PEAKS = [1800.0, 5000.0]

    def setUp(self):
        self.calls = []
        self.segments = build_segments(self.PEAKS, 120, INFO.duration)
        # ダウンロードは区間の3秒前から始まった、という想定
        self.media = [{"path": "C:\\dl\\part%d.mp4" % (i + 1), "offset": s["start"] - 3.0,
                       "duration": s["end"] - s["start"] + 6.0}
                      for i, s in enumerate(self.segments)]

    def fake_transcribe(self, media_path, start, end, engine, report=None, use_cache=True):
        self.calls.append((media_path, start, end))
        offset = 0.0
        for media in self.media:
            if media_path == media["path"]:
                offset = media["offset"]
        # 元の配信の時刻で決まる発話を、ファイルの中の時刻で返す
        found = TestAutoEditRun.fake_transcribe(self, media_path, start + offset, end + offset,
                                                engine)
        found.words = [Word(w.text, w.start - offset, w.end - offset) for w in found.words]
        return found

    def run_auto(self, segment_media):
        options = autoedit.EditOptions(autocut=True, subtitles=True)
        with mock.patch.object(transcribe, "transcribe_segment", side_effect=self.fake_transcribe):
            return autoedit.run(INFO, self.PEAKS, options, margin_sec=120, fps=59.94,
                                media_path="dummy.mp4", segment_media=segment_media)

    def test_each_segment_uses_its_own_file(self):
        result = self.run_auto(self.media)
        # 文字起こしはファイルの中の時刻（区間は先頭から3秒の所）で頼む
        self.assertEqual([(p, round(s, 3)) for p, s, _ in self.calls],
                         [("C:\\dl\\part1.mp4", 3.0), ("C:\\dl\\part2.mp4", 3.0)])
        root = parse(result["content"])
        files = {f.get("id"): f for f in root.iter("file") if f.find("pathurl") is not None}
        self.assertEqual(sorted(files), ["media-1", "media-2"])
        self.assertEqual(files["media-2"].find("pathurl").text,
                         "file://localhost/C:/dl/part2.mp4")
        fps = 59.94
        for item in root.findall("sequence/media/video/track/clipitem"):
            index = int(item.get("id").split("-")[1]) - 1
            src_in = int(item.find("in").text)
            # in はファイル先頭から。名前は元の配信のタイムコードのまま
            name_tc = item.find("name").text.split()[1]
            offset = editexport._frames(self.media[index]["offset"], fps)
            self.assertEqual(editexport.timecode((src_in + offset) / fps, fps), name_tc)
            self.assertEqual(item.find("file").get("id"), "media-%d" % (index + 1))

    def test_captions_match_single_file_export(self):
        """ファイルが分かれても、カットとテロップは配信全体の動画のときと同じ。"""
        split = self.run_auto(self.media)
        self.media = []
        whole = self.run_auto(None)
        self.assertEqual(split["srt_content"], whole["srt_content"])
        self.assertEqual(split["edited_sec"], whole["edited_sec"])

    def test_count_mismatch_is_rejected(self):
        with self.assertRaises(ValueError):
            self.run_auto(self.media[:1])


class TestServerFlow(unittest.TestCase):
    """画面から呼ぶ API。ON のときだけジョブになり、OFF ならその場で返る。"""

    @classmethod
    def setUpClass(cls):
        import time
        from fastapi.testclient import TestClient

        from stream_highlight import server
        from stream_highlight.jobs import manager

        cls.time = time
        cls.tmp = tempfile.mkdtemp(prefix="sh_srv_")
        cls.original = cache.CACHE_DIR
        cache.CACHE_DIR = cls.tmp
        cls.media = os.path.join(cls.tmp, "vod.mp4")
        with open(cls.media, "wb") as fh:
            fh.write(b"0")
        info = StreamInfo("youtube", "auto1234567", "https://www.youtube.com/watch?v=auto1234567",
                          title="自動テスト", duration=4000.0)
        cache.save(info, [])
        manager._loaded.clear()
        cls.client = TestClient(server.app)
        cls.key = "youtube:auto1234567"

    @classmethod
    def tearDownClass(cls):
        import shutil
        cache.CACHE_DIR = cls.original
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def post(self, **extra):
        body = {"video_key": self.key, "peaks": [1000.0, 2500.0], "margin_sec": 60, "fps": 30,
                "media_path": self.media}
        body.update(extra)
        return self.client.post("/api/export/edit", json=body)

    def test_off_returns_immediately(self):
        res = self.post()
        self.assertEqual(res.status_code, 200, res.text)
        self.assertNotIn("job_id", res.json())
        self.assertIn("xmeml", res.json()["content"])

    def test_missing_media_is_rejected_up_front(self):
        res = self.post(autocut=True, media_path="")
        self.assertEqual(res.status_code, 400)
        self.assertIn("動画ファイルの場所", res.json()["detail"])

    def test_missing_dependency_is_rejected_up_front(self):
        with mock.patch.object(transcribe, "missing", return_value=["torch が入っていません"]):
            res = self.post(subtitles=True)
        self.assertEqual(res.status_code, 400)
        self.assertIn("torch", res.json()["detail"])

    def test_job_runs_and_returns_xml_and_srt(self):
        def fake(media_path, start, end, engine, report=None, use_cache=True):
            return Transcript(engine, talk(start + 1, [("こんにちは 今日は", 0), ("よろしく", 5)]))

        with mock.patch.object(transcribe, "missing", return_value=[]), \
                mock.patch.object(transcribe, "transcribe_segment", side_effect=fake):
            res = self.post(autocut=True, subtitles=True, engine="whisper")
            self.assertEqual(res.status_code, 200, res.text)
            job_id = res.json()["job_id"]
            for _ in range(100):
                status = self.client.get("/api/job/%s" % job_id).json()
                if status["status"] in ("done", "error"):
                    break
                self.time.sleep(0.05)
        self.assertEqual(status["status"], "done", status.get("error"))
        result = status["result"]
        self.assertIn("xmeml", result["content"])
        self.assertIn(" --> ", result["srt_content"])
        self.assertEqual(result["engine"], "whisper")

    def test_capabilities(self):
        body = self.client.get("/api/autoedit/capabilities").json()
        self.assertEqual([e["id"] for e in body["engines"]], ["qwen", "whisper"])
        self.assertIn("retake", body)


class TestTranscriptCache(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="sh_tr_")
        self.original = cache.CACHE_DIR
        cache.CACHE_DIR = self.tmp
        self.media = os.path.join(self.tmp, "v.mp4")
        with open(self.media, "wb") as fh:
            fh.write(b"0" * 10)

    def tearDown(self):
        import shutil
        cache.CACHE_DIR = self.original
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_cache_is_per_engine_and_skips_transcription(self):
        raw = Transcript("qwen", [Word("あり", 10, 11), Word("あり", 11, 12), Word("あり", 12, 13)])
        transcribe.save_cached(self.media, 0.0, 60.0, raw)
        self.assertIsNone(transcribe.load_cached(self.media, 0.0, 60.0, "whisper"))
        with mock.patch.object(transcribe, "extract_audio",
                               side_effect=AssertionError("文字起こしし直している")):
            found = transcribe.transcribe_segment(self.media, 0.0, 60.0, "qwen")
        # キャッシュには生の出力、返すときにループ除去をかける
        self.assertEqual([w.text for w in found.words], ["あり"])

    def test_changed_file_invalidates_cache(self):
        transcribe.save_cached(self.media, 0.0, 60.0, Transcript("qwen", [Word("あ", 1, 2)]))
        with open(self.media, "ab") as fh:
            fh.write(b"more")
        self.assertIsNone(transcribe.load_cached(self.media, 0.0, 60.0, "qwen"))


class TestDependencies(unittest.TestCase):
    def test_missing_library_is_explained_in_japanese(self):
        with mock.patch.object(transcribe, "_has", return_value=False), \
                mock.patch.object(transcribe, "find_ffmpeg", return_value="ffmpeg"):
            problems = transcribe.missing("qwen")
        self.assertEqual(len(problems), 1)
        self.assertIn("qwen-asr", problems[0])
        self.assertIn("入っていません", problems[0])

    def test_missing_ffmpeg_is_reported(self):
        with mock.patch.object(transcribe, "find_ffmpeg", return_value=""):
            self.assertTrue(any("ffmpeg" in p for p in transcribe.missing("whisper")))

    def test_exe_build_disables_feature(self):
        with mock.patch.object(transcribe.paths, "is_frozen", return_value=True):
            self.assertIn("exe版", transcribe.missing("qwen")[0])

    def test_gpu_oom_falls_back_to_cpu(self):
        """GPUのメモリが足りないときは、落とさずCPUで読み込み直す。"""
        loads = []

        class OutOfMemoryError(RuntimeError):
            pass

        class FakeEngine(transcribe.Engine):
            name = "fake"

            def load(self, prefer_gpu=True):
                loads.append(prefer_gpu)
                if prefer_gpu:
                    raise OutOfMemoryError("CUDA out of memory")
                self.device = "CPU"

            def unload(self):
                pass

        holder = transcribe.EngineHolder()
        messages = []
        with mock.patch.dict(transcribe.ENGINES, {"fake": FakeEngine}), \
                mock.patch.object(transcribe, "missing", return_value=[]):
            engine = holder.get("fake", report=messages.append)
        self.assertEqual(loads, [True, False])
        self.assertEqual(engine.device, "CPU")
        self.assertTrue(any("CPU" in m for m in messages))

    def test_switching_engine_releases_previous(self):
        events = []

        def make(name):
            class Fake(transcribe.Engine):
                def load(self, prefer_gpu=True):
                    events.append("load " + name)

                def unload(self):
                    events.append("unload " + name)
            Fake.name = name
            return Fake

        holder = transcribe.EngineHolder()
        with mock.patch.dict(transcribe.ENGINES, {"a": make("a"), "b": make("b")}), \
                mock.patch.object(transcribe, "missing", return_value=[]):
            holder.get("a")
            holder.get("a")
            holder.get("b")
        self.assertEqual(events, ["load a", "unload a", "load b"])


if __name__ == "__main__":
    unittest.main()
