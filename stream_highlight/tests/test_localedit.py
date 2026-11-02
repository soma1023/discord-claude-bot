# -*- coding: utf-8 -*-
"""手元の動画の自動編集（localedit.py）と、テロップの整形の追加分を確かめる。

文字起こしと動画の読み取りは差し替える。モデルも ffmpeg も使わない。
"""

import json
import os
import shutil
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from stream_highlight import deps, editexport, localedit, subtitles
from stream_highlight.transcribe import Transcript, Word


def settings(**changes):
    flat = localedit.flatten_profile(localedit.load_profile("default"))
    flat.update(changes)
    return localedit.build_settings(flat)


class TestProfiles(unittest.TestCase):
    def test_bundled_profiles_load(self):
        for name in ("default", "murash", "nekomiya", "euler"):
            self.assertIn(name, localedit.available_profiles())
            s = localedit.build_settings(localedit.flatten_profile(localedit.load_profile(name)))
            self.assertEqual(s.silence_method, "speech")
            self.assertEqual(s.engine, "whisper-large-v3")

    def test_profile_overrides_only_what_it_writes(self):
        s = localedit.build_settings(localedit.flatten_profile(localedit.load_profile("nekomiya")))
        self.assertEqual(s.caption.replacements["ねこみや"], "猫宮")
        self.assertEqual(s.caption.line_chars, 16)          # default.yaml のまま
        self.assertEqual(s.gap_cut_sec, 0.4)

    def test_cli_overrides_profile(self):
        args = localedit.make_parser().parse_args(
            ["in.mp4", "--profile", "nekomiya", "--silence-method", "audio",
             "--silence-threshold", "0.08", "--audio-track", "1", "--no-fillers", "--line-chars", "12"])
        s = localedit.settings_from_args(args)
        self.assertEqual((s.silence_method, s.audio_threshold, s.audio_track), ("audio", 0.08, 1))
        self.assertFalse(s.fillers)
        self.assertEqual(s.caption.line_chars, 12)
        self.assertEqual(s.caption.replacements["ねこみや"], "猫宮")

    def test_profile_from_path_and_bad_values(self):
        folder = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, folder, True)
        path = os.path.join(folder, "client.yaml")
        with open(path, "w", encoding="utf-8") as fh:
            fh.write("cut:\n  gap_cut_sec: 0.8\n  audio_track: 2\n")
        s = localedit.build_settings(localedit.flatten_profile(localedit.load_profile(path)))
        self.assertEqual((s.gap_cut_sec, s.audio_track), (0.8, 2))
        with self.assertRaises(localedit.EditError):
            settings(gap_cut_sec=99)
        with self.assertRaises(localedit.EditError):
            settings(silence_method="loud")
        with self.assertRaises(localedit.EditError):
            settings(fillers="yes")
        with self.assertRaises(localedit.EditError):
            localedit.load_profile("no_such_profile")

    def test_prompt_only_for_whisper(self):
        self.assertTrue(settings().prompt)
        self.assertEqual(settings(engine="qwen").prompt, "")
        self.assertEqual(settings(filler_prompt=False).prompt, "")


class TestFillers(unittest.TestCase):
    def test_whole_word_and_spelling_variants(self):
        words = [Word("えー、", 0.0, 0.5), Word("えーーと", 1.0, 1.5), Word("今日は", 2.0, 2.5),
                 Word("あの〜", 3.0, 3.4)]
        found = localedit.filler_ranges(words, ["えー", "えーと", "あのー"])
        self.assertEqual([(a, b) for a, b, _ in found], [(0.0, 0.5), (1.0, 1.5), (3.0, 3.4)])

    def test_reaction_is_kept(self):
        self.assertEqual(localedit.filler_ranges([Word("えー！", 0, 1)], ["えー"]), [])

    def test_prefix_only_for_long_fillers(self):
        found = localedit.filler_ranges([Word("えーとそれで", 0.0, 1.2)], ["えーと"])
        self.assertEqual(len(found), 1)
        self.assertAlmostEqual(found[0][1], 0.55)        # 6文字中3文字 → 0.6秒の手前
        self.assertEqual(localedit.filler_ranges([Word("まあまあ", 0, 1)], ["まあ"]), [])


class TestPlan(unittest.TestCase):
    def setUp(self):
        self.segments = localedit.whole_file(20.0)
        self.words = [Word("今日は", 1.0, 1.5), Word("えー、", 1.6, 2.2), Word("ゲームを", 2.3, 2.9),
                      Word("やります", 3.0, 3.6), Word("よろしく", 6.0, 6.6)]

    def test_gap_and_filler_are_cut(self):
        plan = localedit.build_plan(self.segments, [self.words], [], settings(), 30.0)
        kept = plan.cuts[0]
        # 「えー」の所と、3.6〜6.0 の間が消える
        self.assertTrue(all(not (lo < 1.9 < hi) for lo, hi in kept), kept)
        self.assertTrue(all(not (lo < 4.8 < hi) for lo, hi in kept), kept)
        reasons = [r for item in plan.removed for r in item["reasons"]]
        self.assertTrue(any(r.startswith("フィラー") for r in reasons), reasons)
        self.assertIn("単語の間", reasons)
        texts = [line for cue in plan.cues for line in cue["lines"]]
        self.assertFalse(any("えー" in t for t in texts), texts)

    def test_fillers_off_keeps_them(self):
        plan = localedit.build_plan(self.segments, [self.words], [], settings(fillers=False), 30.0)
        self.assertTrue(any(lo < 1.9 < hi for lo, hi in plan.cuts[0]))

    def test_short_pieces_are_joined(self):
        words = [Word("あ", 1.0, 1.1), Word("はい", 1.9, 2.0), Word("そうです", 2.8, 3.6)]
        plan = localedit.build_plan(self.segments, [words], [], settings(), 30.0)
        self.assertTrue(all(hi - lo >= 0.3 for lo, hi in plan.cuts[0]), plan.cuts)

    def test_audio_silence_adds_cuts(self):
        quiet = [(10.0, 15.0)]
        words = [Word("ずっと", 9.0, 9.5), Word("喋って", 9.6, 16.0)]   # 語の途中に無音
        plan = localedit.build_plan(self.segments, [words], quiet,
                                    settings(silence_method="audio", gap_cut=False), 30.0)
        self.assertTrue(all(not (lo < 12.0 < hi) for lo, hi in plan.cuts[0]), plan.cuts)

    def test_v3_timeline_to_silence(self):
        v3 = {"timebase": "60/1", "v": [[{"offset": 60, "dur": 120}, {"offset": 300, "dur": 60}]]}
        self.assertEqual(localedit.silence_from_v3(v3, 10.0),
                         [(0.0, 1.0), (3.0, 5.0), (6.0, 10.0)])

    def test_segments_file(self):
        folder = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, folder, True)
        path = os.path.join(folder, "s.json")
        with open(path, "w", encoding="utf-8") as fh:
            json.dump([[50, 70], {"start": 10, "end": 20, "peaks": [15]}, [65, 200]], fh)
        got = localedit.read_segments_file(path, 100.0)
        self.assertEqual([(s["start"], s["end"]) for s in got], [(10, 20), (50, 100)])
        self.assertEqual(got[0]["peaks"], [15.0])


class TestCaptionRules(unittest.TestCase):
    def test_pause_beats_particle(self):
        parts = ["今日は", "新しい", "ゲームを", "やって", "いきます"]
        # 「新しい」の後に長い間があれば、助詞の後（今日は|）より間を選ぶ
        lines = subtitles.wrap(parts, 10, [0, 0.05, 0.6, 0.05, 0.05])
        self.assertEqual(lines[0], "今日は新しい")

    def test_particle_beats_plain(self):
        lines = subtitles.wrap(["配信を", "始め", "ます", "ありがとう"], 8)
        self.assertEqual(lines[0], "配信を始めます")
        lines = subtitles.wrap(["ゲームを", "やり", "ます", "ありがとう"], 9, [0, 0, 0, 0])
        self.assertTrue(all(len(l) <= 9 for l in lines), lines)

    def test_never_breaks_inside_word(self):
        parts = ["ホラーゲームを", "やって", "いきます"]
        lines = subtitles.wrap(parts, 10)
        self.assertEqual(lines, ["ホラーゲームを", "やっていきます"])

    def test_no_one_char_line(self):
        """「〜ですよ／ね」のように、1〜2文字だけの行を作らない。"""
        lines = subtitles.wrap(["多分", "そのぐらいだと", "思いますですよ", "ね"], 16)
        self.assertEqual(lines, ["多分そのぐらいだと", "思いますですよね"])
        self.assertEqual(subtitles.wrap(["今日は", "ね"], 16), ["今日はね"])
        # 短い行が前にあるとき、行末の句読点を除くと短くなるとき
        self.assertEqual(subtitles.wrap(["どう", "見られるかとか", "そういうとこまで"], 16),
                         ["どう見られるかとか", "そういうとこまで"])
        self.assertEqual(subtitles.wrap(["初心者デッキおすすめとかで", "検索を", "して、"], 16),
                         ["初心者デッキおすすめとかで", "検索をして、"])
        lines = subtitles.wrap(["あ" * 17])
        self.assertEqual([len(l) for l in lines], [9, 8])

    def test_replacements_across_words(self):
        words = [Word("ねこ", 0.0, 0.3), Word("みやです", 0.3, 0.8), Word("よろしく", 1.0, 1.5)]
        got = subtitles.apply_replacements(words, {"ねこみや": "猫宮"})
        self.assertEqual([w.text for w in got], ["猫宮です", "よろしく"])
        self.assertEqual((got[0].start, got[0].end), (0.0, 0.8))

    def test_strip_trailing_punct(self):
        segments = [{"start": 0.0, "end": 10.0, "peaks": []}]
        timeline = editexport.Timeline(segments, [[(0.0, 10.0)]], 30.0)
        words = [Word("こんにちは。", 1.0, 1.5), Word("今日は、", 3.0, 3.5), Word("雨です", 3.6, 4.0)]
        style = subtitles.CaptionStyle(strip_trailing_punct=True)
        lines = [l for c in subtitles.make_cues(words, timeline, style) for l in c["lines"]]
        self.assertEqual(lines, ["こんにちは", "今日は、雨です"])

    def test_max_show(self):
        segments = [{"start": 0.0, "end": 30.0, "peaks": []}]
        timeline = editexport.Timeline(segments, [[(0.0, 30.0)]], 30.0)
        words = [Word("ずっと", 1.0 + i * 0.5, 1.4 + i * 0.5) for i in range(30)]
        style = subtitles.CaptionStyle(max_show_sec=4.0)
        for cue in subtitles.make_cues(words, timeline, style):
            self.assertLessEqual(cue["end"] - cue["start"], 4.0 + 1e-6)


class TestRun(unittest.TestCase):
    """文字起こしと動画の読み取りを差し替えて、書き出しまで通す。"""

    def setUp(self):
        self.folder = tempfile.mkdtemp(prefix="sh_local_")
        self.addCleanup(shutil.rmtree, self.folder, True)
        self.media = os.path.join(self.folder, "配信 テスト.mp4")
        with open(self.media, "wb") as fh:
            fh.write(b"\0")
        words = [Word("ねこみやです", 1.0, 1.8), Word("えっと、", 2.0, 2.4),
                 Word("今日は", 2.5, 2.9), Word("ホラーゲームを", 3.0, 3.9),
                 Word("やっていきます。", 4.0, 5.0), Word("よろしく", 8.0, 8.6)]
        self.calls = []

        def fake_transcribe(media, start, end, engine, **kwargs):
            self.calls.append(kwargs)
            return Transcript(engine, [w for w in words if start <= w.start < end],
                              elapsed=3.0, device="GPU")

        for target, value in ((localedit.transcribe, "transcribe_segment"),):
            patcher = mock.patch.object(target, value, side_effect=fake_transcribe)
            patcher.start()
            self.addCleanup(patcher.stop)
        patcher = mock.patch("stream_highlight.realtest.probe", return_value=(10.0, 59.94, 1920, 1080))
        patcher.start()
        self.addCleanup(patcher.stop)
        patcher = mock.patch.object(deps, "require")
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_outputs(self):
        s = localedit.build_settings(localedit.flatten_profile(localedit.load_profile("nekomiya")))
        files, summary = localedit.run(self.media, s, progress=lambda *a: None)
        self.assertEqual(os.path.dirname(files["xml"]), os.path.join(self.folder, "配信 テスト"))
        self.assertTrue(self.calls[0]["prompt"])           # フィラーを含む initial_prompt
        self.assertIsNone(self.calls[0]["audio_track"])    # 1本目は ffmpeg の既定のまま

        with open(files["xml"], encoding="utf-8") as fh:
            root = ET.fromstring(fh.read().split("\n", 2)[2])
        clips = root.findall("./sequence/media/video/track/clipitem")
        self.assertEqual(len(clips), summary["clips"])
        self.assertGreater(len(clips), 1)
        # 1カット＝1クリップで隙間なく並び、素材ファイル全体の長さを持つ（端を引き出して戻せる）
        position = 0
        for clip in clips:
            self.assertEqual(int(clip.find("start").text), position)
            position = int(clip.find("end").text)
        self.assertEqual(root.find("./sequence/media/video/track/clipitem/file/duration").text,
                         str(editexport._frames(10.0, 59.94)))
        self.assertEqual(root.find("./sequence/rate/ntsc").text, "TRUE")

        with open(files["srt"], encoding="utf-8") as fh:
            srt = fh.read()
        self.assertIn("猫宮です", srt)
        self.assertNotIn("えっと", srt)
        self.assertNotIn("。", srt)

        with open(files["json"], encoding="utf-8") as fh:
            payload = json.load(fh)
        self.assertEqual(payload["settings"]["engine"], "whisper-large-v3")
        self.assertIn("ねこみやです", payload["text"])
        self.assertTrue(payload["removed"])
        self.assertEqual(payload["effects"], [])
        self.assertLess(summary["edited_sec"], summary["original_sec"])

    def test_missing_file(self):
        with self.assertRaises(localedit.EditError):
            localedit.run(os.path.join(self.folder, "無い.mp4"), settings(), progress=lambda *a: None)


class TestDeps(unittest.TestCase):
    def test_reports_missing_with_hint(self):
        with mock.patch.object(deps.transcribe, "_has", return_value=False), \
                mock.patch.object(deps, "auto_editor_command", return_value=[]):
            with self.assertRaises(deps.MissingDependency) as caught:
                deps.require("whisper-large-v3", "audio")
        text = str(caught.exception)
        self.assertIn("pip install faster-whisper", text)
        self.assertIn("pip install auto-editor", text)

    def test_auto_editor_not_needed_for_speech(self):
        with mock.patch.object(deps, "auto_editor_command", return_value=[]):
            names = [name for name, _ in deps.problems("whisper-large-v3", "speech")]
        self.assertNotIn("auto-editor", names)


if __name__ == "__main__":
    unittest.main()
