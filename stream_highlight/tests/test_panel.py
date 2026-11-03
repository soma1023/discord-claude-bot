# -*- coding: utf-8 -*-
"""Premiere のオートカットパネル（premiere_panel/）と、サーバー側の判定 API を確かめる。

Premiere は動かせないので、パネルの時刻の計算（cutplan.js）だけを node で確かめる。
"""

import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from stream_highlight import deps, localedit
from stream_highlight.transcribe import Transcript, Word

PANEL = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "premiere_panel")
WORDS = [Word("今日は", 101.0, 101.5), Word("えっと、", 101.7, 102.1), Word("ゲームを", 102.2, 102.8),
         Word("やります", 102.9, 103.5), Word("よろしく", 106.0, 106.6)]


def fake_transcribe(media, start, end, engine, **kwargs):
    return Transcript(engine, [w for w in WORDS if start <= w.start < end], elapsed=2.0, device="GPU")


class TestAnalyzeRange(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.mkdtemp(prefix="sh_panel_")
        self.addCleanup(shutil.rmtree, self.folder, True)
        self.media = os.path.join(self.folder, "素材.mp4")
        with open(self.media, "wb") as fh:
            fh.write(b"\0")
        for target, name, kwargs in ((localedit.transcribe, "transcribe_segment",
                                      {"side_effect": fake_transcribe}),
                                     (deps, "require", {})):
            patcher = mock.patch.object(target, name, **kwargs)
            patcher.start()
            self.addCleanup(patcher.stop)

    def settings(self):
        return localedit.build_settings(localedit.flatten_profile(localedit.load_profile("default")))

    def test_keep_is_in_source_seconds(self):
        result = localedit.analyze_range(self.media, 100.0, 108.0, self.settings())
        keep = result["keep"]
        self.assertGreaterEqual(keep[0][0], 100.0)
        self.assertLessEqual(keep[-1][1], 108.0)
        # 「えっと」と、3.5〜6.0 の間は残さない
        self.assertTrue(all(not (lo < 101.9 < hi) for lo, hi in keep), keep)
        self.assertTrue(all(not (lo < 104.5 < hi) for lo, hi in keep), keep)
        self.assertLess(result["kept_sec"], result["original_sec"])
        self.assertTrue(any("フィラー" in r for item in result["removed"] for r in item["reasons"]))

    def test_errors(self):
        with self.assertRaises(localedit.EditError):
            localedit.analyze_range(os.path.join(self.folder, "無い.mp4"), 0, 5, self.settings())
        with self.assertRaises(localedit.EditError):
            localedit.analyze_range(self.media, 5.0, 5.05, self.settings())


class TestPanelApi(unittest.TestCase):
    def setUp(self):
        from fastapi.testclient import TestClient

        from stream_highlight import server

        self.client = TestClient(server.app)
        self.folder = tempfile.mkdtemp(prefix="sh_panel_api_")
        self.addCleanup(shutil.rmtree, self.folder, True)
        self.media = os.path.join(self.folder, "素材.mp4")
        with open(self.media, "wb") as fh:
            fh.write(b"\0")
        for target, name, kwargs in ((localedit.transcribe, "transcribe_segment",
                                      {"side_effect": fake_transcribe}),
                                     (deps, "require", {}),
                                     (deps, "problems", {"return_value": []})):
            patcher = mock.patch.object(target, name, **kwargs)
            patcher.start()
            self.addCleanup(patcher.stop)

    def test_profiles(self):
        res = self.client.get("/api/panel/profiles")
        self.assertEqual(res.status_code, 200, res.text)
        names = [p["name"] for p in res.json()["profiles"]]
        self.assertEqual(names, ["default", "euler", "murash", "nekomiya"])
        row = res.json()["profiles"][0]
        self.assertEqual((row["gap_cut_sec"], row["pad_sec"], row["fillers"]), (0.4, 0.15, True))

    def test_analyze_job(self):
        res = self.client.post("/api/panel/analyze", json={
            "media_path": self.media, "source_in": 100.0, "source_out": 108.0,
            "profile": "nekomiya", "gap_cut_sec": 0.5, "fillers": False})
        self.assertEqual(res.status_code, 200, res.text)
        job_id = res.json()["job_id"]
        for _ in range(100):
            job = self.client.get("/api/job/" + job_id).json()
            if job["status"] in ("done", "error"):
                break
            time.sleep(0.05)
        self.assertEqual(job["status"], "done", job)
        keep = job["result"]["keep"]
        # フィラーを OFF にしたので「えっと」は残る
        self.assertTrue(any(lo < 101.9 < hi for lo, hi in keep), keep)

    def test_rejects_unknown_profile_and_paths(self):
        res = self.client.post("/api/panel/analyze", json={
            "media_path": self.media, "source_in": 0, "source_out": 5, "profile": "../secret"})
        self.assertEqual(res.status_code, 400)
        res = self.client.post("/api/panel/analyze", json={
            "media_path": os.path.join(self.folder, "無い.mp4"), "source_in": 0, "source_out": 5})
        self.assertEqual(res.status_code, 400)
        self.assertIn("見つかりません", res.json()["detail"])

    def test_bad_override_is_400(self):
        res = self.client.post("/api/panel/analyze", json={
            "media_path": self.media, "source_in": 0, "source_out": 5, "gap_cut_sec": 99})
        self.assertEqual(res.status_code, 400)


class TestCaptions(unittest.TestCase):
    """パネルの「テロップ作成」: タイムラインに置かれたクリップの位置にテロップを合わせる。"""

    def setUp(self):
        from stream_highlight import cache, transcribe

        self.folder = tempfile.mkdtemp(prefix="sh_caps_")
        self.addCleanup(shutil.rmtree, self.folder, True)
        original = cache.CACHE_DIR
        cache.CACHE_DIR = self.folder
        self.addCleanup(setattr, cache, "CACHE_DIR", original)
        self.media = os.path.join(self.folder, "素材.mp4")
        with open(self.media, "wb") as fh:
            fh.write(b"\0")
        os.utime(self.media, (1000, 1000))
        # 判定のときに 100〜110 秒を起こした保存分
        words = [Word("今日は", 101.0, 101.5), Word("ゲームを", 101.6, 102.2), Word("やります。", 102.3, 103.0),
                 Word("よろしく", 106.0, 106.6), Word("お願いします。", 106.7, 107.5)]
        transcribe.save_cached(self.media, 100.0, 110.0, Transcript("whisper-large-v3", words, 3.0, "GPU"))
        self.transcribe = transcribe
        self.settings = localedit.build_settings(localedit.flatten_profile(localedit.load_profile("default")))

    def test_find_covering_cuts_out_the_range(self):
        found = self.transcribe.find_covering(self.media, 105.5, 108.0, "whisper-large-v3")
        self.assertEqual([w.text for w in found.words], ["よろしく", "お願いします。"])
        self.assertIsNone(self.transcribe.find_covering(self.media, 95.0, 108.0, "whisper-large-v3"))
        self.assertIsNone(self.transcribe.find_covering(self.media, 101.0, 102.0, "qwen"))

    def test_captions_follow_timeline_positions(self):
        # カット後: 素材 101〜103 をシーケンス 0 秒に、素材 106〜107.5 をシーケンス 2 秒に置いた
        clips = [{"media_path": self.media, "source_in": 101.0, "source_out": 103.0, "seq_start": 0.0},
                 {"media_path": self.media, "source_in": 106.0, "source_out": 107.5, "seq_start": 2.0}]
        with mock.patch.object(self.transcribe, "transcribe_segment") as again:
            srt, count = localedit.captions_for_clips(clips, 30.0, self.settings)
        again.assert_not_called()          # 保存分だけで足りる。起こし直さない
        self.assertEqual(count, 2)
        blocks = srt.strip().split("\n\n")
        self.assertIn("00:00:00,000 --> 00:00:02,000", blocks[0])
        self.assertIn("今日はゲームをやります", blocks[0])      # 行末の句点は消える（default.yaml）
        self.assertIn("00:00:02,000 -->", blocks[1])
        self.assertIn("よろしくお願いします", blocks[1])

    def test_uncached_range_is_transcribed(self):
        clips = [{"media_path": self.media, "source_in": 200.0, "source_out": 203.0, "seq_start": 10.0}]
        made = Transcript("whisper-large-v3", [Word("新しい", 200.5, 201.0)], 1.0, "GPU")
        with mock.patch.object(self.transcribe, "transcribe_segment", return_value=made) as again:
            srt, count = localedit.captions_for_clips(clips, 30.0, self.settings)
        again.assert_called_once()
        self.assertEqual(count, 1)
        self.assertIn("00:00:10,500", srt)

    def test_api_writes_srt_next_to_project(self):
        from fastapi.testclient import TestClient

        from stream_highlight import server

        project = os.path.join(self.folder, "案件.prproj")
        open(project, "w").close()
        open(os.path.join(self.folder, "本編_テロップ.srt"), "w").close()     # 既にある名前は上書きしない
        client = TestClient(server.app)
        with mock.patch.object(deps, "problems", return_value=[]):
            res = client.post("/api/panel/captions", json={
                "clips": [{"media_path": self.media, "source_in": 101.0, "source_out": 103.0,
                           "seq_start": 0.0}],
                "fps": 29.97, "project_path": project, "name": "本編"})
        self.assertEqual(res.status_code, 200, res.text)
        for _ in range(100):
            job = client.get("/api/job/" + res.json()["job_id"]).json()
            if job["status"] in ("done", "error"):
                break
            time.sleep(0.05)
        self.assertEqual(job["status"], "done", job)
        self.assertEqual(os.path.basename(job["result"]["srt_path"]), "本編_テロップ (2).srt")
        with open(job["result"]["srt_path"], encoding="utf-8") as fh:
            self.assertIn("今日は", fh.read())
        res = client.post("/api/panel/captions", json={"clips": [], "fps": 30, "project_path": project})
        self.assertEqual(res.status_code, 400)
        res = client.post("/api/panel/captions", json={
            "clips": [{"media_path": self.media, "source_in": 0, "source_out": 1, "seq_start": 0}],
            "fps": 30, "project_path": ""})
        self.assertEqual(res.status_code, 400)


NODE_CHECK = r"""
const assert = require("assert");
const plan = require(process.argv[2]);
const T = plan.TICKS_PER_SECOND;
const frame = T / 30;                                   // 30fps
// シーケンス 10秒の位置に、素材の 100秒から 8秒ぶん置いたクリップ
const clip = { startTicks: 10 * T, endTicks: 18 * T, inTicks: 100 * T };

const whole = plan.analysisPortion(clip, null);
assert.strictEqual(whole.sourceIn, 100);
assert.strictEqual(whole.sourceOut, 108);

const pieces = plan.keepPieces(clip, whole, [[100.5, 102.0], [103.0, 107.99]], frame);
assert.deepStrictEqual(pieces.map(p => [p.inTicks / T, p.outTicks / T]), [[100.5, 102], [103, 108]]);
assert.ok(!plan.unchanged(clip, pieces));
assert.ok(plan.unchanged(clip, plan.keepPieces(clip, whole, [[100, 108]], frame)));

// イン〜アウト（シーケンスの 12〜15秒）だけを判定したとき、外側はそのまま残る
const part = plan.analysisPortion(clip, { inTicks: 12 * T, outTicks: 15 * T });
assert.strictEqual(part.sourceIn, 102);
assert.strictEqual(part.sourceOut, 105);
const kept = plan.keepPieces(clip, part, [[102, 103], [104, 105]], frame);
assert.deepStrictEqual(kept.map(p => [p.inTicks / T, p.outTicks / T]), [[100, 103], [104, 108]]);
assert.strictEqual(plan.keptTicks(kept) / T, 7);

// 消す所はシーケンスの位置に（素材 103〜104秒 → シーケンス 13〜14秒）
const marks = plan.removedOnTimeline(part, [{ start: 103, end: 104, reasons: ["単語の間"] }], frame);
assert.deepStrictEqual(marks, [{ startTicks: 13 * T, durationTicks: 1 * T, reasons: "単語の間" }]);

// フレームに揃う（1/30秒より細かい値は丸める）
const odd = plan.keepPieces(clip, whole, [[100.51, 101.0]], frame);
assert.strictEqual((odd[0].inTicks - clip.inTicks) % frame, 0);

// クリップと重ならない範囲
assert.strictEqual(plan.analysisPortion(clip, { inTicks: 20 * T, outTicks: 30 * T }), null);

// レーザー方式: 素材 100.5〜102 と 103〜108 を残す → シーケンス 10.5 / 12 / 13 秒で切り、
// 10〜10.5 と 12〜13 を消す。クリップの端（18秒）は切らない
const cuts = plan.timelineCuts(clip, pieces);
assert.deepStrictEqual(cuts.boundaries.map(t => t / T), [10.5, 12, 13]);
assert.deepStrictEqual(cuts.removed.map(r => r.map(t => t / T)), [[10, 10.5], [12, 13]]);
assert.deepStrictEqual(plan.timelineCuts(clip, [{ inTicks: 100 * T, outTicks: 108 * T }]),
                       { boundaries: [], removed: [] });

// 入れ直し方式で失うもの
const clean = {
  speed: 1, reversed: false, disabled: false, isSequence: false, name: "a.mp4", projectName: "a.mp4",
  audioTracks: [0], transitions: 0,
  video: [{ matchName: "AE.ADBE Motion", displayName: "モーション",
            params: [{ name: "位置", keyframes: 0, value: null }, { name: "スケール", keyframes: 0, value: 100 },
                     { name: "幅のスケール", keyframes: 0, value: 100 }, { name: "x", keyframes: 0, value: null },
                     { name: "回転", keyframes: 0, value: 0 }] },
          { matchName: "AE.ADBE Opacity", displayName: "不透明度", params: [{ name: "不透明度", keyframes: 0, value: 100 }] },
          { matchName: "AE.ADBE Time Remapping", displayName: "タイムリマップ", params: [] }],
  audio: [{ matchName: "AE.ADBE Audio Volume", displayName: "ボリューム", params: [{ name: "レベル", keyframes: 0, value: 0 }] },
          { matchName: "AE.ADBE Audio Panner", displayName: "パンナー", params: [] }],
};
assert.deepStrictEqual(plan.cutLosses(clean), { blocks: [], notes: [] });
const dirty = JSON.parse(JSON.stringify(clean));
dirty.speed = 2;
dirty.video.push({ matchName: "AE.ADBE Lumetri", displayName: "Lumetri カラー", params: [] });
dirty.video[0].params[1].value = 120;
dirty.audio[0].params[0].keyframes = 3;
dirty.name = "OP";
const losses = plan.cutLosses(dirty);
assert.deepStrictEqual(losses.blocks, ["速度の変更（200%）", "スケールの変更（120）",
                                       "エフェクト「Lumetri カラー」", "音声のキーフレーム（ボリューム の レベル）"]);
assert.strictEqual(losses.notes.length, 1);
assert.ok(plan.cutLosses({ ...clean, audioTracks: [] }).blocks[0].includes("音声がリンクしていない"));
assert.ok(plan.cutLosses({ ...clean, audioTracks: [0, 2] }).blocks[0].includes("離れたトラック"));
assert.ok(plan.cutLosses({ ...clean, transitions: 1 }).blocks[0].includes("トランジション"));
console.log("ok");
"""


@unittest.skipUnless(shutil.which("node"), "node が無い")
class TestCutPlanJs(unittest.TestCase):
    def test_cutplan(self):
        folder = tempfile.mkdtemp(prefix="sh_cutplan_")
        self.addCleanup(shutil.rmtree, folder, True)
        script = os.path.join(folder, "check.js")
        with open(script, "w", encoding="utf-8") as fh:
            fh.write(NODE_CHECK)
        done = subprocess.run(["node", script, os.path.join(PANEL, "cutplan.js")],
                              capture_output=True, text=True, encoding="utf-8")
        self.assertEqual(done.returncode, 0, done.stderr)
        self.assertEqual(done.stdout.strip(), "ok")


class TestManifest(unittest.TestCase):
    def test_manifest(self):
        with open(os.path.join(PANEL, "manifest.json"), encoding="utf-8") as fh:
            manifest = json.load(fh)
        self.assertEqual(manifest["manifestVersion"], 5)
        self.assertEqual(manifest["host"]["app"], "premierepro")
        # 127.0.0.1 とポートの指定は Premiere の通信許可で弾かれたので、全体を許可している
        self.assertEqual(manifest["requiredPermissions"]["network"]["domains"], "all")
        with open(os.path.join(PANEL, "index.html"), encoding="utf-8") as fh:
            html = fh.read()
        for element in ("previewBtn", "cutBtn", "captionBtn", "clearBtn", "diagnoseBtn", "useRazor",
                        "strategy", "alignBtn", "lineHeight", "keepX", "alignDryRun"):
            self.assertIn('id="%s"' % element, html)
        for name in (manifest["main"], "main.js", "cutplan.js", "captionalign.js"):
            self.assertTrue(os.path.isfile(os.path.join(PANEL, name)), name)


ALIGN_CHECK = r"""
const assert = require("assert");
const a = require(process.argv[2]);
assert.strictEqual(a.lineCount("一行目\r二行目"), 2);
assert.strictEqual(a.lineCount("一行"), 1);
assert.strictEqual(a.lineCount("一\n\n"), 1);
assert.strictEqual(a.textOf("\u0000{\"mTextParam\":{\"mText\":\"あ\\rい\",\"mFontSize\":64}}"), "あ\rい");
assert.strictEqual(a.textOf("ただの文字"), "ただの文字");
assert.strictEqual(a.fontSizeOf("{\"x\":{\"mFontSize\":64}}"), 64);
// 割合: 2行は 1行分（80/1080）上、1行はそのまま、3行は2行分上
let r = a.planAlign([{lines: 1, x: 0.5, y: 0.85}, {lines: 2, x: 0.5, y: 0.85}, {lines: 3, x: 0.4, y: 0.1}], 80, 1080);
assert.strictEqual(r.unit, "割合");
assert.ok(!r.moves[0].changed);
assert.ok(Math.abs(r.moves[1].to.y - (0.85 - 80 / 1080)) < 1e-9);
assert.ok(Math.abs(r.moves[2].to.y - (0.85 - 160 / 1080)) < 1e-9);
assert.strictEqual(r.moves[2].to.x, 0.5);
// ピクセル・横はそのまま
r = a.planAlign([{lines: 1, x: 960, y: 950}, {lines: 2, x: 900, y: 950}], 80, 1080, {keepX: true});
assert.strictEqual(r.unit, "px");
assert.strictEqual(r.moves[1].to.y, 870);
assert.strictEqual(r.moves[1].to.x, 900);
// 1行のテロップが無ければ止める
assert.throws(() => a.planAlign([{lines: 2, x: 0.5, y: 0.8}], 80, 1080));
// クリップのモーションで 0.06 下げているものが多数。モーションが既定（中央）のものは、その分テキストを下げる
r = a.planAlign([
  {lines: 1, x: 0.125, y: 0.65, dx: 0, dy: 0.06},
  {lines: 1, x: 0.125, y: 0.65, dx: 0, dy: 0.06},
  {lines: 1, x: 0.125, y: 0.65, dx: 0, dy: 0},
  {lines: 2, x: 0.125, y: 0.65, dx: 0, dy: 0},
], 108, 1080);
assert.ok(Math.abs(r.base.y - 0.71) < 1e-9);
assert.ok(!r.moves[0].changed);
assert.ok(Math.abs(r.moves[2].to.y - 0.71) < 1e-9);
assert.ok(Math.abs(r.moves[3].to.y - 0.61) < 1e-9);
// 下段の位置をピクセルで決める（1行のテロップが無くてもよい）。モーションのずれは引く
r = a.planAlign([
  {lines: 2, x: 0.125, y: 0.52, dx: 0, dy: 0.06},
  {lines: 1, x: 0.125, y: 0.65, dx: 0, dy: 0},
], 108, 1080, {baseYPx: 918});
assert.ok(Math.abs(r.base.y - 0.85) < 1e-9);
assert.ok(Math.abs(r.moves[0].to.y - (0.85 - 0.1 - 0.06)) < 1e-9);
assert.ok(Math.abs(r.moves[1].to.y - 0.85) < 1e-9);
console.log("ok");
"""


@unittest.skipUnless(shutil.which("node"), "node が無い")
class TestCaptionAlignJs(unittest.TestCase):
    def test_captionalign(self):
        folder = tempfile.mkdtemp(prefix="sh_align_")
        self.addCleanup(shutil.rmtree, folder, True)
        script = os.path.join(folder, "check.js")
        with open(script, "w", encoding="utf-8") as fh:
            fh.write(ALIGN_CHECK)
        done = subprocess.run(["node", script, os.path.join(PANEL, "captionalign.js")],
                              capture_output=True, text=True, encoding="utf-8")
        self.assertEqual(done.returncode, 0, done.stderr)
        self.assertEqual(done.stdout.strip(), "ok")


if __name__ == "__main__":
    unittest.main()
