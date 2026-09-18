# -*- coding: utf-8 -*-
"""API全体の流れを、ネットワークを使わずに確かめる。

キャッシュに合成チャットを仕込むと取得処理が丸ごと省かれるので、
そのままサーバの動作確認に使える。
"""

import os
import sys
import tempfile
import time
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from fastapi.testclient import TestClient

from stream_highlight import cache
from stream_highlight.jobs import manager
from stream_highlight.server import app
from stream_highlight.tests.test_analyze import DURATION, SPIKES, build_chat

DEMO_ID = "demo1234567"
DEMO_URL = "https://www.youtube.com/watch?v=%s" % DEMO_ID


class TestServer(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmpdir = tempfile.mkdtemp(prefix="sh_cache_")
        cls.original_cache_dir = cache.CACHE_DIR
        cache.CACHE_DIR = cls.tmpdir
        from stream_highlight.sources import StreamInfo
        info = StreamInfo("youtube", DEMO_ID, DEMO_URL, title="テスト配信",
                          channel="テスト配信者", duration=DURATION)
        cls.info = info
        cache.save(info, build_chat())
        cls.client = TestClient(app)

    @classmethod
    def tearDownClass(cls):
        import shutil
        cache.CACHE_DIR = cls.original_cache_dir
        shutil.rmtree(cls.tmpdir, ignore_errors=True)

    def setUp(self):
        manager._loaded.clear()

    def run_job(self, url=DEMO_URL, refresh=False, params=None):
        res = self.client.post("/api/analyze",
                               json={"url": url, "refresh": refresh, "params": params})
        self.assertEqual(res.status_code, 200, res.text)
        job_id = res.json()["job_id"]
        for _ in range(100):
            status = self.client.get("/api/job/%s" % job_id).json()
            if status["status"] in ("done", "error"):
                return status
            time.sleep(0.05)
        self.fail("ジョブが終わらない")

    def test_full_flow(self):
        status = self.run_job()
        self.assertEqual(status["status"], "done", status.get("error"))
        result = status["result"]
        self.assertEqual(result["video_key"], "youtube:%s" % DEMO_ID)
        self.assertEqual(result["video"]["title"], "テスト配信")
        self.assertTrue(result["moments"])
        self.assertTrue(any(
            abs(result["moments"][0]["peak_sec"] - s) <= 20 for s in SPIKES))
        self.assertEqual(len(result["series"]["rate"]), len(result["series"]["baseline"]))
        self.assertTrue(result["categories"])

    def test_reanalyze_without_refetch(self):
        self.run_job()
        strict = self.client.post("/api/reanalyze", json={
            "video_key": "youtube:%s" % DEMO_ID, "params": {"min_z": 8.0, "top_n": 5},
        })
        self.assertEqual(strict.status_code, 200, strict.text)
        loose = self.client.post("/api/reanalyze", json={
            "video_key": "youtube:%s" % DEMO_ID, "params": {"min_z": 1.5, "top_n": 50},
        })
        self.assertGreaterEqual(len(loose.json()["moments"]), len(strict.json()["moments"]))
        self.assertEqual(loose.json()["params"]["min_z"], 1.5)

    def test_keyword_endpoint(self):
        self.run_job()
        res = self.client.post("/api/keyword", json={
            "video_key": "youtube:%s" % DEMO_ID, "keyword": "www"})
        self.assertEqual(res.status_code, 200, res.text)
        body = res.json()
        self.assertGreater(body["hits"], 100)
        self.assertTrue(body["moments"])

        empty = self.client.post("/api/keyword", json={
            "video_key": "youtube:%s" % DEMO_ID, "keyword": "  "})
        self.assertEqual(empty.status_code, 400)

    def test_unknown_video_key(self):
        res = self.client.post("/api/reanalyze", json={"video_key": "youtube:nope"})
        self.assertEqual(res.status_code, 404)

    def test_bad_url_rejected_immediately(self):
        res = self.client.post("/api/analyze", json={"url": "https://example.com/watch"})
        self.assertEqual(res.status_code, 400)
        self.assertIn("YouTube", res.json()["detail"])

    def test_history(self):
        entries = self.client.get("/api/history").json()["entries"]
        self.assertTrue(any(e["video_id"] == DEMO_ID for e in entries))
        self.assertEqual(entries[0]["title"], "テスト配信")

    def test_capabilities(self):
        body = self.client.get("/api/capabilities").json()
        self.assertTrue(body["version"])



    def test_index_and_assets(self):
        for path in ("/", "/static/app.js", "/static/style.css"):
            res = self.client.get(path)
            self.assertEqual(res.status_code, 200, path)

    def test_missing_job(self):
        self.assertEqual(self.client.get("/api/job/deadbeef").status_code, 404)


if __name__ == "__main__":
    unittest.main(verbosity=2)
