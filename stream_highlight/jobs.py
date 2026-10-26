# -*- coding: utf-8 -*-
"""チャット取得をバックグラウンドで走らせ、進捗をUIに返す。"""

import os
import threading
import time
import traceback
import uuid
from collections import OrderedDict

from . import cache
from .analyze import analyze
from .sources import FetchError, fetch_chat, fetch_info, parse_url
from .transcribe import TranscribeError

MAX_JOBS = 12          # 履歴として保持するジョブ数
MAX_LOADED = 4         # メモリに載せておく配信数（1件で数万〜十数万コメント）


class Job:
    def __init__(self, job_id, url, params, refresh):
        self.id = job_id
        self.url = url
        self.params = params
        self.refresh = refresh
        self.status = "pending"      # pending / running / done / error
        self.message = "待機中…"
        self.progress = 0.0
        self.result = None
        self.error = ""
        self.video_key = ""
        self.created_at = time.time()

    def as_dict(self):
        return {
            "job_id": self.id,
            "status": self.status,
            "message": self.message,
            "progress": round(self.progress, 3),
            "error": self.error,
            "video_key": self.video_key,
            "result": self.result,
        }


class JobManager:
    """ジョブとチャット本体を抱えて、再解析を即座に返せるようにする。"""

    def __init__(self):
        self._jobs = OrderedDict()
        self._loaded = OrderedDict()     # video_key -> (StreamInfo, messages)
        self._lock = threading.Lock()

    # -- チャット本体の保持 ------------------------------------------

    def _remember(self, key, info, messages):
        with self._lock:
            self._loaded[key] = (info, messages)
            self._loaded.move_to_end(key)
            while len(self._loaded) > MAX_LOADED:
                self._loaded.popitem(last=False)

    def get_chat(self, video_key):
        """メモリ→ディスクキャッシュの順に探す。見つからなければ None。"""
        with self._lock:
            found = self._loaded.get(video_key)
            if found:
                self._loaded.move_to_end(video_key)
                return found
        if ":" not in video_key:
            return None
        platform, video_id = video_key.split(":", 1)
        loaded = cache.load(platform, video_id)
        if not loaded:
            return None
        self._remember(video_key, loaded[0], loaded[1])
        return loaded

    # -- ジョブ ------------------------------------------------------

    def submit(self, url, params, refresh=False):
        parse_url(url)   # 先にURLを検証して、おかしければここで弾く
        job = Job(uuid.uuid4().hex[:12], url, params, refresh)
        with self._lock:
            self._jobs[job.id] = job
            while len(self._jobs) > MAX_JOBS:
                self._jobs.popitem(last=False)
        threading.Thread(target=self._run, args=(job,), daemon=True).start()
        return job

    def get(self, job_id):
        with self._lock:
            return self._jobs.get(job_id)

    def submit_task(self, work, message="準備中…"):
        """解析以外の時間のかかる処理（文字起こし付きの書き出しなど）を走らせる。

        work(report) を別スレッドで呼び、戻り値をそのまま結果にする。
        report(message, frac) で進み具合を画面に伝えられる。
        進捗は解析と同じ /api/job/{id} で問い合わせる。
        """
        job = Job(uuid.uuid4().hex[:12], "", None, False)
        job.message = message
        with self._lock:
            self._jobs[job.id] = job
            while len(self._jobs) > MAX_JOBS:
                self._jobs.popitem(last=False)
        threading.Thread(target=self._run_task, args=(job, work), daemon=True).start()
        return job

    def _run_task(self, job, work):
        job.status = "running"

        def report(message, frac=None):
            if message:
                job.message = message
            if frac is not None:
                job.progress = max(job.progress, min(1.0, float(frac)))
        try:
            job.result = work(report)
            job.progress = 1.0
            job.message = "完了"
            job.status = "done"
        except (ValueError, TranscribeError) as exc:
            job.status = "error"
            job.error = str(exc)
            job.message = "失敗しました"
        except Exception as exc:                      # noqa: BLE001
            traceback.print_exc()
            job.status = "error"
            job.error = "書き出し中に予期しないエラーが発生しました: %s%s" % (exc, _where(exc))
            job.message = "失敗しました"

    def _run(self, job):
        job.status = "running"
        try:
            platform, video_id = parse_url(job.url)
            key = "%s:%s" % (platform, video_id)
            job.video_key = key

            chat_share = 0.9

            def stage_progress(low, high):
                def report(message, frac):
                    job.message = message
                    if frac is not None:
                        job.progress = max(job.progress, low + frac * (high - low))
                return report

            loaded = None if job.refresh else self.get_chat(key)
            if loaded:
                info, messages = loaded
                job.message = "保存済みのチャットを使用中…"
                job.progress = chat_share
            else:
                job.message = "配信情報を取得中…"
                info = fetch_info(job.url)
                job.progress = 0.05
                messages = fetch_chat(info, progress=stage_progress(0.05, chat_share))
                if not messages:
                    raise FetchError("チャットが1件も取得できませんでした。")
                cache.save(info, messages)
                self._remember(key, info, messages)

            job.message = "盛り上がりを解析中…"
            job.progress = 0.95
            job.result = analyze(messages, info, job.params)
            job.result["video_key"] = key
            job.progress = 1.0
            job.message = "完了"
            job.status = "done"
        except FetchError as exc:
            job.status = "error"
            job.error = str(exc)
            job.message = "失敗しました"
        except Exception as exc:                      # noqa: BLE001
            traceback.print_exc()
            job.status = "error"
            job.error = "解析中に予期しないエラーが発生しました: %s%s" % (exc, _where(exc))
            job.message = "失敗しました"


def _where(exc):
    """発生箇所が分からないと報告を受けても直せないので、画面にも出す。"""
    frames = traceback.extract_tb(exc.__traceback__)
    if not frames:
        return ""
    last = frames[-1]
    return " [%s:%d %s]" % (os.path.basename(last.filename), last.lineno, last.name)


manager = JobManager()
