# -*- coding: utf-8 -*-
"""チャット取得をバックグラウンドで走らせ、進捗をUIに返す。"""

import threading
import time
import traceback
import uuid
from collections import OrderedDict

from . import cache
from .analyze import analyze
from .sources import FetchError, fetch_chat, fetch_info, parse_url

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
        if loaded:
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

    def _run(self, job):
        job.status = "running"
        try:
            platform, video_id = parse_url(job.url)
            key = "%s:%s" % (platform, video_id)
            job.video_key = key

            loaded = None if job.refresh else self.get_chat(key)
            if loaded:
                info, messages = loaded
                job.message = "保存済みのチャットを使用中…"
                job.progress = 0.9
            else:
                def progress(message, frac):
                    job.message = message
                    if frac is not None:
                        job.progress = max(job.progress, frac * 0.9)

                job.message = "配信情報を取得中…"
                info = fetch_info(job.url)
                job.progress = 0.05
                messages = fetch_chat(info, progress=progress)
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
            job.error = "解析中に予期しないエラーが発生しました: %s" % exc
            job.message = "失敗しました"


manager = JobManager()
