# -*- coding: utf-8 -*-
"""チャット取得をバックグラウンドで走らせ、進捗をUIに返す。"""

import os
import threading
import time
import traceback
import uuid
from collections import OrderedDict

from . import audio as audio_mod
from . import cache
from .analyze import analyze
from .sources import FetchError, fetch_chat, fetch_info, parse_url

MAX_JOBS = 12          # 履歴として保持するジョブ数
MAX_LOADED = 4         # メモリに載せておく配信数（1件で数万〜十数万コメント）


class Job:
    def __init__(self, job_id, url, params, refresh, with_audio=False):
        self.id = job_id
        self.url = url
        self.params = params
        self.refresh = refresh
        self.with_audio = with_audio
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

    def _remember(self, key, info, messages, loudness=None):
        with self._lock:
            previous = self._loaded.get(key)
            if loudness is None and previous:
                loudness = previous[2]        # 既に取ってある音量列は捨てない
            self._loaded[key] = (info, messages, loudness)
            self._loaded.move_to_end(key)
            while len(self._loaded) > MAX_LOADED:
                self._loaded.popitem(last=False)

    def get_chat(self, video_key):
        """メモリ→ディスクキャッシュの順に探す。

        戻り値は (StreamInfo, messages, loudness or None)。見つからなければ None。
        """
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
        info, messages = loaded
        loudness = cache.load_audio(platform, video_id)
        self._remember(video_key, info, messages, loudness)
        return info, messages, loudness

    # -- ジョブ ------------------------------------------------------

    def submit(self, url, params, refresh=False, with_audio=False):
        parse_url(url)   # 先にURLを検証して、おかしければここで弾く
        if with_audio:
            audio_mod.find_ffmpeg()   # 落としてから足りないと分かるのを避ける
        job = Job(uuid.uuid4().hex[:12], url, params, refresh, with_audio)
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

            # 音声まで取るときは、チャットは全体の前半分の進捗として扱う
            chat_share = 0.45 if job.with_audio else 0.9

            def stage_progress(low, high):
                def report(message, frac):
                    job.message = message
                    if frac is not None:
                        job.progress = max(job.progress, low + frac * (high - low))
                return report

            loaded = None if job.refresh else self.get_chat(key)
            if loaded:
                info, messages, loudness = loaded
                job.message = "保存済みのチャットを使用中…"
                job.progress = chat_share
            else:
                loudness = None
                job.message = "配信情報を取得中…"
                info = fetch_info(job.url)
                job.progress = 0.05
                messages = fetch_chat(info, progress=stage_progress(0.05, chat_share))
                if not messages:
                    raise FetchError("チャットが1件も取得できませんでした。")
                cache.save(info, messages)
                self._remember(key, info, messages)

            if job.with_audio and loudness is None:
                job.message = "音声を取得中…"
                loudness = audio_mod.fetch_loudness(
                    info, progress=stage_progress(chat_share, 0.92))
                cache.save_audio(info, loudness)
                self._remember(key, info, messages, loudness)

            job.message = "盛り上がりを解析中…"
            job.progress = 0.95
            job.result = analyze(messages, info, job.params,
                                 loudness=loudness if job.with_audio else None)
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
            # 発生箇所が分からないと報告を受けても直せないので、画面にも出す
            frames = traceback.extract_tb(exc.__traceback__)
            where = ""
            if frames:
                last = frames[-1]
                where = " [%s:%d %s]" % (os.path.basename(last.filename),
                                         last.lineno, last.name)
            job.status = "error"
            job.error = "解析中に予期しないエラーが発生しました: %s%s" % (exc, where)
            job.message = "失敗しました"


manager = JobManager()
