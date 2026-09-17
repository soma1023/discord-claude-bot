# -*- coding: utf-8 -*-
"""配信音声のラウドネス（音の大きさ）を時系列で取り出す。

単純な波形の振幅ではなく、ffmpeg の ebur128 フィルタが出す
モーメンタリラウドネス（LUFS / 400ms窓・100ms刻み）を使う。
人間の聞こえ方に合わせた重み付けがされているので、
「配信者の声が大きくなった」を素直に拾いやすい。

ただしアーカイブ音声は配信者のマイクとゲーム音が混ざった1本のトラックなので、
これ単体では爆発音などのSEと声の区別がつかない。
チャットとの照合（analyze.py 側）で誤検知を落とす前提の部品。
"""

import glob
import os
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass

from .sources import FetchError, _run_ytdlp, _ytdlp_error

STEP_SEC = 0.1          # ebur128 が値を出す間隔
SILENCE_FLOOR = -70.0   # 無音は -120 付近まで落ちるので、統計が壊れない値で止める

_META_LINE = re.compile(r"^lavfi\.r128\.M=(-?[\d.]+|-?inf)$")
_TIME_LINE = re.compile(r"^frame:\d+\s+pts:\S+\s+pts_time:([\d.]+)")
_YTDLP_PROGRESS = re.compile(r"\[download\]\s+([\d.]+)%")


@dataclass
class Loudness:
    """100ms刻みのラウドネス列（単位: LUFS）。"""

    step: float
    values: list

    @property
    def duration(self):
        return len(self.values) * self.step

    def bin_max(self, bin_sec, n_bins):
        """チャット側と同じ時間刻みに合わせ、各区間の最大値を取る。

        平均ではなく最大なのは、一瞬の叫びを均して消したくないため。
        """
        per_bin = max(1, int(round(bin_sec / self.step)))
        out = []
        for b in range(n_bins):
            lo = b * per_bin
            hi = min(len(self.values), lo + per_bin)
            out.append(max(self.values[lo:hi]) if lo < hi else SILENCE_FLOOR)
        return out

    def as_dict(self):
        return {"step": self.step, "values": [round(v, 2) for v in self.values]}

    @classmethod
    def from_dict(cls, data):
        return cls(step=float(data["step"]), values=[float(v) for v in data["values"]])


# ---------------------------------------------------------------- ffmpeg

def find_ffmpeg():
    """ffmpeg の実行ファイルを探す。PATH → pip同梱版 の順。"""
    exe = shutil.which("ffmpeg") or shutil.which("ffmpeg.exe")
    if exe:
        return exe
    try:
        import imageio_ffmpeg
        return imageio_ffmpeg.get_ffmpeg_exe()
    except Exception:       # noqa: BLE001  （未導入・DLエラーなど理由を問わず次の案内に倒す）
        pass
    raise FetchError(
        "音声解析には ffmpeg が必要です。\n"
        "`pip install imageio-ffmpeg` を実行すると同梱版が使えるようになります。\n"
        "（チャットだけの解析は ffmpeg なしで動きます）"
    )


def ffmpeg_available():
    try:
        find_ffmpeg()
        return True
    except FetchError:
        return False


# ---------------------------------------------------------------- 取得と解析

def download_audio(url, tmpdir, progress=None):
    """音声トラックだけを落とす（映像はダウンロードしない）。"""
    outtmpl = os.path.join(tmpdir, "audio.%(ext)s")

    def on_line(line):
        m = _YTDLP_PROGRESS.search(line)
        if m and progress:
            pct = float(m.group(1))
            progress("音声をダウンロード中… %.0f%%" % pct, pct / 100.0)

    # --extract-audio は yt-dlp 側に ffmpeg を要求する（PATHに無いと失敗する）。
    # bestaudio はもともと音声だけのトラックなので、変換せずそのまま解析する。
    # 音声だけの形式が無い配信では、数GBの映像を黙って落とさずエラーにする。
    code, lines = _run_ytdlp(
        ["-f", "bestaudio", "-o", outtmpl, url],
        on_line=on_line,
    )
    files = [p for p in glob.glob(os.path.join(tmpdir, "audio.*"))
             if not p.endswith(".part")]
    if not files:
        if any("Requested format is not available" in l for l in lines):
            raise FetchError("この配信には音声だけのトラックがないため、音声解析に対応できません。")
        raise FetchError(_ytdlp_error(lines) if code != 0
                         else "音声トラックを取得できませんでした。")
    return max(files, key=os.path.getsize)


def extract_loudness(path, duration=0, progress=None):
    """ffmpeg にラウドネスを計算させ、100ms刻みの列にする。"""
    exe = find_ffmpeg()
    errlog = tempfile.TemporaryFile()
    proc = subprocess.Popen(
        [exe, "-nostdin", "-hide_banner", "-i", path, "-vn", "-ac", "1",
         "-af", "ebur128=metadata=1,ametadata=print:key=lavfi.r128.M:file=-",
         "-f", "null", "-"],
        stdout=subprocess.PIPE, stderr=errlog,
        text=True, encoding="utf-8", errors="replace",
    )

    values = []
    last_report = 0.0
    try:
        for line in proc.stdout:
            line = line.strip()
            m = _META_LINE.match(line)
            if m:
                raw = m.group(1)
                value = SILENCE_FLOOR if raw.endswith("inf") else float(raw)
                values.append(max(SILENCE_FLOOR, value))
                continue
            t = _TIME_LINE.match(line)
            if t and progress and duration:
                seconds = float(t.group(1))
                if seconds - last_report >= 30:
                    last_report = seconds
                    progress("音量を解析中… %d%%" % min(99, int(seconds / duration * 100)),
                             min(0.99, seconds / duration))
    finally:
        proc.stdout.close()
        proc.wait()

    if proc.returncode != 0 and not values:
        errlog.seek(0)
        detail = errlog.read().decode("utf-8", "replace").strip().split("\n")[-1:]
        errlog.close()
        raise FetchError("音声の解析に失敗しました: %s" % (detail[0] if detail else "原因不明"))
    errlog.close()
    if not values:
        raise FetchError("音声から音量を取得できませんでした。")
    return Loudness(step=STEP_SEC, values=values)


def fetch_loudness(info, progress=None):
    """音声をダウンロードしてラウドネス列を返す。一時ファイルは必ず消す。"""
    find_ffmpeg()      # 落としてから足りないと分かるのは無駄なので先に確認する
    tmpdir = tempfile.mkdtemp(prefix="stream_highlight_audio_")
    try:
        def dl_progress(message, frac):
            if progress:
                progress(message, None if frac is None else frac * 0.7)

        path = download_audio(info.url, tmpdir, progress=dl_progress)

        def an_progress(message, frac):
            if progress:
                progress(message, None if frac is None else 0.7 + frac * 0.3)

        return extract_loudness(path, duration=info.duration, progress=an_progress)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
