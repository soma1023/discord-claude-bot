# -*- coding: utf-8 -*-
"""手元の動画の自動編集（localedit.py）に要るものが揃っているかを確かめる。

足りないときは、何を・どう入れればよいかを日本語で返す。
重いライブラリ（torch・faster-whisper）は読み込まず、入っているかだけを見る。

  python -m stream_highlight.autoedit --check
"""

import shutil
import subprocess
import sys

from . import transcribe

INSTALL_HINTS = {
    "ffmpeg": "winget install Gyan.FFmpeg  （または pip install imageio-ffmpeg）",
    "faster-whisper": "pip install faster-whisper  （GPU を使うなら先に CUDA 版の torch を入れる。"
                      "stream_highlight/requirements-transcribe.txt を参照）",
    "qwen-asr": "stream_highlight/requirements-transcribe.txt の手順で入れる",
    "pyyaml": "pip install pyyaml",
    "auto-editor": "pip install auto-editor  （初回の実行時に本体を自動でダウンロードする）",
}


class MissingDependency(RuntimeError):
    """足りない依存がある。メッセージはそのまま画面に出す。"""


def auto_editor_command():
    """auto-editor の起動コマンド。見つからなければ空のリスト。"""
    exe = shutil.which("auto-editor")
    if exe:
        return [exe]
    if transcribe._has("auto_editor"):
        return [sys.executable, "-m", "auto_editor"]
    return []


def auto_editor_version():
    command = auto_editor_command()
    if not command:
        return ""
    try:
        done = subprocess.run(command + ["--version"], capture_output=True, text=True,
                              timeout=120, creationflags=transcribe._no_window())
    except (OSError, subprocess.SubprocessError):
        return ""
    return (done.stdout or "").strip().splitlines()[-1] if done.returncode == 0 and done.stdout else ""


def problems(engine, silence_method):
    """足りないものの一覧 [(名前, 説明)]。空なら始めてよい。"""
    found = []
    if not transcribe.find_ffmpeg():
        found.append(("ffmpeg", "ffmpeg が見つかりません。"))
    if not transcribe._has("yaml"):
        found.append(("pyyaml", "プロファイル（yaml）を読む PyYAML が入っていません。"))
    lacking = [m for m in transcribe.ENGINE_MODULES.get(engine, []) if not transcribe._has(m)]
    if engine not in transcribe.ENGINE_MODULES:
        found.append((engine, "文字起こしエンジン「%s」はありません（whisper-large-v3 / whisper / qwen）。"
                      % engine))
    elif lacking:
        name = "qwen-asr" if engine == "qwen" else "faster-whisper"
        found.append((name, "文字起こしに使う %s が入っていません（%s）。" % (name, "、".join(lacking))))
    if silence_method == "audio" and not auto_editor_command():
        found.append(("auto-editor", "silence_method: audio には auto-editor が必要です。"))
    return found


def require(engine, silence_method):
    """足りないものがあれば、入れ方を添えて MissingDependency を投げる。"""
    found = problems(engine, silence_method)
    if not found:
        return
    lines = ["必要なものが足りません:"]
    for name, message in found:
        lines.append("  - %s" % message)
        if name in INSTALL_HINTS:
            lines.append("      入れ方: %s" % INSTALL_HINTS[name])
    raise MissingDependency("\n".join(lines))


def report():
    """--check で出す一覧（日本語の行）。"""
    rows = [("ffmpeg", transcribe.find_ffmpeg() or "")]
    rows.append(("PyYAML", "入っています" if transcribe._has("yaml") else ""))
    rows.append(("faster-whisper", "入っています" if transcribe._has("faster_whisper") else ""))
    rows.append(("qwen-asr（任意）", "入っています" if transcribe._has("qwen_asr") else ""))
    rows.append(("auto-editor（silence_method: audio のときだけ）", auto_editor_version()))
    try:
        import importlib.util
        cuda = "torch が無いので不明"
        if importlib.util.find_spec("torch") is not None:
            import torch
            cuda = "使えます（%s）" % torch.cuda.get_device_name(0) if torch.cuda.is_available() \
                else "使えません（CPU で動きます。遅くなります）"
    except Exception as exc:          # noqa: BLE001
        cuda = "確認できませんでした: %s" % exc
    lines = []
    for name, value in rows:
        lines.append("%s: %s" % (name, value or "見つかりません"))
    lines.append("GPU（CUDA）: %s" % cuda)
    return lines
