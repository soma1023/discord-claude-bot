# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller の設定。`build_exe.bat` から使う。

yt-dlp は配信サイトごとの抽出器を実行時に読み込むため、
まとめて同梱しないと「対応していないURL」と言われてしまう。
"""

import os

from PyInstaller.utils.hooks import collect_submodules

hiddenimports = (
    collect_submodules("yt_dlp")
    + collect_submodules("uvicorn")
    + ["fastapi", "pydantic"]
)

datas = [("stream_highlight/static", "stream_highlight/static")]
if os.path.exists("stream_highlight/_build_id.txt"):
    datas.append(("stream_highlight/_build_id.txt", "stream_highlight"))

a = Analysis(
    ["run_app.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    # 使わない重いものは入れない
    excludes=["tkinter", "matplotlib", "numpy", "PIL", "test", "unittest"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    exclude_binaries=True,
    name="StreamHighlight",
    debug=False,
    strip=False,
    upx=False,
    console=True,          # URLとエラーが見えるように、コンソールは出す
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="StreamHighlight",
)
