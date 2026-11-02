# -*- coding: utf-8 -*-
"""Premiere のオートカットパネル（premiere_panel/）を .ccx にまとめて、Premiere に入れる。

  python -m stream_highlight.install_panel

UXP Developer Tools も開発者モードも要らない。Creative Cloud に付いている
Adobe のプラグインインストーラー（UnifiedPluginInstallerAgent）で入れる。
パネルを直したら、もう一度これを実行して Premiere を再起動すれば入れ替わる。
"""

import os
import subprocess
import sys
import zipfile

PANEL_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "premiere_panel")
FILES = ("manifest.json", "index.html", "main.js", "cutplan.js")
PANEL_NAME = "オートカット"       # manifest.json の name。インストーラーはこの名前で外す
INSTALLER = os.path.join(
    os.environ.get("CommonProgramFiles", r"C:\Program Files\Common Files"),
    "Adobe", "Adobe Desktop Common", "RemoteComponents", "UPI",
    "UnifiedPluginInstallerAgent", "UnifiedPluginInstallerAgent.exe")


def build(out_path):
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as package:
        for name in FILES:
            package.write(os.path.join(PANEL_DIR, name), name)
    return out_path


def main():
    out_dir = os.path.join(PANEL_DIR, "dist")
    os.makedirs(out_dir, exist_ok=True)
    package = build(os.path.join(out_dir, "autocut.ccx"))
    print("パネルをまとめました: %s" % package)
    if not os.path.isfile(INSTALLER):
        print("Adobe のプラグインインストーラーが見つかりません（Creative Cloud デスクトップが必要です）。\n"
              "できた .ccx をダブルクリックして入れてください。", file=sys.stderr)
        return 1
    # 古い版が残っていると、Premiere は古いほうを読み込むことがあった（実際に起きた）。
    # 入れる前に、入っている版を全て外す（/remove は1回に1つしか外さない）
    for _ in range(20):
        listed = subprocess.run([INSTALLER, "/list", "all"], capture_output=True, text=True,
                                encoding="utf-8", errors="replace").stdout or ""
        if PANEL_NAME not in listed:
            break
        subprocess.run([INSTALLER, "/remove", PANEL_NAME], capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    done = subprocess.run([INSTALLER, "/install", package], capture_output=True, text=True,
                          encoding="utf-8", errors="replace")
    print((done.stdout or "").strip())
    if done.returncode != 0 or "Successful" not in (done.stdout or ""):
        print("入れられませんでした: %s" % (done.stderr or done.stdout or "").strip(), file=sys.stderr)
        return 1
    print("入れました。Premiere を再起動し、「ウィンドウ」→「UXP プラグイン」→「オートカット」で開いてください。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
