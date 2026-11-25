# -*- coding: utf-8 -*-
"""Premiere のオートカットパネル（premiere_panel/）を .ccx にまとめて、Premiere に入れる。

  python -m stream_highlight.install_panel

UXP Developer Tools も開発者モードも要らない。Creative Cloud に付いている
Adobe のプラグインインストーラー（UnifiedPluginInstallerAgent）で入れる。
パネルを直したら、もう一度これを実行して Premiere を再起動すれば入れ替わる。

サーバー（run_app.py）を起動するショートカットも dist/ に作り、その場所をパネルに書き込む。
パネルは開いたときにサーバーが止まっていれば、このショートカットを開いて起動する。
"""

import json
import os
import subprocess
import sys
import zipfile

PANEL_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "premiere_panel")
FILES = ("manifest.json", "index.html", "main.js", "cutplan.js", "captionalign.js")
PANEL_NAME = "オートカット"       # manifest.json の name。インストーラーはこの名前で外す
INSTALLER = os.path.join(
    os.environ.get("CommonProgramFiles", r"C:\Program Files\Common Files"),
    "Adobe", "Adobe Desktop Common", "RemoteComponents", "UPI",
    "UnifiedPluginInstallerAgent", "UnifiedPluginInstallerAgent.exe")


REPO_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LAUNCH_MODULE = "serverlaunch.js"


def make_launcher(out_dir):
    """サーバーを黒い画面なしで起動するショートカット（.lnk）を作り、その場所を返す。作れなければ None。"""
    pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    if not os.path.isfile(pythonw):
        return None
    link = os.path.join(out_dir, "サーバー起動.lnk")
    script = ("$s = (New-Object -ComObject WScript.Shell).CreateShortcut($env:LNK); "
              "$s.TargetPath = $env:TARGET; $s.Arguments = $env:ARGS; "
              "$s.WorkingDirectory = $env:WORKDIR; $s.WindowStyle = 7; $s.Save()")
    env = dict(os.environ, LNK=link, TARGET=pythonw, WORKDIR=REPO_DIR,
               ARGS='"%s" --no-browser' % os.path.join(REPO_DIR, "run_app.py"))
    done = subprocess.run(["powershell", "-NoProfile", "-Command", script], env=env,
                          capture_output=True, text=True, encoding="utf-8", errors="replace")
    return link if done.returncode == 0 and os.path.isfile(link) else None


def build(out_path, launcher=None):
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as package:
        for name in FILES:
            package.write(os.path.join(PANEL_DIR, name), name)
        package.writestr(LAUNCH_MODULE, "module.exports = %s;\n" % json.dumps(
            {"launcher": launcher}, ensure_ascii=False))
    return out_path


def main():
    out_dir = os.path.join(PANEL_DIR, "dist")
    os.makedirs(out_dir, exist_ok=True)
    launcher = make_launcher(out_dir)
    if launcher:
        print("サーバーを起動するショートカットを作りました: %s" % launcher)
    else:
        print("サーバーを起動するショートカットを作れませんでした。パネルからの自動起動は使えません。",
              file=sys.stderr)
    package = build(os.path.join(out_dir, "autocut.ccx"), launcher)
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
