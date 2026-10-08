# -*- coding: utf-8 -*-
"""配信ハイライト抽出ツールの入口。

exe化したときは、ここが2つの役割を兼ねる。

  StreamHighlight.exe            → アプリとして起動する
  StreamHighlight.exe --ytdlp …  → yt-dlp として振る舞う

exe化すると `sys.executable` は自分自身を指すため、`python -m yt_dlp` の
形で外部プロセスを呼べなくなる。そこで自分自身を yt-dlp として
呼び出せるようにしている。

ソースから使うときは、ここが起動の全てを受け持つ。

  pythonw.exe run_app.py --update   → 更新してから、黒い画面を出さずに起動

pythonw.exe はそもそもコンソールを持たないため、黒い画面が一瞬も出ない。
VBScript や PowerShell を経由していた頃は、それらが無効化されたり
引用符が壊れたりすると「何も起きない」で終わっていた。層を無くした。
"""

import multiprocessing
import sys
from datetime import datetime


def _redirect_output():
    """コンソールを出さない設定では、標準出力の行き先が無くなる。

    そのまま print すると失敗するうえ、エラーの内容も残らないので、
    保存先フォルダのログファイルに向ける。
    """
    if sys.stdout is not None and sys.stderr is not None:
        return
    import os

    from stream_highlight import paths

    path = os.path.join(paths.data_dir(), "app.log")
    try:
        stream = open(path, "a", encoding="utf-8", errors="replace")
    except OSError:
        stream = open(os.devnull, "w", encoding="utf-8")
    if sys.stdout is None:
        sys.stdout = stream
    if sys.stderr is None:
        sys.stderr = stream


def _hide_console_and_log():
    """コンソールを隠し、起動時の状態をログに残す。

    うまくいかなかったときに何が起きたのかを追えるよう、
    コンソールの有無にかかわらずログファイルに必ず書く。
    """
    import os

    from stream_highlight import code_version, paths

    had_console, hidden = paths.hide_own_console()
    lines = [
        "--- 起動 %s ---" % datetime.now().isoformat(timespec="seconds"),
        "コード      : %s" % code_version(),
        "実行ファイル: %s" % sys.executable,
        "コンソール  : %s" % (
            "隠した" if hidden else
            "あり（他と共有しているため、そのまま）" if had_console else
            "なし"),
    ]
    for line in lines:
        print(line, flush=True)

    # 画面にコンソールが出ている場合、print はそちらに行ってしまう。
    # 後から状況を追えるよう、ログファイルには必ず残す。
    try:
        with open(os.path.join(paths.data_dir(), "app.log"), "a",
                  encoding="utf-8", errors="replace") as fh:
            fh.write("\n".join(lines) + "\n")
    except OSError:
        pass


def _no_window():
    """子プロセスの黒い画面を出さないための指定。"""
    import subprocess

    return getattr(subprocess, "CREATE_NO_WINDOW", 0)


def _run(args, timeout=600):
    """コマンドを黒い画面なしで実行し、出力をログに残す。

    戻り値は (成功したか, 出力)。失敗しても例外にしないのは、
    更新に失敗しただけで起動を諦めたくないため。
    """
    import os
    import subprocess

    env = dict(os.environ)
    # 認証を聞かれると、画面が無いまま永久に待ち続ける。聞かずに失敗させる。
    env["GIT_TERMINAL_PROMPT"] = "0"
    env["GCM_INTERACTIVE"] = "never"
    try:
        done = subprocess.run(
            args, capture_output=True, text=True, encoding="utf-8",
            errors="replace", timeout=timeout, env=env,
            creationflags=_no_window(),
        )
    except FileNotFoundError:
        return False, "%s が見つかりません" % args[0]
    except subprocess.TimeoutExpired:
        return False, "%s が %d 秒で終わりませんでした" % (args[0], timeout)
    output = (done.stdout or "") + (done.stderr or "")
    return done.returncode == 0, output.strip()


def _startup_update():
    """起動のたびに更新を取り込み、足りないライブラリを入れる。

    bat が担っていた仕事をこちらに移した。bat だと、非表示で動かしたときに
    どこで止まったのか分からなくなる経路がいくつもあった。
    """
    import os

    root = os.path.dirname(os.path.abspath(__file__))

    if os.path.isdir(os.path.join(root, ".git")):
        ok, out = _run(["git", "-C", root, "diff", "--quiet"], timeout=60)
        if not ok:
            print("手元に変更があるため、更新は取り込みません。", flush=True)
        else:
            ok, out = _run(["git", "-C", root, "pull", "--ff-only"], timeout=180)
            print("更新: %s" % (out or ("成功" if ok else "失敗")), flush=True)

    try:
        import fastapi, uvicorn, yt_dlp      # noqa: F401
        return
    except ImportError as exc:
        print("ライブラリが足りません（%s）。入れます。" % exc, flush=True)

    req = os.path.join(root, "stream_highlight", "requirements.txt")
    ok, out = _run([sys.executable, "-m", "pip", "install", "-r", req])
    print("インストール: %s" % ("成功" if ok else "失敗"), flush=True)
    if not ok:
        print(out, flush=True)
        raise SystemExit("ライブラリを入れられませんでした。")


def _show_log():
    """失敗したことに気づけるよう、ログを開く。

    画面を出さない起動では、例外を吐いても誰にも見えない。
    """
    import os
    import subprocess

    from stream_highlight import paths

    path = os.path.join(paths.data_dir(), "app.log")
    if sys.platform != "win32":
        print("ログ: %s" % path, flush=True)
        return
    try:
        subprocess.Popen(["notepad.exe", path])
    except OSError:
        pass


def main():
    if getattr(sys, "frozen", False) and len(sys.argv) > 1 and sys.argv[1] == "--ytdlp":
        import yt_dlp
        sys.exit(yt_dlp.main(sys.argv[2:]))

    # --update はこちらの引数。サーバの引数解析に渡さないよう先に抜く。
    updating = "--update" in sys.argv
    if updating:
        sys.argv = [a for a in sys.argv if a != "--update"]

    # pythonw.exe では標準出力の行き先が無いので、先にログへ向ける。
    _redirect_output()
    if getattr(sys, "frozen", False):
        _hide_console_and_log()

    try:
        if updating:
            print("--- 起動 %s ---" % datetime.now().isoformat(timespec="seconds"),
                  flush=True)
            _startup_update()
        from stream_highlight.server import main as serve
        serve()
    except SystemExit:
        raise
    except BaseException:
        import traceback

        traceback.print_exc()
        try:
            sys.stdout.flush()
        except Exception:      # noqa: BLE001
            pass
        if updating or getattr(sys, "frozen", False):
            _show_log()
        raise


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
