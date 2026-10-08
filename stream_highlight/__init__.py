"""配信アーカイブから「見せ場」候補を抽出するツール。"""

__version__ = "0.1.0"


import os


def _read_build_id():
    """exe化したときに埋め込まれるファイルから版を読む。"""
    from . import paths

    for base in (paths.bundle_dir(), os.path.dirname(os.path.abspath(__file__))):
        stamp = os.path.join(base, "_build_id.txt")
        if not os.path.exists(stamp):
            continue
        try:
            with open(stamp, encoding="utf-8") as fh:
                value = fh.read().strip()
            if value:
                return value
        except OSError:
            pass
    return ""


def _read_git_head():
    """リポジトリから、いまチェックアウトしている版を読む。"""
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    git = os.path.join(root, ".git")
    try:
        with open(os.path.join(git, "HEAD"), encoding="utf-8") as fh:
            head = fh.read().strip()
        if not head.startswith("ref:"):
            return head[:7]
        ref = head.split(" ", 1)[1].strip()
        path = os.path.join(git, ref)
        if os.path.exists(path):
            with open(path, encoding="utf-8") as fh:
                return fh.read().strip()[:7]
        # ref がファイルとして無い場合は packed-refs を見る
        with open(os.path.join(git, "packed-refs"), encoding="utf-8") as fh:
            for line in fh:
                if line.rstrip().endswith(" " + ref):
                    return line.split(" ", 1)[0][:7]
    except OSError:
        pass
    return ""


def code_version():
    """いま動いているコードがどのコミットかを返す。

    サーバを再起動し忘れると、画面だけ新しくコードが古いという状態になり、
    原因の切り分けが難しくなる。画面に出して一目で分かるようにする。

    リポジトリから動かしているときは .git が正しい。
    exeのビルド時に作られる _build_id.txt はビルド後も残るため、
    これを優先すると、更新しても古い版を表示し続けてしまう。
    """
    from . import paths

    if paths.is_frozen():
        return _read_build_id() or __version__
    return _read_git_head() or _read_build_id() or __version__
