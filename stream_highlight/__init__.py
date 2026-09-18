"""配信アーカイブから「見せ場」候補を抽出するツール。"""

__version__ = "0.1.0"


import os


def code_version():
    """いま動いているコードがどのコミットかを返す。

    サーバを再起動し忘れると、画面だけ新しくコードが古いという状態になり、
    原因の切り分けが難しくなる。画面に出して一目で分かるようにする。
    """
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
    return __version__
