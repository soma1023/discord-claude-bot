# -*- coding: utf-8 -*-
"""起動の入口（run_app.py）の振る舞いを確かめる。

非表示で起動すると、失敗しても画面に何も出ない。そのため
「黙って何も起きない」を生む経路をここで固定しておく。
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import run_app
from stream_highlight import server


class TestRunCommand(unittest.TestCase):
    def test_success_returns_output(self):
        ok, out = run_app._run([sys.executable, "-c", "print('hi')"])
        self.assertTrue(ok)
        self.assertEqual(out, "hi")

    def test_missing_command_is_not_an_exception(self):
        """コマンドが無くても落とさない。更新の失敗で起動を諦めないため。"""
        ok, out = run_app._run(["sh_definitely_not_a_command"])
        self.assertFalse(ok)
        self.assertIn("見つかりません", out)

    def test_failure_returns_output(self):
        ok, out = run_app._run([sys.executable, "-c", "import sys; sys.exit(3)"])
        self.assertFalse(ok)

    def test_git_is_told_not_to_ask(self):
        """認証を聞かれると、画面が無いまま永久に待つ。聞かせない。"""
        ok, out = run_app._run([
            sys.executable, "-c",
            "import os; print(os.environ['GIT_TERMINAL_PROMPT'])"])
        self.assertTrue(ok)
        self.assertEqual(out, "0")

    def test_timeout_is_reported(self):
        ok, out = run_app._run([sys.executable, "-c", "import time; time.sleep(5)"],
                               timeout=1)
        self.assertFalse(ok)
        self.assertIn("終わりませんでした", out)


class TestUpdateFlag(unittest.TestCase):
    """--update は入口の引数。サーバの引数解析に渡してはいけない。"""

    def setUp(self):
        self.argv = sys.argv[:]
        self.served = []
        self.updated = []
        self.real_update = run_app._startup_update
        self.real_serve = server.main
        run_app._startup_update = lambda: self.updated.append(True)
        server.main = lambda: self.served.append(sys.argv[:])

    def tearDown(self):
        sys.argv = self.argv
        run_app._startup_update = self.real_update
        server.main = self.real_serve

    def test_update_flag_is_removed_before_serving(self):
        sys.argv = ["run_app.py", "--update", "--no-browser"]
        run_app.main()
        self.assertEqual(self.updated, [True])
        self.assertEqual(self.served, [["run_app.py", "--no-browser"]])

    def test_without_flag_no_update_runs(self):
        sys.argv = ["run_app.py", "--no-browser"]
        run_app.main()
        self.assertEqual(self.updated, [])
        self.assertEqual(self.served, [["run_app.py", "--no-browser"]])


if __name__ == "__main__":
    unittest.main(verbosity=2)
