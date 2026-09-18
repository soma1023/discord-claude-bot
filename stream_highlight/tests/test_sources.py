# -*- coding: utf-8 -*-
"""チャット取得まわりのパーサを、実際のレスポンス構造を模したデータで確かめる。

ネットワークには出ない。YouTubeの live_chat.json と TwitchのGQL応答は、
それぞれ実物と同じ入れ子構造で組み立てている。
"""

import json
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from stream_highlight import sources
from stream_highlight.sources import FetchError, parse_url


def text_item(offset_ms, author, runs, msg_id="x"):
    return json.dumps({
        "replayChatItemAction": {
            "actions": [{
                "addChatItemAction": {
                    "item": {
                        "liveChatTextMessageRenderer": {
                            "message": {"runs": runs},
                            "authorName": {"simpleText": author},
                            "timestampUsec": "1700000000000000",
                            "id": msg_id,
                        }
                    }
                }
            }],
            "videoOffsetTimeMsec": str(offset_ms),
        }
    }, ensure_ascii=False)


class TestYouTubeParsing(unittest.TestCase):
    def parse(self, line, origin=None):
        return sources._parse_live_chat_line(line, origin)

    def test_plain_message(self):
        msg = self.parse(text_item(125_500, "視聴者A", [{"text": "www"}]))
        self.assertEqual(msg.offset, 125.5)
        self.assertEqual(msg.author, "視聴者A")
        self.assertEqual(msg.text, "www")
        self.assertEqual(msg.kind, "text")

    def test_emoji_runs_are_kept(self):
        msg = self.parse(text_item(1000, "A", [
            {"text": "かわいい"},
            {"emoji": {"emojiId": "abc", "shortcuts": [":_kusa:"], "isCustomEmoji": True}},
        ]))
        self.assertEqual(msg.text, "かわいい:_kusa:")

    def test_negative_offset_for_pre_stream_chat(self):
        """配信開始前の待機所コメントは負のオフセットで来る。"""
        msg = self.parse(text_item(-51966, "A", [{"text": "楽しみ"}]))
        self.assertLess(msg.offset, 0)

    def test_superchat(self):
        line = json.dumps({
            "replayChatItemAction": {
                "actions": [{"addChatItemAction": {"item": {
                    "liveChatPaidMessageRenderer": {
                        "message": {"runs": [{"text": "応援してます"}]},
                        "authorName": {"simpleText": "支援者"},
                        "purchaseAmountText": {"simpleText": "￥1,000"},
                    }}}}],
                "videoOffsetTimeMsec": "60000",
            }
        }, ensure_ascii=False)
        msg = self.parse(line)
        self.assertEqual(msg.kind, "paid")
        self.assertEqual(msg.amount, "￥1,000")
        self.assertEqual(msg.offset, 60.0)

    def test_membership_without_message_body(self):
        line = json.dumps({
            "replayChatItemAction": {
                "actions": [{"addChatItemAction": {"item": {
                    "liveChatMembershipItemRenderer": {
                        "authorName": {"simpleText": "新メンバー"},
                        "headerSubtext": {"runs": [{"text": "メンバーになりました"}]},
                    }}}}],
                "videoOffsetTimeMsec": "90000",
            }
        }, ensure_ascii=False)
        msg = self.parse(line)
        self.assertEqual(msg.kind, "member")
        self.assertEqual(msg.text, "メンバーになりました")

    def test_fallback_to_timestamp_when_offset_missing(self):
        """videoOffsetTimeMsec が無い形式では、先頭からの相対時刻で補う。"""
        line = json.dumps({
            "actions": [{"addChatItemAction": {"item": {
                "liveChatTextMessageRenderer": {
                    "message": {"runs": [{"text": "草"}]},
                    "authorName": {"simpleText": "A"},
                    "timestampUsec": "1700000030000000",
                }}}}]
        }, ensure_ascii=False)
        msg = self.parse(line, origin=1700000000000000)
        self.assertEqual(msg.offset, 30.0)

    def test_ignores_non_chat_lines(self):
        self.assertIsNone(self.parse("こわれた行"))
        self.assertIsNone(self.parse(json.dumps({"replayChatItemAction": {"actions": []}})))
        self.assertIsNone(self.parse(json.dumps({
            "replayChatItemAction": {
                "actions": [{"markChatItemAsDeletedAction": {"targetItemId": "x"}}],
                "videoOffsetTimeMsec": "1000"}})))

    def test_survives_broken_shapes(self):
        """実データで起きうる「入れ子がnull」に落とされないこと。"""
        broken = [
            "null",
            "[]",
            '"文字列だけ"',
            json.dumps({"replayChatItemAction": None}),
            json.dumps({"replayChatItemAction": {"actions": None}}),
            json.dumps({"replayChatItemAction": {"actions": [None], "videoOffsetTimeMsec": "1000"}}),
            json.dumps({"replayChatItemAction": {"actions": [{"addChatItemAction": None}],
                                                 "videoOffsetTimeMsec": "1000"}}),
            json.dumps({"replayChatItemAction": {"actions": [{"addChatItemAction": {"item": None}}],
                                                 "videoOffsetTimeMsec": "1000"}}),
            json.dumps({"replayChatItemAction": {
                "actions": [{"addChatItemAction": {"item": {"liveChatTextMessageRenderer": None}}}],
                "videoOffsetTimeMsec": "1000"}}),
            json.dumps({"replayChatItemAction": {
                "actions": [{"addChatItemAction": {"item": {
                    "liveChatTextMessageRenderer": {"message": None,
                                                    "authorName": None}}}}],
                "videoOffsetTimeMsec": "1000"}}),
            json.dumps({"replayChatItemAction": {
                "actions": [{"addChatItemAction": {"item": {
                    "liveChatTextMessageRenderer": {"message": {"runs": None}}}}}],
                "videoOffsetTimeMsec": "1000"}}),
            json.dumps({"replayChatItemAction": {
                "actions": [{"addChatItemAction": {"item": {
                    "liveChatTextMessageRenderer": {"message": {"runs": [None, {"emoji": None},
                                                                        {"text": None}]}}}}}],
                "videoOffsetTimeMsec": "1000"}}),
            json.dumps({"replayChatItemAction": {
                "actions": [{"addChatItemAction": {"item": {
                    "liveChatTextMessageRenderer": {"message": {"runs": [{"text": "ok"}]}}}}}],
                "videoOffsetTimeMsec": "こわれた値"}}),
        ]
        for line in broken:
            self.parse(line)            # 例外を出さないこと自体が要件

    def test_unknown_renderer_types_are_skipped(self):
        """広告・アンケート・プレースホルダなど、発言以外の要素を無視すること。"""
        for renderer in ("liveChatViewerEngagementMessageRenderer",
                         "liveChatPlaceholderItemRenderer",
                         "liveChatPollRenderer",
                         "liveChatBannerRenderer"):
            line = json.dumps({"replayChatItemAction": {
                "actions": [{"addChatItemAction": {"item": {renderer: {"id": "x"}}}}],
                "videoOffsetTimeMsec": "1000"}})
            self.assertIsNone(self.parse(line), renderer)

    def test_non_chat_action_types_are_skipped(self):
        for action in ("markChatItemAsDeletedAction", "addLiveChatTickerItemAction",
                       "replaceChatItemAction"):
            line = json.dumps({"replayChatItemAction": {
                "actions": [{action: {"id": "x"}}], "videoOffsetTimeMsec": "1000"}})
            self.assertIsNone(self.parse(line), action)

    def test_gift_author_from_header(self):
        """ギフト告知は名前がheaderの中にあるので、そこから拾うこと。"""
        line = json.dumps({"replayChatItemAction": {
            "actions": [{"addChatItemAction": {"item": {
                "liveChatSponsorshipsGiftPurchaseAnnouncementRenderer": {
                    "header": {"liveChatSponsorshipsHeaderRenderer": {
                        "authorName": {"simpleText": "太っ腹さん"},
                        "primaryText": {"runs": [{"text": "5個ギフトしました"}]}}}}}}}],
            "videoOffsetTimeMsec": "5000"}}, ensure_ascii=False)
        msg = self.parse(line)
        self.assertEqual(msg.author, "太っ腹さん")
        self.assertEqual(msg.kind, "gift")

    def test_subprocess_suppresses_console_window(self):
        """Windowsで子プロセスの黒い画面が開かないように指定していること。"""
        import subprocess

        import yt_dlp  # noqa: F401  差し替える前に読み込ませておく

        captured = {}

        # subprocess.Popen を継承しているコードがあるので、関数ではなくクラスで差し替える
        class FakePopen:
            def __init__(self, cmd, **kwargs):
                captured.update(kwargs)
                self.stdout = []
                self.returncode = 0

            def wait(self):
                return 0

        original = subprocess.Popen
        subprocess.Popen = FakePopen
        try:
            sources._run_ytdlp(["--version"])
        finally:
            subprocess.Popen = original

        self.assertIn("creationflags", captured)
        self.assertEqual(captured["creationflags"], sources._NO_CONSOLE)

    def test_error_messages_are_readable(self):
        self.assertIn("メンバー限定", sources._ytdlp_error(["ERROR: Join this channel members-only"]))
        self.assertIn("見つかりません", sources._ytdlp_error(["ERROR: Video unavailable"]))


class TestTwitchParsing(unittest.TestCase):
    def setUp(self):
        # 再試行の待ち時間でテストを遅くしない
        self._wait = sources._NULL_RETRY_WAIT
        sources._NULL_RETRY_WAIT = 0

    def tearDown(self):
        sources._NULL_RETRY_WAIT = self._wait

    def test_node_to_message(self):
        node = {
            "id": "c1",
            "contentOffsetSeconds": 742,
            "commenter": {"displayName": "viewer1"},
            "message": {"fragments": [{"text": "LUL "}, {"text": "KEKW"}]},
        }
        msg = sources._node_to_message(node)
        self.assertEqual(msg.offset, 742.0)
        self.assertEqual(msg.author, "viewer1")
        self.assertEqual(msg.text, "LUL KEKW")

    def test_missing_commenter_is_tolerated(self):
        """BAN済みユーザーなどは commenter が null で来ることがある。"""
        msg = sources._node_to_message({"contentOffsetSeconds": 5, "commenter": None,
                                        "message": {"fragments": [{"text": "hi"}]}})
        self.assertEqual(msg.author, "")
        self.assertEqual(msg.text, "hi")

    def test_system_notice_is_separated(self):
        """サブスク告知などは本人の発言ではないので切り分けること。"""
        cases = [
            ("でいじ subscribed with Prime. They've subscribed for 43 months! ajak0nHi",
             "ajak0nHi", "text"),
            ("まじょ_ subscribed at Tier 1. They've subscribed for 13 months, "
             "currently on a 9 month streak!", "", "system"),
            ("raro_ga watched 55 consecutive streams and sparked a watch streak!",
             "", "system"),
            ("someone gifted a Tier 1 Sub to viewer!", "", "system"),
            ("15 raiders from otherchannel have joined!", "", "system"),
        ]
        for text, expected_body, expected_kind in cases:
            node = {"id": "x", "contentOffsetSeconds": 1,
                    "commenter": {"displayName": "u"},
                    "message": {"fragments": [{"text": text}]}}
            msg = sources._node_to_message(node)
            self.assertEqual(msg.text, expected_body, text)
            self.assertEqual(msg.kind, expected_kind, text)

    def test_normal_comments_are_untouched(self):
        for text in ("それもえぐいww", "subscribe しようかな", "普通のコメント",
                     "ajak0nWww", "888888"):
            node = {"id": "x", "contentOffsetSeconds": 1,
                    "commenter": {"displayName": "u"},
                    "message": {"fragments": [{"text": text}]}}
            msg = sources._node_to_message(node)
            self.assertEqual(msg.text, text)
            self.assertEqual(msg.kind, "text")

    def test_request_shapes(self):
        first = sources._comment_request("123", offset=600)
        self.assertEqual(first["variables"]["contentOffsetSeconds"], 600)
        self.assertNotIn("cursor", first["variables"])
        nxt = sources._comment_request("123", cursor="abc")
        self.assertEqual(nxt["variables"]["cursor"], "abc")
        self.assertNotIn("contentOffsetSeconds", nxt["variables"])
        self.assertEqual(nxt["extensions"]["persistedQuery"]["sha256Hash"], sources._GQL_HASH)

    def test_segment_paging_and_dedupe(self):
        """ページングが進み、区間の外に出たら止まり、重複が除かれること。"""
        pages = {
            None: {"edges": [
                {"cursor": "c1", "node": {"id": "a", "contentOffsetSeconds": 1,
                                          "commenter": {"displayName": "u"},
                                          "message": {"fragments": [{"text": "1"}]}}},
                {"cursor": "c2", "node": {"id": "b", "contentOffsetSeconds": 2,
                                          "commenter": {"displayName": "u"},
                                          "message": {"fragments": [{"text": "2"}]}}},
            ], "pageInfo": {"hasNextPage": True}},
            "c2": {"edges": [
                {"cursor": "c3", "node": {"id": "b", "contentOffsetSeconds": 2,
                                          "commenter": {"displayName": "u"},
                                          "message": {"fragments": [{"text": "2"}]}}},
                {"cursor": "c4", "node": {"id": "c", "contentOffsetSeconds": 99,
                                          "commenter": {"displayName": "u"},
                                          "message": {"fragments": [{"text": "out"}]}}},
            ], "pageInfo": {"hasNextPage": True}},
        }
        calls = []

        def fake_post(payload, retries=4, **kwargs):
            cursor = payload[0]["variables"].get("cursor")
            calls.append(cursor)
            return [{"data": {"video": {"comments": pages[cursor]}}}]

        original = sources._gql_post
        sources._gql_post = fake_post
        try:
            import threading
            collected, seen = [], set()
            sources._fetch_twitch_segment("1", 0, 50, seen, threading.Lock(),
                                          collected, lambda _: None)
        finally:
            sources._gql_post = original

        self.assertEqual(calls, [None, "c2"])
        self.assertEqual([m.text for m in collected], ["1", "2"])   # 重複と区間外を除外

    def test_null_comments_does_not_crash(self):
        """comments が null でも落ちず、理由の分かるエラーになること。"""
        import threading
        original = sources._gql_post
        sources._gql_post = lambda payload, retries=4, **kw: [
            {"data": {"video": {"id": "1", "comments": None}}}]
        failures = []
        try:
            sources._fetch_twitch_segment("1", 0, None, set(), threading.Lock(),
                                          [], lambda _: None, failures)
        finally:
            sources._gql_post = original
        self.assertTrue(failures)
        self.assertIn("Twitchがチャットを返しませんでした", failures[0])

    def test_graphql_errors_are_surfaced(self):
        """Twitchが返したエラー本文を、そのまま利用者に見せること。"""
        comments, reason = sources._extract_comments(
            [{"errors": [{"message": "service timeout"}]}])
        self.assertIsNone(comments)
        self.assertIn("service timeout", reason)

    def test_deleted_video_message(self):
        comments, reason = sources._extract_comments([{"data": {"video": None}}])
        self.assertIsNone(comments)
        self.assertIn("見つかりません", reason)

    def test_malformed_payloads(self):
        for payload in (None, [], [None], "文字列", [{}], [{"data": None}]):
            comments, reason = sources._extract_comments(payload)
            self.assertIsNone(comments)
            self.assertTrue(reason)

    def test_partial_failure_keeps_what_was_collected(self):
        """一部の区間が取れなくても、取れた分は返すこと。"""
        calls = {"n": 0}

        def flaky(payload, retries=4, **kwargs):
            calls["n"] += 1
            if calls["n"] <= 2:      # 1回目は接続確認、2回目が実際の取得
                return [{"data": {"video": {"comments": {
                    "edges": [{"cursor": "c1", "node": {
                        "id": "a", "contentOffsetSeconds": 1,
                        "commenter": {"displayName": "u"},
                        "message": {"fragments": [{"text": "残るコメント"}]}}}],
                    "pageInfo": {"hasNextPage": False}}}}}]
            return [{"data": {"video": {"comments": None}}}]

        original = sources._gql_post
        sources._gql_post = flaky
        try:
            messages = sources.fetch_twitch_chat("1", duration=0)
        finally:
            sources._gql_post = original
        self.assertEqual([m.text for m in messages], ["残るコメント"])

    def test_probe_falls_back_to_second_client_id(self):
        """1つ目のClient-IDが弾かれたら、2つ目で取りに行くこと。"""
        used = []

        def by_client(payload, retries=4, client_id=None, **kwargs):
            used.append(client_id)
            if client_id == sources._GQL_CLIENT_IDS[0]:
                return [{"errors": [{"message": "failed integrity check"}]}]
            return [{"data": {"video": {"comments": {"edges": [], "pageInfo": {}}}}}]

        original = sources._gql_post
        sources._gql_post = by_client
        try:
            chosen = sources.probe_twitch("1")
        finally:
            sources._gql_post = original
        self.assertEqual(chosen, sources._GQL_CLIENT_IDS[1])
        self.assertEqual(used, list(sources._GQL_CLIENT_IDS))

    def test_probe_reports_every_client_id_that_failed(self):
        """全部だめなら、どのIDがどう失敗したかを添えて知らせること。"""
        original = sources._gql_post
        sources._gql_post = lambda payload, retries=4, **kw: [
            {"errors": [{"message": "failed integrity check"}]}]
        try:
            with self.assertRaises(FetchError) as ctx:
                sources.probe_twitch("1")
        finally:
            sources._gql_post = original
        detail = str(ctx.exception)
        self.assertIn("failed integrity check", detail)
        for client_id in sources._GQL_CLIENT_IDS:
            self.assertIn(client_id[:10], detail)

    def test_request_headers_match_twitch_client(self):
        """Twitch自身のクライアントと同じヘッダーで投げていること。"""
        import urllib.request
        captured = {}

        def fake_urlopen(req, timeout=None):
            captured["headers"] = dict(req.header_items())
            raise urllib.error.URLError("停止")

        original = urllib.request.urlopen
        urllib.request.urlopen = fake_urlopen
        try:
            with self.assertRaises(FetchError):
                sources._gql_post([{"operationName": "x"}], retries=1)
        finally:
            urllib.request.urlopen = original
        headers = {k.lower(): v for k, v in captured["headers"].items()}
        self.assertEqual(headers["Content-type".lower()], "text/plain;charset=UTF-8")
        self.assertEqual(headers["Client-id".lower()], sources._GQL_CLIENT_IDS[0])
        self.assertIn("twitch.tv", headers["Origin".lower()])

    def test_all_segments_empty_raises_with_reason(self):
        original = sources._gql_post
        sources._gql_post = lambda payload, retries=4, **kw: [
            {"errors": [{"message": "failed integrity check"}]}]
        try:
            with self.assertRaises(FetchError) as ctx:
                sources.fetch_twitch_chat("1", duration=0)
        finally:
            sources._gql_post = original
        self.assertIn("failed integrity check", str(ctx.exception))

    def test_missing_video_is_reported(self):
        import threading
        original = sources._gql_post
        sources._gql_post = lambda payload, retries=4, **kw: [{"data": {"video": None}}]
        failures = []
        try:
            sources._fetch_twitch_segment("1", 0, None, set(), threading.Lock(),
                                          [], lambda _: None, failures)
        finally:
            sources._gql_post = original
        self.assertTrue(failures)
        self.assertIn("削除済み", failures[0])


class TestUrls(unittest.TestCase):
    def test_youtube_forms(self):
        for url in ["https://www.youtube.com/watch?v=abcdefghijk",
                    "https://youtu.be/abcdefghijk?t=90",
                    "https://www.youtube.com/live/abcdefghijk",
                    "https://m.youtube.com/watch?app=desktop&v=abcdefghijk",
                    "abcdefghijk"]:
            self.assertEqual(parse_url(url), ("youtube", "abcdefghijk"), url)

    def test_twitch_forms(self):
        for url in ["https://www.twitch.tv/videos/123456789",
                    "https://twitch.tv/someone/video/123456789",
                    "https://www.twitch.tv/videos/123456789?t=1h2m3s"]:
            self.assertEqual(parse_url(url), ("twitch", "123456789"), url)

    def test_rejects_others(self):
        for url in ["", "https://example.com/", "https://www.nicovideo.jp/watch/sm9"]:
            with self.assertRaises(FetchError):
                parse_url(url)

    def test_time_url(self):
        yt = sources.StreamInfo("youtube", "abcdefghijk", "u")
        self.assertEqual(yt.time_url(3725.9), "https://www.youtube.com/watch?v=abcdefghijk&t=3725s")
        tw = sources.StreamInfo("twitch", "123", "u")
        self.assertEqual(tw.time_url(3725), "https://www.twitch.tv/videos/123?t=1h2m5s")
        self.assertEqual(tw.time_url(-5), "https://www.twitch.tv/videos/123?t=0h0m0s")


if __name__ == "__main__":
    unittest.main(verbosity=2)
