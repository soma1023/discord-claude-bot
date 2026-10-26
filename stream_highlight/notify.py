# -*- coding: utf-8 -*-
"""長い処理の進み具合を Discord に知らせる（任意）。

.env の STREAM_HIGHLIGHT_DISCORD_WEBHOOK に Webhook の URL があるときだけ投稿する。
無いとき・送れなかったときは何もしない。通知の失敗で本来の処理を止めないため。
"""

import json
import urllib.request

from . import retake

WEBHOOK_ENV = "STREAM_HIGHLIGHT_DISCORD_WEBHOOK"
# Discord は既定の Python の User-Agent を弾くことがあるので、名乗っておく
USER_AGENT = "DiscordBot (https://github.com/soma1023/discord-claude-bot, 1.0)"
MAX_CHARS = 1900            # 1投稿の上限（2000字）より少し手前で切る
_ALLOWED = ("https://discord.com/api/webhooks/", "https://discordapp.com/api/webhooks/",
            "https://ptb.discord.com/api/webhooks/", "https://canary.discord.com/api/webhooks/")


def webhook_url():
    url = retake.setting(WEBHOOK_ENV).strip()
    # 打ち間違いで別のサイトに送ってしまわないよう、Discord の Webhook だけを受け付ける
    return url if url.startswith(_ALLOWED) else ""


def enabled():
    return bool(webhook_url())


def post(text, url=None, opener=urllib.request.urlopen):
    """1件投稿する。送れたら True。"""
    url = url or webhook_url()
    if not url or not text:
        return False
    if len(text) > MAX_CHARS:
        text = text[:MAX_CHARS - 1] + "…"
    request = urllib.request.Request(
        url, data=json.dumps({"content": text}).encode("utf-8"),
        headers={"Content-Type": "application/json", "User-Agent": USER_AGENT})
    try:
        with opener(request, timeout=10):
            return True
    except Exception:          # noqa: BLE001（回線・URL の誤り・Discord 側の不調）
        return False
