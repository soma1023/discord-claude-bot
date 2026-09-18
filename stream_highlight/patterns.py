# -*- coding: utf-8 -*-
"""コメントの正規化と分類ルール。

チャット欄の反応を「草」「驚き」などのカテゴリに振り分ける。
1つのコメントが複数カテゴリに入ることもある（例: 「まじかwww」→ 草 + 驚き）。
"""

import re
import unicodedata

# 同じ文字の繰り返しを3文字までに潰す（wwwwww → www、！！！！ → ！！！）
_RUNS = re.compile(r"(.)\1{2,}")
_SPACES = re.compile(r"\s+")


def normalize(text):
    """コメントを比較用に正規化する。全角半角・大文字小文字・連打を吸収する。"""
    t = unicodedata.normalize("NFKC", text or "")
    t = t.strip().lower()
    t = _SPACES.sub(" ", t)
    return _RUNS.sub(r"\1\1\1", t)


# カテゴリ定義。patterns は normalize() 済みの文字列に対して当てる。
# 配信サービスが出すシステム通知。本人の発言ではないので見せ場から外す。
# 取得時ではなく解析時に判定するので、保存済みのチャットにも後から効く。
SYSTEM_NOTICE = re.compile(
    r"subscribed with Prime"
    r"|subscribed at Tier \d"
    r"|(?:T|t)hey've subscribed for \d+ month"
    r"|is continuing the Gift Sub"
    r"|gifted a Tier \d+ Sub"
    r"|is gifting \d+ Tier \d+ Sub"
    r"|watched \d+ consecutive stream"
    r"|sparked a watch streak"
    r"|\d+ raiders? from"
    r"|converted from a Prime Sub",
    re.IGNORECASE,
)


def split_notice(text):
    """システム通知と、本人が書いたコメントを切り分ける。

    通知のあとに本人のひとことが続くことがある
    （例: 「... 43 months! ajak0nHi」の末尾）。その部分は本物なので残す。
    戻り値は (本人のコメント, 通知だったか)。
    """
    found = SYSTEM_NOTICE.search(text or "")
    if not found:
        return text, False
    tail = text[found.end():]
    mark = tail.find("!")
    return (tail[mark + 1:].strip() if mark >= 0 else ""), True


CATEGORIES = [
    {
        "id": "laugh",
        "label": "草",
        "color": "#7ee787",
        "patterns": [
            r"(?<![a-z0-9.])w{2,}(?![a-z0-9.])",   # www / ｗｗｗ（単独の連打）
            # Twitchのチャンネル絵文字は ○○Www の形が多いので、語中でも拾う。
            # URLの www は直後が . なので当たらない。
            r"(?<![./])w{3,}(?![.])",
            r"草|くさ|大草原|草原|草生",
            r"笑|わろ|爆笑|ぷぷ",
            r"\blol\b|\blmao\b|\brofl\b|\bxd\b",
            r"lul|kek|omegalul|icant|lmfao|pepelaugh|lolw",
        ],
    },
    {
        "id": "surprise",
        "label": "驚き",
        "color": "#ffa657",
        "patterns": [
            r"えぇ|ええ|えっ|うぉ|おぉ|おお|ふぁ|ファ",
            r"まじ|マジ|嘘|うそ|ウソ|やば|ヤバ|えぐ|ちょ|待って|まって",
            r"は\?|なんだと|何それ|なにそれ|どうなって",
            r"\bomg\b|\bwtf\b|pog|poggers|pogchamp|monkas",
        ],
    },
    {
        "id": "praise",
        "label": "称賛",
        "color": "#79c0ff",
        "patterns": [
            r"うま|上手|すご|凄|つよ|強い|神|天才|最高|かっこ|かっけ",
            r"ナイス|\bnice\b|\bgg\b|\bgoat\b|\bpogu\b|\bclap\b",
            r"888|ぱちぱち|拍手",
            r"^w$",   # Twitchの「W」（= win）
        ],
    },
    {
        "id": "scream",
        "label": "悲鳴",
        "color": "#ff7b72",
        "patterns": [
            r"うわ|ぎゃ|ひぇ|ヒェ|あぁ|ああ|やめ|いた|痛|こわ|怖",
            r"どんまい|ドンマイ|惜し|おし|残念|ざんねん|おわた|オワタ",
            r"sadge|pepehands|monkaw|\bfeelsbad\b",
        ],
    },
    {
        "id": "cute",
        "label": "かわいい",
        "color": "#ff9ec6",
        "patterns": [
            r"かわい|かわよ|可愛|きゃわ|kawaii|\bcute\b|尊い|とうと",
        ],
    },
    {
        "id": "confused",
        "label": "困惑",
        "color": "#d2a8ff",
        "patterns": [
            r"\?\?|？？|なんで|何で|どういうこと|意味不|わからん|わかんない|謎",
        ],
    },
]

_COMPILED = [
    (c["id"], re.compile("|".join(c["patterns"])))
    for c in CATEGORIES
]

CATEGORY_IDS = [c["id"] for c in CATEGORIES]
CATEGORY_LABELS = {c["id"]: c["label"] for c in CATEGORIES}


def classify(normalized_text):
    """正規化済みコメントが該当するカテゴリIDのリストを返す。"""
    return [cid for cid, rx in _COMPILED if rx.search(normalized_text)]
