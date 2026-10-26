# -*- coding: utf-8 -*-
"""言い直しを見つけて、その「前の側」を消す範囲を返す（任意機能・既定OFF）。

「今日はホラーゲ…あ、今日はホラーゲームをやります」のような言い直しは、
時刻付きの文字起こしを Claude API に渡して判定してもらう。

.env に ANTHROPIC_API_KEY があるときだけ使える。失敗しても言い直しの削除を
諦めるだけで、書き出し自体は続ける（呼び出し側が RetakeError を受け止める）。

使うモデルは .env の STREAM_HIGHLIGHT_RETAKE_MODEL で変えられる。
"""

import json
import os

from .transcribe import join_words

DEFAULT_MODEL = "claude-opus-5-5"
MODEL_ENV = "STREAM_HIGHLIGHT_RETAKE_MODEL"
KEY_ENV = "ANTHROPIC_API_KEY"
UTTERANCE_GAP_SEC = 0.3        # これ以上の間で行を分けて渡す（読みやすくするため）
MAX_REMOVE_RATIO = 0.3         # 3割以上消せと言われたら、判定がおかしいと見なして使わない

# 拒否されたときに別のモデルでやり直してくれる仕組みに対応しているモデル
_FALLBACK_MODELS = ("claude-opus-5-5", "claude-opus-5", "claude-fable-5-1", "claude-sonnet-5-5")

PROMPT = """以下はゲーム配信の切り抜き区間の文字起こしです。各語の前の [番号] は語の番号です。
配信者が言い直している箇所（言いかけてやめ、同じ内容を言い直している所）を探し、
消すべき「前の側」（言い損じた方）の語の範囲を返してください。

- 言い直した後の側は残します。消すのは前の側だけです。
- 確信が持てない所は返さないでください。単なる繰り返しや強調（「やばいやばい」など）は言い直しではありません。
- 見つからなければ空の配列を返してください。

{transcript}"""

SCHEMA = {
    "type": "object",
    "properties": {
        "remove": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "first": {"type": "integer"},
                    "last": {"type": "integer"},
                },
                "required": ["first", "last"],
                "additionalProperties": False,
            },
        },
    },
    "required": ["remove"],
    "additionalProperties": False,
}


class RetakeError(Exception):
    """言い直しの削除だけを諦める失敗。メッセージは画面に出す。"""


def _env_files():
    """.env を探す場所。リポジトリの直下と、保存先フォルダ（exe 版）。

    .env.for-partner などの派生ファイルは読まない。
    """
    from . import paths

    here = os.path.dirname(os.path.abspath(__file__))
    return [os.path.join(os.path.dirname(here), ".env"),
            os.path.join(paths.data_dir(), ".env")]


def _read_env_file(path):
    values = {}
    try:
        with open(path, encoding="utf-8-sig") as fh:
            for line in fh:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                if key.startswith("export "):
                    key = key[7:].strip()
                values[key] = value.strip().strip('"').strip("'")
    except OSError:
        pass
    return values


def setting(name):
    """環境変数 → .env の順に探す。"""
    if os.environ.get(name):
        return os.environ[name]
    for path in _env_files():
        value = _read_env_file(path).get(name)
        if value:
            return value
    return ""


def available():
    """使えるなら空文字、使えないなら理由（日本語）。"""
    try:
        import anthropic  # noqa: F401
    except ImportError:
        return "anthropic が入っていません（pip install anthropic）。"
    if not setting(KEY_ENV):
        return ".env に ANTHROPIC_API_KEY がありません。"
    return ""


def model_name():
    return setting(MODEL_ENV) or DEFAULT_MODEL


def format_transcript(words):
    """語に番号を付け、間で行を分けたテキストにする。"""
    lines = []
    current = []
    last_end = None
    for index, w in enumerate(words):
        if current and last_end is not None and w.start - last_end >= UTTERANCE_GAP_SEC:
            lines.append(" ".join(current))
            current = []
        if not current:
            current.append("(%.1f秒)" % w.start)
        current.append("[%d]%s" % (index, w.text))
        last_end = w.end
    if current:
        lines.append(" ".join(current))
    return "\n".join(lines)


def ranges_from_answer(words, answer):
    """API の答え（語番号の範囲）を、消す時刻の範囲に直す。おかしな答えは捨てる。"""
    picked = set()
    for item in answer.get("remove") or []:
        try:
            first, last = int(item["first"]), int(item["last"])
        except (KeyError, TypeError, ValueError):
            continue
        if first > last:
            first, last = last, first
        if first < 0 or last >= len(words):
            continue
        picked.update(range(first, last + 1))
    if not picked:
        return []
    if len(picked) > len(words) * MAX_REMOVE_RATIO:
        raise RetakeError("言い直しの判定が多すぎるため、使いませんでした。")
    ranges = []
    for index in sorted(picked):
        start, end = words[index].start, words[index].end
        if ranges and index - 1 in picked and ranges[-1][2] == index - 1:
            ranges[-1] = (ranges[-1][0], end, index)
        else:
            ranges.append((start, end, index))
    return [(lo, hi) for lo, hi, _ in ranges]


def detect(words, client=None):
    """消す範囲 [(開始秒, 終了秒)] を返す。失敗したら RetakeError。"""
    if not words:
        return []
    reason = available() if client is None else ""
    if reason:
        raise RetakeError(reason)
    import anthropic

    if client is None:
        client = anthropic.Anthropic(api_key=setting(KEY_ENV))
    model = model_name()
    request = dict(
        model=model, max_tokens=16000,
        output_config={"format": {"type": "json_schema", "schema": SCHEMA}},
        messages=[{"role": "user",
                   "content": PROMPT.format(transcript=format_transcript(words))}],
    )
    try:
        if model in _FALLBACK_MODELS:
            response = client.beta.messages.create(
                betas=["server-side-fallback-2026-07-01"], fallbacks="default", **request)
        else:
            response = client.messages.create(**request)
    except anthropic.AuthenticationError:
        raise RetakeError("Claude API のキーが無効です（.env の ANTHROPIC_API_KEY）。")
    except anthropic.NotFoundError:
        raise RetakeError("モデル「%s」が見つかりません（.env の %s）。" % (model, MODEL_ENV))
    except anthropic.APIStatusError as exc:
        raise RetakeError("Claude API がエラーを返しました（%s）。" % exc.status_code)
    except anthropic.APIConnectionError:
        raise RetakeError("Claude API に接続できませんでした。")

    if response.stop_reason == "refusal":
        raise RetakeError("Claude API が判定を断りました。")
    text = "".join(b.text for b in response.content if b.type == "text")
    try:
        answer = json.loads(text)
    except ValueError:
        raise RetakeError("Claude API の答えを読めませんでした。")
    return ranges_from_answer(words, answer)


def describe(words, ranges):
    """何を消したかを人に見せる形にする（書き出しの報告用）。"""
    out = []
    for lo, hi in ranges:
        out.append(join_words(w.text for w in words if w.start >= lo - 1e-6 and w.end <= hi + 1e-6))
    return out


def main():
    """API に実際に通るかを、短い例文で確かめる。

        python -m stream_highlight.retake
    """
    import sys

    from .transcribe import Word

    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass
    sample = ("えー 今日は ホラゲ あ 今日は ホラーゲームを やって いきます "
              "最後まで 見て いって ください ね それでは 始めて いきましょう").split()
    words = [Word(text, i * 0.5, i * 0.5 + 0.4) for i, text in enumerate(sample)]
    print("モデル: %s" % model_name())
    print("入力  : %s" % join_words(sample))
    try:
        ranges = detect(words)
    except RetakeError as exc:
        print("失敗  : %s" % exc)
        raise SystemExit(1)
    print("削除  : %s" % (" / ".join(describe(words, ranges)) or "（なし）"))
    print("API への接続は成功しました。")


if __name__ == "__main__":
    main()
