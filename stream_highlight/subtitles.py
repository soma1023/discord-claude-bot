# -*- coding: utf-8 -*-
"""文字起こしからテロップ（SRT）を作る。

Premiere は FCP7 XML に埋め込んだテキストを正しく読めないことが多いので、
テロップは XML とは別の SRT にして、キャプショントラックとして読み込んでもらう。

時刻はカット後のシーケンス上の時刻（元の配信の時刻ではない）。
editexport.Timeline と同じフレームの丸めを使うので、映像とずれない。
"""

import re
from dataclasses import dataclass, field

from .transcribe import Word, join_words

LINE_CHARS = 16          # 1行の最大文字数
MAX_LINES = 2
MIN_SHOW_SEC = 0.7       # これより短い表示は読めない
MAX_SHOW_SEC = 5.0       # これより長く同じテロップを出さない
PAUSE_SEC = 0.6          # これ以上の間があれば別のテロップにする
SENTENCE_END = "。！？!?…"
SOFT_BREAK = "、，,"
BREAK_PAUSE_SEC = 0.25   # 語と語の間がこれ以上あれば「発話の間」として折り返しの候補にする
# 文節の終わりとみなす助詞。語（文節に近い単位）がこれで終わっていれば、その後ろで折りやすい
BREAK_PARTICLES = "はがをにでもへとのねよ"
TRAILING_PUNCT = "、。，．,."
SHORT_LINE_CHARS = 2     # これ以下の文字数だけの行は作らない（前の行と折り直す）


@dataclass
class CaptionStyle:
    """テロップの体裁。チャンネルごとの設定（profiles.py）で変えられる。既定値は従来どおり。"""
    line_chars: int = LINE_CHARS
    max_lines: int = MAX_LINES
    min_show_sec: float = MIN_SHOW_SEC
    max_show_sec: float = MAX_SHOW_SEC
    strip_trailing_punct: bool = False     # 各行の末尾の「、」「。」を消す
    replacements: dict = field(default_factory=dict)    # 表記の置換辞書 {"ねこみや": "猫宮"}

    @classmethod
    def from_values(cls, line_chars=None, max_lines=None, min_show_sec=None, max_show_sec=None):
        """画面や設定ファイルから来た値を、使える範囲に収めて作る。"""
        def pick(value, default, low, high, kind):
            try:
                return min(high, max(low, kind(value)))
            except (TypeError, ValueError):
                return default
        style = cls(pick(line_chars, LINE_CHARS, 6, 40, int),
                    pick(max_lines, MAX_LINES, 1, 4, int),
                    pick(min_show_sec, MIN_SHOW_SEC, 0.2, 5.0, float),
                    pick(max_show_sec, MAX_SHOW_SEC, 1.0, 15.0, float))
        style.max_show_sec = max(style.max_show_sec, style.min_show_sec)
        return style


def _break_score(before, pause):
    """before（語）の後ろで折り返すときの良さ。句読点 ＞ 発話の間 ＞ 助詞の後。"""
    score = 0
    if before[-1] in SOFT_BREAK + SENTENCE_END:
        score += 4
    if pause is not None and pause >= BREAK_PAUSE_SEC:
        score += 2
    if before[-1] in BREAK_PARTICLES:
        score += 1
    return score


def wrap(parts, width=LINE_CHARS, pauses=None):
    """単語の並びを、1行 width 文字以内の行に分ける。

    行は単語（文節に近い単位）の切れ目で折り、単語の途中では切らない。
    溢れたときは、行の後半（width の半分以上）にある切れ目のうち、
    句読点の後 ＞ 発話の間 ＞ 助詞の後 の順に良い所で折る。同じ良さなら行を長く取る。
    折った結果、1〜2文字だけの行（「〜ですよ／ね」）ができたら、前の行と折り直す（_rebalance）。
    pauses は各語の直前の間（秒）。無ければ間は見ない。
    1語で width を超えるものだけは途中で切る。
    """
    items = [((part or "").strip(), None if pauses is None else pauses[i])
             for i, part in enumerate(parts)]
    items = [item for item in items if item[0]]
    lines = []             # 行ごとの [(語, 直前の間)]
    current = []
    for part, pause in items:
        texts = [c[0] for c in current]
        if len(join_words(texts + [part])) <= width:
            current.append((part, pause))
            continue
        if current:
            # k 語目の前で折る候補。k == len(current) は「今の語を次の行へ送る」
            best = None
            for k in range(1, len(current) + 1):
                if k < len(current) and len(join_words(texts[:k])) < width // 2:
                    continue
                after = current[k][1] if k < len(current) else pause
                score = (_break_score(texts[k - 1], after), k)
                if best is None or score > best:
                    best = score
            k = best[1]
            lines.append(current[:k])
            current = current[k:]
            if len(join_words([c[0] for c in current] + [part])) <= width:
                current.append((part, pause))
                continue
            if current:
                lines.append(current)
            current = []
        if len(part) > width:
            # 1語で収まらないときだけ途中で切る。「16文字＋1文字」にならないよう、均等に分ける
            count = -(-len(part) // width)
            size = -(-len(part) // count)
            pieces = [part[i:i + size] for i in range(0, len(part), size)]
            lines.extend([(piece, pause if i == 0 else None)] for i, piece in enumerate(pieces[:-1]))
            part, pause = pieces[-1], None
        current = [(part, pause)] if part else []
    if current:
        lines.append(current)
    return [join_words([w for w, _ in line]) for line in _rebalance(lines, width)]


def _rebalance(lines, width):
    """1〜2文字だけの行を、隣の行と合わせて折り直す。

    短い行が後ろ（「〜ですよ／ね」）でも前（「どう／見られるかとか〜」）でも、2行の語をつなげ、
    両方が SHORT_LINE_CHARS より長く width 以内に収まる切れ目のうち、折り返しの良さ
    （_break_score）が一番高い所で折る。そういう切れ目が無く、1行に収まるならつなげる。
    どちらもできなければそのまま。長さは行末の句読点を除いて数える（消すことがあるため）。
    """
    def text(line):
        return join_words([w for w, _ in line])

    def short(line):
        return len(_strip_line(text(line))) <= SHORT_LINE_CHARS

    out = []
    for line in lines:
        if out and (short(line) or short(out[-1])):
            merged = out[-1] + line
            best = None
            for k in range(1, len(merged)):
                head, tail = merged[:k], merged[k:]
                if (not short(head) and not short(tail)
                        and len(text(head)) <= width and len(text(tail)) <= width):
                    score = (_break_score(merged[k - 1][0], merged[k][1]), k)
                    if best is None or score > best:
                        best = score
            if best is not None:
                out[-1:] = [merged[:best[1]], merged[best[1]:]]
                continue
            if len(text(merged)) <= width:
                out[-1] = merged
                continue
        out.append(line)
    return out


def apply_replacements(words, table):
    """置換辞書で表記を直す。語をまたぐ表記（「ねこ」「みや」）は1語にまとめてから直す。"""
    if not table:
        return list(words)
    keys = sorted((k for k in table if k), key=len, reverse=True)
    out = []
    i = 0
    while i < len(words):
        merged = None
        # 2〜4語をつなげたときに、境目をまたいで辞書の語が現れるか
        for span in range(4, 1, -1):
            if i + span > len(words):
                continue
            chunk = words[i:i + span]
            text = "".join(w.text for w in chunk)
            first_len = len(chunk[0].text)
            last_from = len(text) - len(chunk[-1].text)
            for key in keys:
                at = text.find(key)
                if at >= 0 and at < first_len and at + len(key) > last_from:
                    merged = Word(text, chunk[0].start, chunk[-1].end)
                    break
            if merged:
                i += span
                break
        if merged is None:
            merged = Word(words[i].text, words[i].start, words[i].end)
            i += 1
        out.append(merged)
    for w in out:
        for key in keys:
            w.text = w.text.replace(key, table[key])
    return out


def _strip_line(line):
    return re.sub("[%s]+$" % re.escape(TRAILING_PUNCT), "", line)


def make_cues(words, timeline, style=None):
    """[{"start", "end"（シーケンス上の秒）, "lines"}] を返す。

    カットされた所の語は出さない。発話がカット点をまたぐときは、
    そこでテロップを分ける（つなぎ目の前後で同じテロップが続いて見えないように）。
    """
    style = style or CaptionStyle()
    fps = timeline.fps
    words = apply_replacements(words, style.replacements)
    placed = []        # (クリップ番号, 開始フレーム, 終了フレーム, 語)
    for w in words:
        clip = timeline.locate((w.start + w.end) / 2.0)
        if clip is None:
            continue
        start = timeline.to_sequence(w.start, clip)
        end = timeline.to_sequence(w.end, clip)
        placed.append((clip, start, max(start, end), w.text))

    groups = []
    for item in placed:
        clip, start, end, text = item
        if groups:
            group = groups[-1]
            last = group[-1]
            fits = len(wrap([g[3] for g in group] + [text], style.line_chars,
                            _pauses(group + [item], fps))) <= style.max_lines
            joined = (clip == last[0]
                      and (start - last[2]) / fps < PAUSE_SEC
                      and (end - group[0][1]) / fps <= style.max_show_sec
                      and last[3][-1] not in SENTENCE_END
                      and fits)
            if joined:
                group.append(item)
                continue
        groups.append([item])

    cues = []
    for group in groups:
        start = group[0][1] / fps
        end = group[-1][2] / fps
        lines = wrap([g[3] for g in group], style.line_chars, _pauses(group, fps))
        if style.strip_trailing_punct:
            lines = [l for l in (_strip_line(l) for l in lines) if l]
        cues.append({"start": start, "end": end, "lines": lines})

    # 表示時間を整える。短いものは延ばすが、次のテロップには重ねない
    for index, cue in enumerate(cues):
        limit = cues[index + 1]["start"] if index + 1 < len(cues) else None
        end = max(cue["end"], cue["start"] + style.min_show_sec)
        end = min(end, cue["start"] + style.max_show_sec)
        if limit is not None:
            end = min(end, limit)
        cue["end"] = max(end, cue["start"])
    return [c for c in cues if c["end"] > c["start"] and c["lines"]]


def _pauses(group, fps):
    """テロップにまとめた語の、それぞれ直前の間（秒）。先頭は間なし。"""
    return [0.0] + [(group[k][1] - group[k - 1][2]) / fps for k in range(1, len(group))]


def _srt_time(seconds):
    ms = int(round(max(0.0, seconds) * 1000))
    h, ms = divmod(ms, 3600000)
    m, ms = divmod(ms, 60000)
    s, ms = divmod(ms, 1000)
    return "%02d:%02d:%02d,%03d" % (h, m, s, ms)


def to_srt(cues):
    blocks = []
    for index, cue in enumerate(cues, 1):
        blocks.append("%d\n%s --> %s\n%s\n" % (index, _srt_time(cue["start"]),
                                               _srt_time(cue["end"]), "\n".join(cue["lines"])))
    return "\n".join(blocks)
