# -*- coding: utf-8 -*-
"""文字起こしからテロップ（SRT）を作る。

Premiere は FCP7 XML に埋め込んだテキストを正しく読めないことが多いので、
テロップは XML とは別の SRT にして、キャプショントラックとして読み込んでもらう。

時刻はカット後のシーケンス上の時刻（元の配信の時刻ではない）。
editexport.Timeline と同じフレームの丸めを使うので、映像とずれない。
"""

from .transcribe import join_words

LINE_CHARS = 16          # 1行の最大文字数
MAX_LINES = 2
MIN_SHOW_SEC = 0.7       # これより短い表示は読めない
MAX_SHOW_SEC = 5.0       # これより長く同じテロップを出さない
PAUSE_SEC = 0.6          # これ以上の間があれば別のテロップにする
SENTENCE_END = "。！？!?…"
SOFT_BREAK = "、，,"


def wrap(parts, width=LINE_CHARS):
    """単語の並びを、1行 width 文字以内の行に分ける。

    行は単語（文節に近い単位）の切れ目で折る。句読点の直後で折れる位置が
    行の後半にあれば、そちらを優先する。1語で width を超えるものは途中で切る。
    """
    lines = []
    current = []
    for part in parts:
        part = (part or "").strip()
        if not part:
            continue
        candidate = join_words(current + [part])
        if len(candidate) <= width:
            current.append(part)
            continue
        if current:
            # 句読点で終わる語の所まで戻せるなら、そこで折る
            cut = None
            for k in range(len(current) - 1, 0, -1):
                if current[k - 1][-1] in SOFT_BREAK + SENTENCE_END and \
                        len(join_words(current[:k])) >= width // 2:
                    cut = k
                    break
            if cut:
                lines.append(join_words(current[:cut]))
                current = current[cut:]
                if len(join_words(current + [part])) <= width:
                    current.append(part)
                    continue
            lines.append(join_words(current))
            current = []
        while len(part) > width:
            lines.append(part[:width])
            part = part[width:]
        current = [part] if part else []
    if current:
        lines.append(join_words(current))
    return lines


def make_cues(words, timeline):
    """[{"start", "end"（シーケンス上の秒）, "lines"}] を返す。

    カットされた所の語は出さない。発話がカット点をまたぐときは、
    そこでテロップを分ける（つなぎ目の前後で同じテロップが続いて見えないように）。
    """
    fps = timeline.fps
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
            fits = len(wrap([g[3] for g in group] + [text])) <= MAX_LINES
            joined = (clip == last[0]
                      and (start - last[2]) / fps < PAUSE_SEC
                      and (end - group[0][1]) / fps <= MAX_SHOW_SEC
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
        cues.append({"start": start, "end": end,
                     "lines": wrap([g[3] for g in group])})

    # 表示時間を整える。短いものは延ばすが、次のテロップには重ねない
    for index, cue in enumerate(cues):
        limit = cues[index + 1]["start"] if index + 1 < len(cues) else None
        end = max(cue["end"], cue["start"] + MIN_SHOW_SEC)
        end = min(end, cue["start"] + MAX_SHOW_SEC)
        if limit is not None:
            end = min(end, limit)
        cue["end"] = max(end, cue["start"])
    return [c for c in cues if c["end"] > c["start"] and c["lines"]]


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
