# -*- coding: utf-8 -*-
"""書き出す区間から「喋っていない間」を詰める。

配信の音声はマイクとゲーム音が1本に混ざっているので、音量（silencedetect）
では喋っていない所を見分けられない。文字起こしの単語の時刻を発話区間とし、
発話と発話の間が長い所だけを詰める。

ただしチャットのピークの前後は詰めない。喋っていなくても、ゲーム画面で
何かが起きて視聴者が反応している場面がありうる。切り抜きではそこが
一番大事な場面であることが多い。

ここで決めるのは「元の配信のどこを残すか」（秒）だけ。フレームへの丸めと
シーケンス上の位置は editexport.Timeline が受け持つ。
"""

from dataclasses import dataclass

DEFAULT_GAP_SEC = 1.5          # これより長い沈黙を詰める
DEFAULT_PAD_SEC = 0.15         # カット点の前後に残す余白（語尾・息継ぎを切らない）
DEFAULT_PROTECT_SEC = 5.0      # ピークの前後この秒数はカットしない
DEFAULT_MIN_PIECE_SEC = 0.5    # これより短い断片は残さず、前後とまとめる
MIN_REMOVE_SEC = 0.1           # これより短いカットは、詰めても意味が無いのでやらない


@dataclass
class CutSettings:
    gap_sec: float = DEFAULT_GAP_SEC
    pad_sec: float = DEFAULT_PAD_SEC
    protect_sec: float = DEFAULT_PROTECT_SEC
    min_piece_sec: float = DEFAULT_MIN_PIECE_SEC

    @classmethod
    def from_values(cls, gap_sec=None, pad_sec=None, protect_sec=None, min_piece_sec=None):
        def pick(value, default, low, high):
            try:
                value = float(value)
            except (TypeError, ValueError):
                return default
            return min(max(value, low), high)
        return cls(gap_sec=pick(gap_sec, DEFAULT_GAP_SEC, 0.3, 30.0),
                   pad_sec=pick(pad_sec, DEFAULT_PAD_SEC, 0.0, 2.0),
                   protect_sec=pick(protect_sec, DEFAULT_PROTECT_SEC, 0.0, 120.0),
                   min_piece_sec=pick(min_piece_sec, DEFAULT_MIN_PIECE_SEC, 0.0, 10.0))


def _union(ranges):
    out = []
    for lo, hi in sorted(r for r in ranges if r[1] > r[0]):
        if out and lo <= out[-1][1]:
            out[-1][1] = max(out[-1][1], hi)
        else:
            out.append([lo, hi])
    return [tuple(r) for r in out]


def _subtract(ranges, holes):
    """ranges から holes を取り除く。どちらも (開始, 終了) の一覧。"""
    out = []
    holes = _union(holes)
    for lo, hi in ranges:
        pieces = [(lo, hi)]
        for h_lo, h_hi in holes:
            nxt = []
            for p_lo, p_hi in pieces:
                if h_hi <= p_lo or h_lo >= p_hi:
                    nxt.append((p_lo, p_hi))
                    continue
                if h_lo > p_lo:
                    nxt.append((p_lo, h_lo))
                if h_hi < p_hi:
                    nxt.append((h_hi, p_hi))
            pieces = nxt
        out.extend(pieces)
    return [r for r in out if r[1] > r[0]]


def silence_ranges(start, end, words, settings):
    """詰めてよい沈黙（元の配信の秒）。余白を残した後の範囲を返す。

    区間の頭と終わりの沈黙も、長ければ詰める。区間の端は発話の端と
    同じ扱いで、余白は付けない（つなぎ目ではないため）。
    単語が1つも無い区間は詰めない。文字起こしの失敗と区別できないので、
    丸ごと消すより残すほうが安全。
    """
    spoken = sorted((max(start, w.start), min(end, w.end)) for w in words
                    if w.end > start and w.start < end)
    if not spoken:
        return []
    pad = settings.pad_sec
    out = []
    # 発話どうしの間。重なる単語があるので、ここまでに喋り終えた最後の時刻を追う
    edge = start
    first = True
    for lo, hi in spoken:
        gap = lo - edge
        if gap > settings.gap_sec:
            left = edge if first else edge + pad
            right = lo - pad
            if right > left:
                out.append((left, right))
        edge = max(edge, hi)
        first = False
    if end - edge > settings.gap_sec:
        out.append((edge + pad, end))
    return out


def plan_cuts(start, end, words, peaks, settings=None, extra_removals=()):
    """区間 start〜end のうち、残す部分を [(開始, 終了)] で返す（元の配信の秒）。

    extra_removals は言い直しなど、沈黙以外に消したい範囲。
    ピークの保護範囲はどのカットよりも優先する。
    """
    settings = settings or CutSettings()
    removals = list(silence_ranges(start, end, words, settings))
    removals.extend((max(start, lo), min(end, hi)) for lo, hi in extra_removals)
    protect = [(p - settings.protect_sec, p + settings.protect_sec) for p in peaks]
    removals = [r for r in _subtract(_union(removals), protect)
                if r[1] - r[0] >= MIN_REMOVE_SEC]
    kept = _subtract([(start, end)], removals)
    return _absorb_short(kept, settings.min_piece_sec)


def _absorb_short(kept, min_piece):
    """短すぎる断片を、隣の断片とつなげて残す（間のカットを取りやめる）。

    0.5秒に満たない断片が一瞬だけ映ると、編集ではただのノイズになる。
    消すと発話を落とすかもしれないので、消すのではなく隣とまとめる。
    つなぐ相手は、間のカットが短いほう（元の流れを崩しにくい）。
    """
    pieces = [list(p) for p in kept]
    while len(pieces) > 1:
        short = [i for i, p in enumerate(pieces) if p[1] - p[0] < min_piece]
        if not short:
            break
        i = short[0]
        left_gap = pieces[i][0] - pieces[i - 1][1] if i > 0 else None
        right_gap = pieces[i + 1][0] - pieces[i][1] if i + 1 < len(pieces) else None
        if right_gap is None or (left_gap is not None and left_gap <= right_gap):
            pieces[i - 1][1] = pieces[i][1]
        else:
            pieces[i + 1][0] = pieces[i][0]
        del pieces[i]
    return [tuple(p) for p in pieces]
