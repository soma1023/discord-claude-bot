# -*- coding: utf-8 -*-
"""チャットの盛り上がりを数値化して、見せ場候補を抽出する。

考え方:
  1. コメントを一定間隔（既定5秒）のビンに入れる
  2. 前後30秒の移動合計を取って「その瞬間の勢い」にする
  3. 前後10分の中央値を「その時間帯の平常運転」とする（視聴者数の増減に追従させる）
  4. z = (勢い - 平常) / sqrt(平常 + 1) で「普段よりどれだけ跳ねたか」を測る
     コメント数はポアソン的なばらつきをするので、平常値の平方根で割ると
     チャットが静かな時間帯と賑やかな時間帯を同じ物差しで比べられる
  5. zのピークを拾い、近いものをまとめて上位から並べる

同じ仕組みを「草だけ」「特定ワードだけ」の系列にも使えるようにしてある。
"""

import math
from collections import Counter
from dataclasses import dataclass, asdict

from .patterns import (CATEGORIES, CATEGORY_IDS, CATEGORY_LABELS, classify,
                       normalize, split_notice)


@dataclass
class Params:
    bin_sec: int = 5           # 集計の刻み
    window_sec: int = 30       # 勢いを測る窓
    baseline_sec: int = 600    # 平常運転を測る窓（前後それぞれ）
    min_z: float = 3.0         # 見せ場と判定するしきい値（感度）
    merge_sec: int = 90        # これ以内のピークは1つにまとめる
    top_n: int = 30            # 出す候補の数
    skip_start_sec: int = 0    # 配信開始から何秒を無視するか（挨拶ラッシュ対策）
    lead_sec: int = 30         # 切り出し開始をピークの何秒前にするか
    tail_sec: int = 15         # 切り出し終了をピークの何秒後にするか


    def as_dict(self):
        return asdict(self)

    @classmethod
    def from_dict(cls, data):
        """UIから来た値を安全な範囲に丸めて取り込む。"""
        data = data or {}
        limits = {
            "bin_sec": (1, 60), "window_sec": (5, 300), "baseline_sec": (60, 3600),
            "min_z": (0.5, 10.0), "merge_sec": (5, 600), "top_n": (1, 200),
            "lead_sec": (0, 300), "tail_sec": (0, 300),
            "skip_start_sec": (0, 1800),
        }
        floats = {"min_z"}
        kwargs = {}
        for field_name, (low, high) in limits.items():
            if field_name not in data or data[field_name] is None:
                continue
            try:
                value = float(data[field_name])
            except (TypeError, ValueError):
                continue
            value = max(low, min(high, value))
            kwargs[field_name] = value if field_name in floats else int(value)
        params = cls(**kwargs)
        if params.window_sec < params.bin_sec:
            params.window_sec = params.bin_sec
        return params


# 見せ場と呼ぶのに最低限必要なコメント数（窓あたり）。
# これ未満は統計的に跳ねて見えても、切り抜く価値のある反応とは言えない。
MIN_MOMENT_MESSAGES = 5.0




# ------------------------------------------------------------ 数値ユーティリティ

def _prefix_sums(values):
    out = [0.0] * (len(values) + 1)
    total = 0.0
    for i, v in enumerate(values):
        total += v
        out[i + 1] = total
    return out


def rolling_sum(values, half):
    """各点を中心に ±half ビンの合計を取る（端は存在する範囲だけ）。"""
    prefix = _prefix_sums(values)
    n = len(values)
    out = [0.0] * n
    for i in range(n):
        lo = max(0, i - half)
        hi = min(n, i + half + 1)
        out[i] = prefix[hi] - prefix[lo]
    return out


def _median(sorted_values):
    n = len(sorted_values)
    if n == 0:
        return 0.0
    mid = n // 2
    if n % 2:
        return float(sorted_values[mid])
    return (sorted_values[mid - 1] + sorted_values[mid]) / 2.0


def _quantile(sorted_values, q):
    """分位点。q=0.5 なら中央値。"""
    n = len(sorted_values)
    if n == 0:
        return 0.0
    if q <= 0:
        return float(sorted_values[0])
    if q >= 1:
        return float(sorted_values[-1])
    pos = q * (n - 1)
    low = int(pos)
    high = min(n - 1, low + 1)
    frac = pos - low
    return sorted_values[low] * (1 - frac) + sorted_values[high] * frac


def rolling_quantile(values, half, q=0.5, stride=None):
    """移動分位点。全点で厳密に計算すると重いので、間引いて線形補間する。

    平均ではなく分位点を使うのは、盛り上がり自体に平常値を引っ張られないため。
    チャットは中央値(q=0.5)、音量は「普段の大きめの音」を基準にしたいので
    高めの分位点を使う（無音区間に基準を引き下げられないため）。
    """
    n = len(values)
    if n == 0:
        return []
    if half <= 0:
        return list(values)
    if stride is None:
        stride = max(1, half // 10)

    anchors = list(range(0, n, stride))
    if anchors[-1] != n - 1:
        anchors.append(n - 1)

    sampled = []
    for i in anchors:
        lo = max(0, i - half)
        hi = min(n, i + half + 1)
        sampled.append(_quantile(sorted(values[lo:hi]), q))

    out = [0.0] * n
    for idx in range(len(anchors) - 1):
        a, b = anchors[idx], anchors[idx + 1]
        va, vb = sampled[idx], sampled[idx + 1]
        span = b - a
        for i in range(a, b):
            out[i] = va + (vb - va) * ((i - a) / span)
    out[n - 1] = sampled[-1]
    return out


def rolling_median(values, half, stride=None):
    return rolling_quantile(values, half, q=0.5, stride=stride)



def zscores(density, baseline):
    """平常値からの跳ね具合。ポアソン分布を仮定して sqrt(平常) で正規化する。"""
    return [
        (d - b) / math.sqrt(b + 1.0)
        for d, b in zip(density, baseline)
    ]


# ------------------------------------------------------------ 時系列

class Timeline:
    """メッセージをビンに詰めて、系列計算をまとめて担当する。"""

    def __init__(self, messages, duration, params):
        self.params = params
        self.messages = messages
        self.bin_sec = params.bin_sec
        span = max(duration or 0.0, messages[-1].offset if messages else 0.0)
        self.duration = max(span, self.bin_sec)
        self.n = int(self.duration // self.bin_sec) + 1
        self.times = [i * self.bin_sec for i in range(self.n)]

        # 各ビンに属するメッセージの添字。
        # システム通知（Twitchのサブスク告知など）は本人の発言ではないので、
        # 盛り上がりの判定にも代表コメントにも使わない。
        self.buckets = [[] for _ in range(self.n)]
        self.skipped_system = 0
        # 通知の判定はここで行う。取得時にやると、保存済みチャットに
        # 後からルールを足しても効かなくなるため。
        self.body = []
        for msg in messages:
            text, is_notice = split_notice(msg.text)
            self.body.append(text if is_notice else msg.text)

        for idx, msg in enumerate(messages):
            if msg.offset < params.skip_start_sec:
                # 配信開始前の待機チャットと、開始直後の挨拶ラッシュを外す。
                # 既定は0秒なので、指定しない限り何も隠さない。
                continue
            if msg.kind == "system" or not self.body[idx].strip():
                self.skipped_system += 1
                continue
            b = int(msg.offset // self.bin_sec)
            if 0 <= b < self.n:
                self.buckets[b].append(idx)

        # コメント正規化とカテゴリ判定は1回だけ
        self.normalized = [normalize(t) for t in self.body]
        self.categories = [classify(t) for t in self.normalized]

        self.half_window = max(0, int(round(params.window_sec / self.bin_sec)) // 2)
        self.half_baseline = max(1, int(round(params.baseline_sec / self.bin_sec)))
        self.window_sec_actual = (self.half_window * 2 + 1) * self.bin_sec

    # -- 系列づくり -------------------------------------------------

    def counts(self, predicate=None):
        out = [0.0] * self.n
        for b, idxs in enumerate(self.buckets):
            if predicate is None:
                out[b] = float(len(idxs))
            else:
                out[b] = float(sum(1 for i in idxs if predicate(i)))
        return out

    def category_counts(self, category_id):
        return self.counts(lambda i: category_id in self.categories[i])

    def keyword_counts(self, keyword):
        needle = normalize(keyword)
        if not needle:
            return [0.0] * self.n
        return self.counts(lambda i: needle in self.normalized[i])

    def series(self, counts):
        """勢い・平常値・zスコアをまとめて返す。"""
        density = rolling_sum(counts, self.half_window)
        baseline = rolling_median(density, self.half_baseline)
        return density, baseline, zscores(density, baseline)

    def per_minute(self, density):
        factor = 60.0 / self.window_sec_actual
        return [d * factor for d in density]

    # -- ピーク検出 -------------------------------------------------

    def detect(self, counts, top_n=None, min_z=None, min_count=None):
        """勢いの系列からピークを拾って、重ならないように上位を選ぶ。"""
        p = self.params
        top_n = p.top_n if top_n is None else top_n
        min_z = p.min_z if min_z is None else min_z
        density, baseline, z = self.series(counts)

        if min_count is None:
            # 過疎チャットで数コメントが1位になるのを防ぐ絶対的な下限。
            # zスコアは「相対的な跳ね」しか見ないので、ここで実数の足切りをする。
            mean_density = sum(density) / len(density) if density else 0.0
            min_count = max(MIN_MOMENT_MESSAGES, mean_density * 0.5)

        merge_bins = max(1, int(round(p.merge_sec / self.bin_sec)))
        candidates = [
            i for i in range(self.n)
            if z[i] >= min_z and density[i] >= min_count
        ]
        candidates.sort(key=lambda i: z[i], reverse=True)

        picked = []
        for i in candidates:
            if any(abs(i - j) < merge_bins for j in picked):
                continue
            picked.append(i)
            if len(picked) >= top_n:
                break
        picked.sort(key=lambda i: z[i], reverse=True)
        return picked, density, baseline, z

    def top_bins(self, counts, top_n=20):
        """しきい値を使わず、単純に密度の高い順に重ならないビンを拾う。

        キーワード検索など「山でなくても出現箇所を知りたい」ときに使う。
        """
        density = rolling_sum(counts, self.half_window)
        merge_bins = max(1, int(round(self.params.merge_sec / self.bin_sec)))
        order = sorted((i for i in range(self.n) if density[i] > 0),
                       key=lambda i: density[i], reverse=True)
        picked = []
        for i in order:
            if any(abs(i - j) < merge_bins for j in picked):
                continue
            picked.append(i)
            if len(picked) >= top_n:
                break
        return picked

    def _span(self, peak, z, min_z):
        """ピークの前後で、盛り上がりが続いている範囲を求める。"""
        floor = max(1.0, min_z * 0.4)
        limit = max(1, int(round(180 / self.bin_sec)))
        start = peak
        while start > 0 and peak - start < limit and z[start - 1] >= floor:
            start -= 1
        end = peak
        while end < self.n - 1 and end - peak < limit and z[end + 1] >= floor:
            end += 1
        return start, end

    def messages_between(self, start_sec, end_sec):
        lo = max(0, int(start_sec // self.bin_sec))
        hi = min(self.n - 1, int(end_sec // self.bin_sec))
        out = []
        for b in range(lo, hi + 1):
            for i in self.buckets[b]:
                if start_sec <= self.messages[i].offset <= end_sec:
                    out.append(i)
        return out


# ------------------------------------------------------------ 見せ場の組み立て

def _top_comments(timeline, idxs, limit=5):
    """よく飛んでいたコメントを、表記を代表させて集計する。"""
    groups = {}
    for i in idxs:
        key = timeline.normalized[i]
        if not key:
            continue
        slot = groups.setdefault(key, {"count": 0, "forms": Counter()})
        slot["count"] += 1
        slot["forms"][timeline.body[i]] += 1
    ranked = sorted(groups.values(), key=lambda g: g["count"], reverse=True)
    return [
        {"text": g["forms"].most_common(1)[0][0], "count": g["count"]}
        for g in ranked[:limit]
    ]


def _build_moment(timeline, info, peak, z, density, baseline, min_z):
    p = timeline.params
    bin_sec = timeline.bin_sec
    peak_sec = peak * bin_sec + bin_sec / 2.0
    start_bin, end_bin = timeline._span(peak, z, min_z)

    clip_start = max(0.0, peak_sec - p.lead_sec)
    clip_end = min(timeline.duration, peak_sec + p.tail_sec)
    idxs = timeline.messages_between(clip_start, clip_end)

    cat_counts = Counter()
    for i in idxs:
        for cid in timeline.categories[i]:
            cat_counts[cid] += 1
    authors = {timeline.messages[i].author for i in idxs if timeline.messages[i].author}
    supers = [i for i in idxs if timeline.messages[i].kind in ("paid", "sticker")]

    base = max(baseline[peak], 0.5)
    rate = density[peak] * 60.0 / timeline.window_sec_actual

    tags = [
        {"id": cid, "label": CATEGORY_LABELS[cid], "count": n}
        for cid, n in cat_counts.most_common(3) if n >= 3
    ]

    return {
        "peak_sec": round(peak_sec, 1),
        "span_start": round(start_bin * bin_sec, 1),
        "span_end": round((end_bin + 1) * bin_sec, 1),
        "clip_start": round(clip_start, 1),
        "clip_end": round(clip_end, 1),
        "score": round(z[peak], 2),
        "ratio": round(density[peak] / base, 2),
        "rate": round(rate, 1),
        "baseline_rate": round(baseline[peak] * 60.0 / timeline.window_sec_actual, 1),
        "messages": len(idxs),
        "authors": len(authors),
        "superchats": len(supers),
        "tags": tags,
        "top_comments": _top_comments(timeline, idxs),
        "samples": [
            {
                "t": round(timeline.messages[i].offset, 1),
                "author": timeline.messages[i].author,
                "text": timeline.body[i],
            }
            for i in idxs[:: max(1, len(idxs) // 8)][:8]
        ],
        "url": info.time_url(clip_start),
    }


def _downsample(values, step):
    if step <= 1:
        return [round(v, 2) for v in values]
    out = []
    for i in range(0, len(values), step):
        chunk = values[i:i + step]
        out.append(round(max(chunk), 2))
    return out


def analyze(messages, info, params=None, max_points=1800):
    """解析のメイン。UIにそのまま渡せる辞書を返す。"""
    params = params or Params()
    timeline = Timeline(messages, info.duration, params)

    counts = timeline.counts()
    picked, density, baseline, z = timeline.detect(counts)

    moments = []
    for rank, peak in enumerate(picked, 1):
        moment = _build_moment(timeline, info, peak, z, density, baseline, params.min_z)
        moment["rank"] = rank
        moments.append(moment)

    # カテゴリ別のランキング（「wwwが多かったところ」など）
    category_rankings = {}
    category_series = {}
    step = max(1, timeline.n // max_points)
    for cat in CATEGORIES:
        cid = cat["id"]
        cat_counts = timeline.category_counts(cid)
        if sum(cat_counts) < 10:
            continue
        cat_picked, cat_density, cat_baseline, cat_z = timeline.detect(
            cat_counts, top_n=10
        )
        category_series[cid] = _downsample(timeline.per_minute(cat_density), step)
        category_rankings[cid] = [
            dict(_build_moment(timeline, info, peak, cat_z, cat_density, cat_baseline,
                               params.min_z), rank=i)
            for i, peak in enumerate(cat_picked, 1)
        ]

    counted = [m for i, m in enumerate(messages)
               if m.offset >= 0 and m.kind != "system" and timeline.body[i].strip()]
    total = len(counted)
    stats = {
        "messages": total,
        "authors": len({m.author for m in counted if m.author}),
        "duration": round(timeline.duration, 1),
        "per_minute": round(total / (timeline.duration / 60.0), 1) if timeline.duration else 0,
        "superchats": len([m for m in messages if m.kind in ("paid", "sticker")]),
        "system_notices": timeline.skipped_system,
    }

    notice = ""
    if total == 0:
        notice = "チャットのコメントが取得できませんでした。"
    elif not moments:
        notice = (
            "しきい値を超える盛り上がりが見つかりませんでした。"
            "感度を下げるか、コメント数の多い配信で試してください。"
            if total >= 200 else
            "コメントが%d件と少ないため、盛り上がりを判定できませんでした。" % total
        )

    return {
        "video": info.as_dict(),
        "stats": stats,
        "notice": notice,
        "params": params.as_dict(),
        "series": {
            "bin_sec": timeline.bin_sec * step,
            "start": 0,
            "rate": _downsample(timeline.per_minute(density), step),
            "baseline": _downsample(timeline.per_minute(baseline), step),
            "categories": category_series,
        },
        "categories": [
            {"id": c["id"], "label": c["label"], "color": c["color"]}
            for c in CATEGORIES if c["id"] in category_series
        ],
        "moments": moments,
        "category_moments": category_rankings,
    }


def analyze_keyword(messages, info, keyword, params=None, max_points=1800):
    """任意のワードに絞って、出現タイミングと山を返す。"""
    params = params or Params()
    timeline = Timeline(messages, info.duration, params)
    counts = timeline.keyword_counts(keyword)
    hits = int(sum(counts))
    if hits == 0:
        return {"keyword": keyword, "hits": 0, "series": [], "bin_sec": timeline.bin_sec,
                "moments": []}

    picked, density, baseline, z = timeline.detect(counts, top_n=20)
    if not picked:
        # 山と呼べるほど跳ねていなくても、出現箇所は知りたいので素朴に拾う
        picked = timeline.top_bins(counts, top_n=20)
    step = max(1, timeline.n // max_points)
    moments = [
        dict(_build_moment(timeline, info, peak, z, density, baseline, params.min_z), rank=i)
        for i, peak in enumerate(picked, 1)
    ]
    return {
        "keyword": keyword,
        "hits": hits,
        "bin_sec": timeline.bin_sec * step,
        "series": _downsample(timeline.per_minute(density), step),
        "moments": moments,
    }
