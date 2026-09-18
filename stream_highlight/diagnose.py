# -*- coding: utf-8 -*-
"""取得がうまくいかないときに、どこで詰まっているかを調べる。

    python -m stream_highlight.diagnose <配信のURL>

開発側からは YouTube / Twitch に接続できないため、
実際に何が返ってきているかはこのコマンドの出力が頼りになる。
出力をそのまま貼れば原因を特定できる形にしてある。
"""

import json
import sys

from . import cache, sources
from .analyze import AUDIO_BASELINE_Q, Params, _quantile, mad, rolling_quantile
from .audio import SILENCE_FLOOR
from .sources import FetchError, parse_url


def line(text=""):
    print(text, flush=True)


def head(title):
    line()
    line("=" * 60)
    line(title)
    line("=" * 60)


def short(value, limit=1200):
    text = json.dumps(value, ensure_ascii=False)[:limit]
    return text


def check_metadata(url):
    head("[1] 配信の情報を取得（yt-dlp）")
    try:
        info = sources.fetch_info(url)
    except FetchError as exc:
        line("失敗: %s" % exc)
        return None
    line("タイトル : %s" % (info.title or "(取得できず)"))
    line("配信者   : %s" % (info.channel or "(取得できず)"))
    line("長さ     : %.0f 秒 (%.1f 時間)" % (info.duration, info.duration / 3600))
    if not info.duration:
        line("注意: 長さが取れていません。区間分割ができないため取得が遅くなります。")
    return info


def check_twitch(video_id):
    head("[2] Twitch GQL に単発リクエスト")
    working = None
    last_payload = None

    for client_id in sources._GQL_CLIENT_IDS:
        line()
        line("--- Client-ID: %s ---" % client_id)
        try:
            payload = sources._gql_post(
                [sources._comment_request(video_id, offset=0)],
                retries=1, client_id=client_id,
            )
        except FetchError as exc:
            line("通信失敗: %s" % exc)
            continue

        last_payload = payload
        if isinstance(payload, list) and payload and isinstance(payload[0], dict):
            first = payload[0]
            line("応答のキー   : %s" % ", ".join(sorted(first.keys())))
            if first.get("errors"):
                line("errors       : %s" % short(first["errors"], 400))
            data = first.get("data")
            if isinstance(data, dict):
                video = data.get("video")
                line("data.video   : %s" % ("あり" if isinstance(video, dict)
                                            else repr(video)))
                if isinstance(video, dict):
                    comments = video.get("comments")
                    if isinstance(comments, dict):
                        edges = comments.get("edges") or []
                        line("comments     : 取得成功（%d件）" % len(edges))
                        if edges:
                            node = edges[0].get("node") or {}
                            msg = sources._node_to_message(node)
                            line("最初のコメント: [%.0f秒] %s: %s"
                                 % (msg.offset, msg.author, msg.text[:40]))
                        working = working or client_id
                    else:
                        line("comments     : %r  ← ここが問題" % comments)
        else:
            line("想定外の応答形式: %s" % short(payload, 300))

    head("[3] 判定")
    if working:
        line("成功: Client-ID %s… でチャットを取得できます。" % working[:10])
        line("このIDが本番の取得でも使われます。")
        return True

    line("失敗: どのClient-IDでもチャットを取得できませんでした。")
    line()
    if last_payload is not None:
        saved = sources._save_debug(last_payload)
        if saved:
            line("生の応答を保存しました: %s" % saved)
            line("このファイルの中身を貼ってもらえれば原因を特定できます。")
        line()
        line("--- 応答の先頭 ---")
        line(short(last_payload, 1500))
    line()
    line("考えられること:")
    line(" - サブスク限定・削除済み・非公開のVOD")
    line(" - TwitchがAPIの仕様を変更した（この場合はコード側の対応が必要）")
    line(" - 一時的な混雑やアクセス制限（時間をおくと通ることがある）")
    return False


def check_youtube(url):
    head("[2] YouTube のチャットリプレイを取得")
    line("yt-dlp でチャットを取りに行きます。件数が増えれば成功です。")
    line()
    try:
        messages = sources.fetch_youtube_chat(
            url, progress=lambda message, frac: line("  %s" % message))
    except FetchError as exc:
        head("[3] 判定")
        line("失敗: %s" % exc)
        return False
    head("[3] 判定")
    line("成功: %d 件のコメントを取得しました。" % len(messages))
    if messages:
        first, last = messages[0], messages[-1]
        line("最初: [%.0f秒] %s: %s" % (first.offset, first.author, first.text[:40]))
        line("最後: [%.0f秒] %s: %s" % (last.offset, last.author, last.text[:40]))
    return True


def check_audio(platform, video_id):
    """保存済みの音量データの分布を見る。

    音量の平常値をどう取るべきかは配信の音づくりで変わるので、
    実際のデータの形が分からないと調整できない。
    """
    head("[4] 音量データの分布")
    loudness = cache.load_audio(platform, video_id)
    if loudness is None:
        line("音量データはまだ保存されていません。")
        line("「音声も解析」にチェックを入れて解析すると作られます。")
        return

    values = loudness.values
    ordered = sorted(values)
    line("サンプル数: %d（%.1f 時間ぶん）" % (len(values), loudness.duration / 3600))
    line()
    line("音量の分布（LUFS）:")
    for q in (0.05, 0.25, 0.50, 0.75, 0.90, 0.99):
        line("  %3.0f%%点 : %7.1f" % (q * 100, _quantile(ordered, q)))
    silent = sum(1 for v in values if v <= SILENCE_FLOOR + 1)
    line("  無音に近い割合: %.1f%%" % (100.0 * silent / len(values)))

    params = Params()
    n_bins = int(loudness.duration // params.bin_sec) + 1
    level = loudness.bin_max(params.bin_sec, n_bins)
    half = max(1, int(round(params.baseline_sec / params.bin_sec)))

    line()
    line("平常値の取り方による「跳ね」の出方:")
    for q, label in ((0.5, "中央値"), (AUDIO_BASELINE_Q, "%.1f分位（現行）" % AUDIO_BASELINE_Q)):
        base = rolling_quantile(level, half, q=q)
        excess = [lv - bs for lv, bs in zip(level, base)]
        spread = max(mad(excess), 1.0)
        calm = sorted(excess)
        over = sum(1 for i, e in enumerate(excess)
                   if e / spread >= params.audio_min_z and e >= params.audio_min_db)
        line("  %-16s ばらつき %4.1f dB / 50%%点 %5.1f dB / 90%%点 %5.1f dB / 検出 %d箇所"
             % (label, spread, _quantile(calm, 0.5), _quantile(calm, 0.9), over))

    base = rolling_quantile(level, half, q=AUDIO_BASELINE_Q)
    excess = [lv - bs for lv, bs in zip(level, base)]
    ranked = sorted(range(len(excess)), key=lambda i: excess[i], reverse=True)[:10]
    line()
    line("音量が上がった箇所 上位10件（現行の基準）:")
    for i in sorted(ranked):
        seconds = i * params.bin_sec
        line("  %2d:%02d:%02d  +%.1f dB（音量 %.1f LUFS）"
             % (seconds // 3600, seconds % 3600 // 60, seconds % 60,
                excess[i], level[i]))


def main():
    if len(sys.argv) < 2:
        line("使い方: python -m stream_highlight.diagnose <配信のURL>")
        return 2

    url = sys.argv[1].strip()
    head("配信ハイライト抽出ツール 診断")
    line("URL: %s" % url)

    try:
        platform, video_id = parse_url(url)
    except FetchError as exc:
        line("失敗: %s" % exc)
        return 1
    line("プラットフォーム: %s / ID: %s" % (platform, video_id))

    info = check_metadata(url)
    if "--audio" in sys.argv:
        # 音量だけ見たいときは取得の確認を飛ばす
        check_audio(platform, video_id)
        head("おわり")
        line("この出力をそのまま貼ってください。")
        return 0

    ok = check_twitch(video_id) if platform == "twitch" else check_youtube(url)
    check_audio(platform, video_id)

    head("おわり")
    line("この出力をそのまま貼ってください。")
    return 0 if ok and info else 1


if __name__ == "__main__":
    sys.exit(main())
