# 作業状況メモ

最終更新: 2026-09-18 / ブランチ `claude/stream-highlight-extractor-lzeupt`

再開するときは、まずこのファイルを読めば続きから進められる。

## いま何ができているか

配信アーカイブのURLから見せ場候補をランキングするWebアプリ。
`start_highlight.bat` を実行するとローカルサーバが立ち、ブラウザが開く。
Discordボットとは無関係で、`bot.py` は1行も変更していない。

- チャットの盛り上がり検出（時間帯ごとの平常値と比べたスコア）
- 「草」「驚き」などカテゴリ別ランキング、任意ワード検索
- 音声のラウドネス解析（任意・チャットと照合してSEの誤検知を落とす）
- コメント密度グラフ、タイムスタンプ/チャプター/CSV書き出し
- 取得したチャットと音量列はキャッシュし、設定変更時は再取得なし

詳しい仕組みと設定は `README.md` を参照。テストは65件（`python -m unittest discover -s stream_highlight/tests -t .`）。

## 検証できている範囲（重要）

| 対象 | 状態 |
| --- | --- |
| 解析ロジック | 合成データで検証済み。仕込んだ盛り上がりを全件検出 |
| 音声の誤検知フィルタ | 合成データで検証済み。SE誤通過0%（しきい値2.0） |
| 画面・API一式 | ブラウザ自動操作で検証済み |
| ffmpegでの音量抽出 | 実ファイルを通して検証済み |
| **YouTubeからの実取得** | **未検証**（開発環境がYouTubeに接続できないため） |
| **Twitchからの実取得** | **未検証。実機で失敗を確認（下記）** |

開発環境のプロキシが youtube.com と gql.twitch.tv を遮断しているため、
取得部分だけは実機でしか確かめられない。

## 次にやること

1. **TwitchのVODで試す**（利用者のメイン用途）。失敗したら診断コマンドの出力を見る。
2. YouTubeのアーカイブでも試す。まだ一度も成功していない。
   「音声も解析」はオフのまま試すのが早い。

## 未解決の問題

### Twitchが comments を null で返す（対策済み・実機未確認）

実機で `data.video.comments` が null になり取得できなかった。

**最有力の原因**: Client-ID が古かった。
手元の yt-dlp（2026.08.19）のTwitch実装を読んだところ、
`kimne78kx3ncx6brgo4mv6wki5h1ko`（旧Web版）ではなく
`ue6666qo983tsx6so1t0vnawi233wa` を使っており、
Content-Type も `application/json` ではなく `text/plain;charset=UTF-8` だった。
旧IDは整合性チェックで弾かれている可能性が高い。

対策として入れたもの:

- Client-ID を2つ用意し、実際に通るほうを実行時に選ぶ
- ヘッダーをTwitch自身のクライアントに合わせる（Content-Type / Origin / Referer）
- 並列取得の前に単発リクエストで疎通を確認する
- 並列数を8から3に下げ、分割は30分以上のVODのみ

**まだ実機で確認できていない。** 次に失敗したときは診断コマンドを使う:

```
python -m stream_highlight.diagnose <VODのURL>
```

どのClient-IDが何を返したかが出る。失敗時は生の応答が
`twitch_debug.json` に保存されるので、その中身が決め手になる。

## 直したバグ（実機で判明したもの）

- `requirements.txt` の日本語コメント → 日本語WindowsのpipがUnicodeDecodeError
  （pipはロケール既定のcp932で読むため、このファイルに非ASCIIを置いてはいけない）
- 起動バッチが、原因を問わず「Pythonが入っているか確認してください」と誤案内
- Twitchのnull応答で `AttributeError` → 解析全体が失敗
- チャットログの入れ子nullで落ちる箇所を全面的に防御（1行の崩れで全体を落とさない）

## セキュリティ（要確認）

`.env.for-partner` にDiscordボットトークンとAnthropic APIキーが平文で入っていた。
**両方の再発行が済んでいるか未確認。** 未対応なら以下から:

- Discord: https://discord.com/developers/applications → Bot → Reset Token
- Anthropic: https://console.anthropic.com/settings/keys → Revoke して再作成

`.gitignore` は `.env.*` を除外するよう修正済み（`.env.example` は対象外）。

## 手元での動かし方

```
cd /d C:\Users\somas\discord-claude-bot
git pull
start_highlight.bat
```

Python 3.10 / pip は導入済みであることを確認済み。

## 触れるデモ

実データを焼き込んだ公開版（PC不要・URL取得は不可）:
https://claude.ai/artifact/BhpxbMAEAxuDDZrX3Ng3qk

## まだ手を付けていないこと

- Discordボットからの呼び出し（今回は意図的に対象外）
- ライブ配信のリアルタイム監視（アーカイブのみ対応）
- マイク音とゲーム音の分離（重すぎるため見送り）
