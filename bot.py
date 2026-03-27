import discord
import subprocess
import os
import asyncio
import base64
import aiohttp
import anthropic
import openpyxl
import send2trash
from dotenv import load_dotenv

# ローカル .env（相手ユーザーのPC用）を優先して読み込む
load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), '.env'))
# オーナーPC用のパスからも読み込む
load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), '..', '.claude', 'channels', 'discord', '.env'))
load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), '..', 'discord-job-manager', '.env'))

TOKEN = os.getenv('DISCORD_BOT_TOKEN')
ANTHROPIC_API_KEY = os.getenv('ANTHROPIC_API_KEY')

# .envから設定を読み込む（各PCで独自に設定）
ALLOWED_USER_IDS = {
    int(uid.strip())
    for uid in os.getenv('ALLOWED_USER_IDS', '').split(',')
    if uid.strip()
}
WORK_DIR = os.path.expanduser('~')
MAX_HISTORY = 20  # 1ユーザーあたりの最大メッセージ数（往復10回）

SYSTEM_PROMPT = """あなたはClaudeです。ユーザーのDiscordから指示を受け、PCを操作するAIアシスタントです。
ファイルの読み書き、コード実行、コマンド実行など、Claude Codeと同等の作業ができます。
作業内容は日本語で簡潔に報告してください。
OSはWindows 11です。コマンドはbashまたはPowerShell構文で実行できます。"""

TOOLS = [
    {
        "name": "bash",
        "description": "シェルコマンドを実行する。Windows環境なのでbash/PowerShell両方使える。",
        "input_schema": {
            "type": "object",
            "properties": {
                "command": {"type": "string", "description": "実行するコマンド"}
            },
            "required": ["command"]
        }
    },
    {
        "name": "read_file",
        "description": "テキストファイルの内容を読み込む。",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "読み込むファイルのパス"}
            },
            "required": ["path"]
        }
    },
    {
        "name": "write_file",
        "description": "ファイルに内容を書き込む（上書き）。",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "書き込むファイルのパス"},
                "content": {"type": "string", "description": "書き込む内容"}
            },
            "required": ["path", "content"]
        }
    },
    {
        "name": "send_file",
        "description": "PCにあるファイル（画像・テキスト・コードなど）をDiscordチャットに送信する。",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "送信するファイルのパス"},
                "caption": {"type": "string", "description": "ファイルに添えるコメント（省略可）"}
            },
            "required": ["path"]
        }
    },
    {
        "name": "read_image",
        "description": "PC上の画像ファイルを読み込んで内容を解析する。スクリーンショット・図・写真など。",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "解析する画像ファイルのパス（png/jpg/gif/webp）"}
            },
            "required": ["path"]
        }
    },
    {
        "name": "list_files",
        "description": "ディレクトリのファイル一覧を取得する。",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "一覧を取得するディレクトリのパス（省略時はホームディレクトリ）"}
            },
            "required": []
        }
    },
    {
        "name": "search_files",
        "description": "ファイル内のテキストを検索する（grep相当）。",
        "input_schema": {
            "type": "object",
            "properties": {
                "pattern": {"type": "string", "description": "検索パターン（正規表現可）"},
                "path": {"type": "string", "description": "検索対象のパスまたはディレクトリ"}
            },
            "required": ["pattern", "path"]
        }
    },
    {
        "name": "trash",
        "description": "ファイルまたはフォルダをゴミ箱に移動する（完全削除ではなく復元可能）。",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "ゴミ箱に移動するファイル/フォルダのパス"}
            },
            "required": ["path"]
        }
    },
    {
        "name": "excel_read",
        "description": "Excelファイルのセル範囲を読み込む。シート名省略時はアクティブシートを使用。",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "Excelファイルのパス（.xlsx）"},
                "sheet": {"type": "string", "description": "シート名（省略時はアクティブシート）"},
                "range": {"type": "string", "description": "セル範囲（例: A1:C10）。省略時はデータ全体"}
            },
            "required": ["path"]
        }
    },
    {
        "name": "excel_write",
        "description": "Excelファイルの指定セルに値を書き込む。ファイルが存在しない場合は新規作成。",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "Excelファイルのパス（.xlsx）"},
                "sheet": {"type": "string", "description": "シート名（省略時はアクティブシート）"},
                "cell": {"type": "string", "description": "セル番地（例: B3）"},
                "value": {"description": "書き込む値（文字列・数値・日付など）"}
            },
            "required": ["path", "cell", "value"]
        }
    },
    {
        "name": "excel_append",
        "description": "Excelファイルの最終行の次に新しい行を追加する。",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "Excelファイルのパス（.xlsx）"},
                "sheet": {"type": "string", "description": "シート名（省略時はアクティブシート）"},
                "values": {
                    "type": "array",
                    "description": "追加する行のデータ（例: [\"山田\", 25, \"東京\"]）",
                    "items": {}
                }
            },
            "required": ["path", "values"]
        }
    }
]

intents = discord.Intents.default()
intents.message_content = True
client = discord.Client(intents=intents)
ai = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

# ユーザーごとの会話履歴 { user_id: [{role, content}, ...] }
history: dict[int, list] = {}


def get_history(user_id: int) -> list:
    return history.setdefault(user_id, [])


def add_to_history(user_id: int, role: str, content: str):
    h = get_history(user_id)
    h.append({"role": role, "content": content})
    while len(h) > MAX_HISTORY:
        h.pop(0)
        h.pop(0)


def execute_tool(name: str, inp: dict) -> str:
    try:
        if name == "bash":
            result = subprocess.run(
                inp["command"], shell=True, capture_output=True,
                text=True, encoding='utf-8', errors='replace',
                timeout=60, cwd=WORK_DIR
            )
            out = (result.stdout + result.stderr).strip()
            return out[:8000] if out else "（出力なし）"

        elif name == "read_file":
            path = os.path.expanduser(inp["path"])
            with open(path, 'r', encoding='utf-8', errors='replace') as f:
                content = f.read()
            return content[:8000]

        elif name == "write_file":
            path = os.path.expanduser(inp["path"])
            os.makedirs(os.path.dirname(path), exist_ok=True) if os.path.dirname(path) else None
            with open(path, 'w', encoding='utf-8') as f:
                f.write(inp["content"])
            return f"書き込み完了: {path}"

        elif name == "list_files":
            path = os.path.expanduser(inp.get("path", WORK_DIR))
            entries = os.listdir(path)
            lines = []
            for e in sorted(entries):
                full = os.path.join(path, e)
                tag = "/" if os.path.isdir(full) else ""
                lines.append(f"{e}{tag}")
            return "\n".join(lines) or "（空）"

        elif name == "search_files":
            result = subprocess.run(
                ["grep", "-r", "-n", inp["pattern"], os.path.expanduser(inp["path"])],
                capture_output=True, text=True, encoding='utf-8',
                errors='replace', timeout=30
            )
            out = result.stdout.strip()
            return out[:8000] if out else "マッチなし"

        elif name == "trash":
            path = os.path.expanduser(inp["path"])
            send2trash.send2trash(path)
            return f"ゴミ箱に移動しました: {path}"

        elif name == "excel_read":
            path = os.path.expanduser(inp["path"])
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb[inp["sheet"]] if inp.get("sheet") else wb.active
            if inp.get("range"):
                cells = ws[inp["range"]]
                if not isinstance(cells, (tuple, list)):
                    # 単一セル（例: "A1"）
                    rows = [[cells.value]]
                elif cells and not isinstance(cells[0], (tuple, list)):
                    # 単一行（例: "A1:C1"が1行の場合）
                    rows = [[c.value for c in cells]]
                else:
                    rows = [[c.value for c in row] for row in cells]
            else:
                rows = [[c.value for c in row] for row in ws.iter_rows()]
            lines = []
            for row in rows:
                lines.append("\t".join("" if v is None else str(v) for v in row))
            return "\n".join(lines)[:8000] or "（データなし）"

        elif name == "excel_write":
            path = os.path.expanduser(inp["path"])
            if os.path.exists(path):
                wb = openpyxl.load_workbook(path)
            else:
                wb = openpyxl.Workbook()
            ws = wb[inp["sheet"]] if inp.get("sheet") and inp["sheet"] in wb.sheetnames else wb.active
            ws[inp["cell"]] = inp["value"]
            wb.save(path)
            return f"{inp['cell']} に '{inp['value']}' を書き込みました: {path}"

        elif name == "excel_append":
            path = os.path.expanduser(inp["path"])
            if os.path.exists(path):
                wb = openpyxl.load_workbook(path)
            else:
                wb = openpyxl.Workbook()
            ws = wb[inp["sheet"]] if inp.get("sheet") and inp["sheet"] in wb.sheetnames else wb.active
            ws.append(inp["values"])
            wb.save(path)
            return f"行を追加しました（{len(inp['values'])}列）: {path}"

    except subprocess.TimeoutExpired:
        return "タイムアウト（60秒）"
    except Exception as e:
        return f"エラー: {e}"


async def update_status_from_headers(headers):
    """レート制限ヘッダーからDiscordステータスを更新する"""
    try:
        remaining = headers.get('anthropic-ratelimit-tokens-remaining')
        limit = headers.get('anthropic-ratelimit-tokens-limit')
        if remaining and limit:
            pct = int(remaining) / int(limit)
            if pct > 0.6:
                status = discord.Status.online
            elif pct > 0.3:
                status = discord.Status.idle
            else:
                status = discord.Status.dnd
            await client.change_presence(status=status)
    except Exception:
        pass


async def run_agent(user_id: int, user_content, channel) -> str:
    """ツール使用を含むエージェントループ"""
    # 履歴＋今回のメッセージでmessagesを構築
    messages = list(get_history(user_id)) + [{"role": "user", "content": user_content}]

    tool_count = 0
    while True:
        raw = await asyncio.get_event_loop().run_in_executor(
            None,
            lambda msgs=messages: ai.messages.with_raw_response.create(
                model="claude-opus-4-6",
                max_tokens=4096,
                system=SYSTEM_PROMPT,
                tools=TOOLS,
                messages=msgs
            )
        )
        response = raw.parse()
        await update_status_from_headers(raw.headers)

        if response.stop_reason != "tool_use":
            # 最終回答
            final = next(
                (b.text for b in response.content if hasattr(b, 'text')),
                "（応答なし）"
            )
            # 履歴に追加（ユーザーメッセージ＋最終回答のみ）
            user_text = user_content if isinstance(user_content, str) else "[画像+テキスト]"
            add_to_history(user_id, "user", user_text)
            add_to_history(user_id, "assistant", final)
            return final

        # ツール実行
        messages.append({"role": "assistant", "content": response.content})
        tool_results = []

        for block in response.content:
            if block.type != "tool_use":
                continue

            tool_count += 1
            await channel.send(f"🔧 `{block.name}` 実行中... ({tool_count}回目)")

            # send_file: 非同期でDiscordにファイルを送信
            if block.name == "send_file":
                path = os.path.expanduser(block.input["path"])
                caption = block.input.get("caption", "")
                try:
                    await channel.send(
                        content=caption if caption else None,
                        file=discord.File(path)
                    )
                    result_content = f"ファイルを送信しました: {path}"
                except Exception as e:
                    result_content = f"送信エラー: {e}"
                tool_results.append({
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": result_content
                })

            # read_image: 画像をbase64で読み込んでClaudeに渡す
            elif block.name == "read_image":
                path = os.path.expanduser(block.input["path"])
                try:
                    ext = os.path.splitext(path)[1].lower()
                    media_map = {".png": "image/png", ".jpg": "image/jpeg",
                                 ".jpeg": "image/jpeg", ".gif": "image/gif",
                                 ".webp": "image/webp"}
                    media_type = media_map.get(ext, "image/png")
                    with open(path, "rb") as f:
                        img_b64 = base64.standard_b64encode(f.read()).decode()
                    result_content = [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": img_b64
                            }
                        }
                    ]
                except Exception as e:
                    result_content = f"画像読み込みエラー: {e}"
                tool_results.append({
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": result_content
                })

            # その他のツール
            else:
                result = await asyncio.get_event_loop().run_in_executor(
                    None, lambda b=block: execute_tool(b.name, b.input)
                )
                tool_results.append({
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": result
                })

        messages.append({"role": "user", "content": tool_results})

        # 無限ループ防止
        if tool_count >= 20:
            return "ツール実行が20回を超えたため中断しました。"


async def send_long(message, response: str):
    """2000文字超えを分割送信"""
    if len(response) > 1900:
        chunks = [response[i:i+1900] for i in range(0, len(response), 1900)]
        for i, chunk in enumerate(chunks):
            prefix = f"[{i+1}/{len(chunks)}]\n" if len(chunks) > 1 else ""
            await message.reply(f"{prefix}```\n{chunk}\n```")
    else:
        await message.reply(response)


@client.event
async def on_ready():
    print(f"起動完了: {client.user}")
    await client.change_presence(status=discord.Status.online)


@client.event
async def on_message(message):
    if message.author == client.user:
        return
    if message.author.id not in ALLOWED_USER_IDS:
        return
    # DMのみ受け付ける
    if not isinstance(message.channel, discord.DMChannel):
        return

    content = message.content.strip()
    user_id = message.author.id

    # コマンド処理
    if content == "!reset":
        history[user_id] = []
        await message.reply("🗑️ 会話履歴をリセットしました。")
        return

    if content == "!help":
        await message.reply(
            "**コマンド一覧**\n"
            "`!reset` — 自分の会話履歴をリセット\n"
            "`!help` — このメッセージを表示\n\n"
            "**できること**\n"
            "- ファイル読み書き\n"
            "- シェルコマンド実行\n"
            "- コード作成・修正\n"
            "- 画像解析（スクショ送付）"
        )
        return

    has_images = any(
        a.content_type and a.content_type.startswith("image/")
        for a in message.attachments
    )

    if not content and not has_images:
        return

    async with message.channel.typing():
        try:
            if has_images:
                user_content = []
                async with aiohttp.ClientSession() as session:
                    for att in message.attachments:
                        if att.content_type and att.content_type.startswith("image/"):
                            async with session.get(att.url) as resp:
                                img_b64 = base64.standard_b64encode(await resp.read()).decode()
                                user_content.append({
                                    "type": "image",
                                    "source": {
                                        "type": "base64",
                                        "media_type": att.content_type.split(";")[0],
                                        "data": img_b64
                                    }
                                })
                user_content.append({
                    "type": "text",
                    "text": content if content else "この画像について説明してください。"
                })
            else:
                user_content = content

            response = await run_agent(user_id, user_content, message.channel)

        except Exception as e:
            response = f"エラー: {e}"

    await send_long(message, response)


client.run(TOKEN)
