
import os
import asyncio
import json
import logging
from datetime import datetime, timezone

import discord
import gspread
from discord.ext import tasks
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials

load_dotenv()

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("airbet")

TOKEN = os.environ["DISCORD_BOT_TOKEN"]
SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
CREDENTIALS_FILE = os.environ.get("GOOGLE_CREDENTIALS_FILE", "")
CREDENTIALS_JSON = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
# 固定チャンネルは.envから取得
CHANNELS = {
    "総合": int(os.environ["DISCORD_CHANNEL_GENERAL"]),
    "最高オッズ": int(os.environ["DISCORD_CHANNEL_ODDS"]),
    "最高利益ポイント": int(os.environ["DISCORD_CHANNEL_PROFIT"])
}

# 参加者情報シート名
PARTICIPANT_SHEET = "参加者情報"

SHEET_NAME = "通知キュー"

# 最初は15秒ごとに確認
POLL_SECONDS = 15


def connect_sheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    if CREDENTIALS_JSON.strip():
        # Recommended for hosting: keep the entire service account JSON
        # in a secret environment variable rather than uploading the key file.
        credentials = Credentials.from_service_account_info(
            json.loads(CREDENTIALS_JSON), scopes=scopes
        )
    elif CREDENTIALS_FILE:
        # Backwards compatible with local development.
        credentials = Credentials.from_service_account_file(
            CREDENTIALS_FILE, scopes=scopes
        )
    else:
        raise RuntimeError(
            "GOOGLE_CREDENTIALS_JSON or GOOGLE_CREDENTIALS_FILE is required"
        )

    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)

    return spreadsheet.worksheet(SHEET_NAME)


sheet = connect_sheet()


def get_participant_channels():
    """
    参加者情報シートから
    参加者IDとDiscordチャンネルIDの対応を取得する。
    """

    spreadsheet = sheet.spreadsheet
    participant_sheet = spreadsheet.worksheet(
        PARTICIPANT_SHEET
    )

    # 見出し名から列を探すため、列順に依存しない
    rows = participant_sheet.get_all_records(
        numericise_ignore=["all"]
    )

    channels = {}

    for row in rows:
        participant_id = str(
            row.get("参加者ID", "")
        ).strip()

        channel_id = str(
            row.get("DiscordチャンネルID", "")
        ).strip()

        if not participant_id and not channel_id:
            continue

        if not participant_id:
            raise ValueError(
                "参加者IDが空欄の行があります"
            )

        if not channel_id:
            # チャンネル未設定の参加者は除外
            continue

        if participant_id in channels:
            raise ValueError(
                f"参加者IDが重複しています: {participant_id}"
            )

        if not channel_id.isascii() or not channel_id.isdigit():
            raise ValueError(
                f"{participant_id}のチャンネルIDが不正です"
            )

        channels[participant_id] = int(channel_id)

    return channels


def fetch_pending():
    """未送信通知を取得する"""
    records = sheet.get_all_values()

    pending = []

    for row_number, row in enumerate(records[1:], start=2):
        row = row + [""] * (8 - len(row))

        if row[3].strip() != "未送信":
            continue

        pending.append({
            "row": row_number,
            "id": row[0],
            "destination": row[1],
            "payload": row[2]
        })

    return pending


# 通知キューのD列は状態、E列は作成日時、F列は送信日時、
# G列はDiscordメッセージID、H列はエラー。
# E列の作成日時は書き換えない。
def update_notification(row_number, status,
                        message_id="", error=""):
    sheet.update(
        range_name=f"D{row_number}",
        values=[[status]],
        value_input_option="RAW"
    )

    if status == "送信済":
        sheet.update(
            range_name=f"F{row_number}:H{row_number}",
            values=[[
                datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                str(message_id),
                ""
            ]],
            value_input_option="RAW"
        )

    elif error:
        sheet.update(
            range_name=f"H{row_number}",
            values=[[error[:1000]]],
            value_input_option="RAW"
        )


intents = discord.Intents.default()
bot = discord.Client(intents=intents)

async def send_notification(notification):
    row = notification["row"]
    destination = notification["destination"]

    # まず送信中にする
    await asyncio.to_thread(
        update_notification, row, "送信中"
    )

    try:
        # 固定チャンネルなら直接取得
        if destination in CHANNELS:
            channel_id = CHANNELS[destination]

        else:
            # 参加者情報シートから検索
            participant_channels = await asyncio.to_thread(
                get_participant_channels
            )

            if destination not in participant_channels:
                raise ValueError(
                    f"通知先が見つかりません：{destination}"
                )

            channel_id = participant_channels[destination]

        channel = bot.get_channel(channel_id)

        if channel is None:
            channel = await bot.fetch_channel(channel_id)

        payload = json.loads(notification["payload"])

        content = payload.get("content")
        embeds = [
            discord.Embed.from_dict(e)
            for e in payload.get("embeds", [])
        ]

        sent = await channel.send(
            content=content,
            embeds=embeds,
            allowed_mentions=discord.AllowedMentions.none()
        )

        await asyncio.to_thread(
            update_notification,
            row,
            "送信済",
            str(sent.id)
        )

        logger.info(
            "送信成功: %s / %s",
            notification["id"], sent.id
        )

    except Exception as e:
        # 通信失敗時に投稿だけ成功している可能性がある。
        # 自動再送は行わず手動確認する。
        try:
            await asyncio.to_thread(
                update_notification,
                row,
                "要確認",
                "",
                str(e)
            )
        except Exception:
            logger.exception("状態更新にも失敗しました")

        logger.exception(
            "通知処理失敗: %s",
            notification["id"]
        )


@tasks.loop(seconds=POLL_SECONDS)
async def poll_notifications():
    try:
        pending = await asyncio.to_thread(fetch_pending)

        for notification in pending:
            await send_notification(notification)

    except Exception:
        logger.exception("通知キューの取得に失敗しました")


@poll_notifications.before_loop
async def before_poll():
    await bot.wait_until_ready()



@bot.event
async def on_ready():
    logger.info("Bot起動: %s", bot.user)
    if not poll_notifications.is_running():
        poll_notifications.start()


bot.run(TOKEN)
