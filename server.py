import discord
from discord.ext import commands
from flask import Flask
from threading import Thread
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from collections import defaultdict
import math
import datetime
import requests
import json

# --- Renderの環境変数から設定を読み込む ---
BOT_TOKEN = os.environ.get('DISCORD_TOKEN')
TARGET_CHANNEL_ID = int(os.environ.get('DISCORD_CHANNEL_ID'))
SHEET_ID = os.environ.get('SHEET_ID')
MESSAGE_LIMIT = 10000 # 取得するメッセージの上限

# --- Flask (Webサーバー) の設定 ---
app = Flask('')
@app.route('/')
def home():
    return "Bot is running!"
def run_flask():
    app.run(host='0.0.0.0', port=8080)
def keep_alive():
    t = Thread(target=run_flask)
    t.start()

# --- Google Sheets への接続設定 ---
def get_sheet():
    # 環境変数からJSON文字列を読み込む
    creds_json_str = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if not creds_json_str:
        raise ValueError("環境変数 GOOGLE_CREDENTIALS_JSON が設定されていません。")

    creds_dict = json.loads(creds_json_str)

    # 辞書から認証情報を読み込む
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SHEET_ID).sheet1
    return sheet

# --- 時間の平均を計算する関数 ---
def calculate_average_time(times):
    sin_sum, cos_sum = 0, 0
    for dt in times:
        seconds_from_midnight = dt.hour * 3600 + dt.minute * 60 + dt.second
        angle = (seconds_from_midnight / 86400) * 2 * math.pi
        sin_sum += math.sin(angle)
        cos_sum += math.cos(angle)
    avg_angle = math.atan2(sin_sum, cos_sum)
    if avg_angle < 0:
        avg_angle += 2 * math.pi
    avg_seconds = (avg_angle / (2 * math.pi)) * 86400
    avg_hour = int(avg_seconds // 3600)
    avg_minute = int((avg_seconds % 3600) // 60)
    return f"{avg_hour:02d}:{avg_minute:02d}"

# --- Discordボットの本体 ---
intents = discord.Intents.default()
intents.messages = True
intents.message_content = True
bot = commands.Bot(command_prefix='!', intents=intents)

@bot.event
async def on_ready():
    print(f'{bot.user}としてログインしました。')

@bot.command()
async def analyze(ctx):
    if ctx.channel.id != TARGET_CHANNEL_ID:
        return

    await ctx.send(f"分析を開始します。最大{MESSAGE_LIMIT}件のメッセージを処理します。少しお待ちください...")

    try:
        target_channel = bot.get_channel(TARGET_CHANNEL_ID)
        user_messages = defaultdict(list)

        async for message in target_channel.history(limit=MESSAGE_LIMIT):
            if not message.author.bot:
                jst_time = message.created_at + datetime.timedelta(hours=9) # JSTに変換
                user_messages[message.author.display_name].append(jst_time)

        await ctx.send("メッセージの集計が完了しました。スプレッドシートに書き込んでいます...")

        sheet = get_sheet()
        sheet.clear()
        header = ["ユーザー名", "平均投稿時間 (JST)", "投稿数"]
        sheet.append_row(header)

        sorted_users = sorted(user_messages.items(), key=lambda item: len(item[1]), reverse=True)

        rows_to_add = []
        for user_name, timestamps in sorted_users:
            if timestamps:
                average_time = calculate_average_time(timestamps)
                rows_to_add.append([user_name, average_time, len(timestamps)])

        if rows_to_add:
            sheet.append_rows(rows_to_add)

        sheet_url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
        await ctx.send(f"分析が完了しました！\n結果はこちら: {sheet_url}")

    except Exception as e:
        await ctx.send(f"エラーが発生しました: {e}")
        print(f"エラー: {e}")

# --- 実行 ---
keep_alive()
bot.run(BOT_TOKEN)
