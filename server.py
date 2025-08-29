import discord
from discord.ext import commands
from flask import Flask, request
from threading import Thread
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from collections import defaultdict
import datetime
import requests
import json
import time
import asyncio
from functools import wraps

# --- Renderの環境変数から設定を読み込む ---
BOT_TOKEN = os.environ.get('DISCORD_TOKEN')
TARGET_CHANNEL_ID = int(os.environ.get('DISCORD_CHANNEL_ID'))
SHEET_ID = os.environ.get('SHEET_ID')
TRIGGER_SECRET = os.environ.get('TRIGGER_SECRET') # GASからの合図を認証する秘密のキー
MESSAGE_LIMIT = 20000

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
# (この部分は変更ありません)
def get_sheet():
    creds_json_str = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if not creds_json_str:
        raise ValueError("環境変数 GOOGLE_CREDENTIALS_JSON が設定されていません。")
    creds_dict = json.loads(creds_json_str)
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SHEET_ID).sheet1
    return sheet

# --- 時刻計算のヘルパー関数 ---
# (この部分は変更ありません)
def time_to_seconds(dt):
    return dt.hour * 3600 + dt.minute * 60 + dt.second
def seconds_to_time_str(seconds):
    if seconds is None: return "--:--"
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    return f"{h:02d}:{m:02d}"
def format_delta_seconds(seconds):
    if seconds is None: return "N/A"
    total_minutes = round(seconds / 60)
    sign = "+" if total_minutes >= 0 else "-"
    return f"{sign}{abs(total_minutes)}分"

# --- メインの分析ロジック ---
# (この部分は変更ありません)
async def perform_analysis():
    print("分析処理を開始します...")
    # ... (前回のコードと同じ内容のため省略) ...
    target_channel = bot.get_channel(TARGET_CHANNEL_ID)
    if not target_channel:
        return None, "指定されたチャンネルが見つかりませんでした。"
    now_jst = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    current_month = now_jst.month
    current_year = now_jst.year
    prev_month_date = now_jst.replace(day=1) - datetime.timedelta(days=1)
    prev_month = prev_month_date.month
    prev_year = prev_month_date.year
    user_daily_first_post = defaultdict(dict)
    async for message in target_channel.history(limit=MESSAGE_LIMIT):
        if message.author.bot: continue
        timestamp_jst = message.created_at + datetime.timedelta(hours=9)
        if timestamp_jst.hour >= 17: continue
        date_str = timestamp_jst.strftime("%Y-%m-%d")
        user_name = message.author.global_name or message.author.username
        if date_str not in user_daily_first_post[user_name] or timestamp_jst < user_daily_first_post[user_name][date_str]:
            user_daily_first_post[user_name][date_str] = timestamp_jst
    user_monthly_stats = defaultdict(lambda: {"current": [], "previous": []})
    for user_name, daily_posts in user_daily_first_post.items():
        for post_time in daily_posts.values():
            seconds = time_to_seconds(post_time)
            if post_time.year == current_year and post_time.month == current_month:
                user_monthly_stats[user_name]["current"].append(seconds)
            elif post_time.year == prev_year and post_time.month == prev_month:
                user_monthly_stats[user_name]["previous"].append(seconds)
    analysis_data = []
    for user_name, stats in user_monthly_stats.items():
        current_avg = sum(stats["current"]) / len(stats["current"]) if stats["current"] else None
        previous_avg = sum(stats["previous"]) / len(stats["previous"]) if stats["previous"] else None
        delta = current_avg - previous_avg if current_avg is not None and previous_avg is not None else None
        if current_avg is not None:
            analysis_data.append({
                'userName': user_name, 'current_avg_sec': current_avg, 'current_count': len(stats["current"]),
                'previous_avg_sec': previous_avg, 'previous_count': len(stats["previous"]), 'delta_sec': delta
            })
    analysis_data.sort(key=lambda x: x['current_avg_sec'])
    return analysis_data, None


# --- スプレッドシート更新ロジック ---
# (この部分は変更ありません)
def update_spreadsheet(analysis_data):
    print("スプレッドシートの更新を開始します...")
    # ... (前回のコードと同じ内容のため省略) ...
    try:
        sheet = get_sheet()
        sheet.clear()
        now_jst = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
        timestamp_str = f"起床時刻ランキング (最終更新: {now_jst.strftime('%Y/%m/%d %H:%M')})"
        sheet.update('A1', [[timestamp_str]])
        headers = ['順位', 'ユーザー名', '今月の平均', '先月の平均', '変化', '今月の記録', '先月の記録']
        sheet.update('A2', [headers])
        rows = []
        for index, user in enumerate(analysis_data):
            rows.append([
                index + 1, user['userName'], seconds_to_time_str(user['current_avg_sec']),
                seconds_to_time_str(user['previous_avg_sec']), format_delta_seconds(user['delta_sec']),
                user['current_count'], user['previous_count']
            ])
        if rows:
            sheet.update('A3', rows)
        requests = [
            { "updateSheetProperties": { "properties": { "sheetId": sheet.id, "gridProperties": { "frozenRowCount": 2 } }, "fields": "gridProperties.frozenRowCount" } },
            { "mergeCells": { "range": { "sheetId": sheet.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 7 }, "mergeType": "MERGE_ALL" } },
            { "repeatCell": { "range": { "sheetId": sheet.id, "startRowIndex": 0, "endRowIndex": 1 }, "cell": { "userEnteredFormat": { "textFormat": { "bold": True, "fontSize": 12 }, "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE" } }, "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment)" } },
            { "repeatCell": { "range": { "sheetId": sheet.id, "startRowIndex": 1, "endRowIndex": 2 }, "cell": { "userEnteredFormat": { "backgroundColor": { "red": 0.2, "green": 0.2, "blue": 0.2 }, "textFormat": { "foregroundColor": { "red": 1, "green": 1, "blue": 1 }, "bold": True }, "horizontalAlignment": "CENTER" } }, "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)" } },
            { "repeatCell": { "range": { "sheetId": sheet.id, "startRowIndex": 2 }, "cell": { "userEnteredFormat": { "verticalAlignment": "MIDDLE" } }, "fields": "userEnteredFormat.verticalAlignment" } },
            { "repeatCell": { "range": { "sheetId": sheet.id, "startRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 1 }, "cell": { "userEnteredFormat": { "horizontalAlignment": "CENTER" } }, "fields": "userEnteredFormat.horizontalAlignment" } },
            { "repeatCell": { "range": { "sheetId": sheet.id, "startRowIndex": 2, "startColumnIndex": 2, "endColumnIndex": 7 }, "cell": { "userEnteredFormat": { "horizontalAlignment": "CENTER" } }, "fields": "userEnteredFormat.horizontalAlignment" } },
            { "updateDimensionProperties": { "range": { "sheetId": sheet.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1 }, "properties": { "pixelSize": 40 }, "fields": "pixelSize" } },
            { "updateDimensionProperties": { "range": { "sheetId": sheet.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2 }, "properties": { "pixelSize": 120 }, "fields": "pixelSize" } },
            { "updateDimensionProperties": { "range": { "sheetId": sheet.id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 5 }, "properties": { "pixelSize": 75 }, "fields": "pixelSize" } },
            { "updateDimensionProperties": { "range": { "sheetId": sheet.id, "dimension": "COLUMNS", "startIndex": 5, "endIndex": 7 }, "properties": { "pixelSize": 60 }, "fields": "pixelSize" } },
        ]
        sheet.spreadsheet.batch_update({"requests": requests})
        print("スプレッドシートの更新が完了しました。")
    except Exception as e:
        print(f"スプレッドシート更新中にエラーが発生: {e}")
        raise

# --- 自動集計（Webhook経由） ---
async def scheduled_analysis():
    print("スケジュールされた自動集計を開始します...")
    try:
        analysis_data, error = await perform_analysis()
        if error:
            print(f"自動集計エラー: {error}")
            return
        update_spreadsheet(analysis_data)
    except Exception as e:
        print(f"自動集計中に致命的なエラーが発生しました: {e}")

def require_secret(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not TRIGGER_SECRET or request.headers.get('X-Trigger-Secret') != TRIGGER_SECRET:
            return 'Unauthorized', 401
        return f(*args, **kwargs)
    return decorated_function

@app.route('/trigger-analysis', methods=['POST'])
@require_secret
def handle_trigger_analysis():
    def run_in_background():
        loop = bot.loop
        asyncio.run_coroutine_threadsafe(scheduled_analysis(), loop)
    
    analysis_thread = Thread(target=run_in_background)
    analysis_thread.start()
    return 'Analysis triggered.', 200

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
    if ctx.channel.id != TARGET_CHANNEL_ID: return
    await ctx.send("テスト分析を開始します。少しお待ちください...")
    try:
        analysis_data, error = await perform_analysis()
        if error:
            await ctx.send(f"エラー: {error}")
            return
        update_spreadsheet(analysis_data)
        sheet_url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
        await ctx.send(f"分析が完了しました！\n結果はこちら: {sheet_url}")
    except Exception as e:
        await ctx.send(f"エラーが発生しました: {e}")
        print(f"コマンド実行エラー: {e}")

# --- 実行 ---
keep_alive()
bot.run(BOT_TOKEN)
