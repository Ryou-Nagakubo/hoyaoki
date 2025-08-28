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
import time
import schedule

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
def time_to_seconds(dt):
    return dt.hour * 3600 + dt.minute * 60 + dt.second

def seconds_to_time_str(seconds):
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    return f"{h:02d}:{m:02d}"

# --- メインの分析ロジック ---
async def perform_analysis():
    print("分析処理を開始します...")
    target_channel = bot.get_channel(TARGET_CHANNEL_ID)
    if not target_channel:
        print(f"エラー: チャンネルID {TARGET_CHANNEL_ID} が見つかりません。")
        return None, "指定されたチャンネルが見つかりませんでした。"

    # ユーザーごと、日付ごとの最初の投稿を記録
    # { user_name: { "YYYY-MM-DD": earliest_datetime } }
    user_daily_first_post = defaultdict(dict)
    
    # JSTでの「今日」の日付を取得
    today_jst = (datetime.datetime.utcnow() + datetime.timedelta(hours=9)).date()

    async for message in target_channel.history(limit=MESSAGE_LIMIT):
        if message.author.bot:
            continue

        timestamp_utc = message.created_at
        timestamp_jst = timestamp_utc + datetime.timedelta(hours=9)
        
        # 17:00以降の投稿は無視
        if timestamp_jst.hour >= 17:
            continue

        date_str = timestamp_jst.strftime("%Y-%m-%d")
        user_name = message.author.global_name or message.author.username

        # その日の最初の投稿でなければ記録を更新
        if date_str not in user_daily_first_post[user_name] or timestamp_jst < user_daily_first_post[user_name][date_str]:
            user_daily_first_post[user_name][date_str] = timestamp_jst

    # 平均起床時間を計算
    analysis_data = []
    for user_name, daily_posts in user_daily_first_post.items():
        wake_up_times_seconds = [time_to_seconds(dt) for dt in daily_posts.values()]
        if not wake_up_times_seconds:
            continue
        
        average_seconds = sum(wake_up_times_seconds) / len(wake_up_times_seconds)
        analysis_data.append({
            'userName': user_name,
            'postCount': len(wake_up_times_seconds),
            'averageWakeUpSeconds': average_seconds
        })

    # 投稿数でソート
    analysis_data.sort(key=lambda x: x['postCount'], reverse=True)
    return analysis_data, None

# --- スプレッドシート更新ロジック ---
def update_spreadsheet(analysis_data):
    print("スプレッドシートの更新を開始します...")
    try:
        sheet = get_sheet()
        
        # --- シートの初期化 ---
        # 値をすべてクリアする (行や列の削除は行わない)
        sheet.clear()

        headers = ['順位', 'ユーザー名', '記録日数', '平均起床時間']
        
        # ヘッダーを書き込み
        sheet.update('A1', [headers])
        
        rows = []
        for index, user in enumerate(analysis_data):
            avg_time_str = seconds_to_time_str(user['averageWakeUpSeconds'])
            rows.append([
                index + 1,
                user['userName'],
                user['postCount'],
                avg_time_str
            ])

        if rows:
            # A2からデータを書き込み
            sheet.update('A2', rows)

        # --- 書式設定と列幅指定をまとめて実行 ---
        requests = [
            # ヘッダー行を固定
            { "updateSheetProperties": { "properties": { "sheetId": sheet.id, "gridProperties": { "frozenRowCount": 1 } }, "fields": "gridProperties.frozenRowCount" } },
            # ヘッダーの書式設定 (A1:D1)
            { "repeatCell": { "range": { "sheetId": sheet.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 4 }, "cell": { "userEnteredFormat": { "backgroundColor": { "red": 0.06, "green": 0.68, "blue": 0.86 }, "textFormat": { "foregroundColor": { "red": 1.0, "green": 1.0, "blue": 1.0 }, "bold": True }, "horizontalAlignment": "CENTER" } }, "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)" } },
            # 全体の垂直方向の配置 (A:D)
            { "repeatCell": { "range": { "sheetId": sheet.id, "startColumnIndex": 0, "endColumnIndex": 4 }, "cell": { "userEnteredFormat": { "verticalAlignment": "MIDDLE" } }, "fields": "userEnteredFormat.verticalAlignment" } },
            # 特定列の水平方向の配置 (A, C, D)
            { "repeatCell": { "range": { "sheetId": sheet.id, "startColumnIndex": 0, "endColumnIndex": 1 }, "cell": { "userEnteredFormat": { "horizontalAlignment": "CENTER" } }, "fields": "userEnteredFormat.horizontalAlignment" } },
            { "repeatCell": { "range": { "sheetId": sheet.id, "startColumnIndex": 2, "endColumnIndex": 4 }, "cell": { "userEnteredFormat": { "horizontalAlignment": "CENTER" } }, "fields": "userEnteredFormat.horizontalAlignment" } },
            # 列幅を手動で指定
            { "updateDimensionProperties": { "range": { "sheetId": sheet.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1 }, "properties": { "pixelSize": 60 }, "fields": "pixelSize" } }, # A列: 順位
            { "updateDimensionProperties": { "range": { "sheetId": sheet.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2 }, "properties": { "pixelSize": 180 }, "fields": "pixelSize" } }, # B列: ユーザー名
            { "updateDimensionProperties": { "range": { "sheetId": sheet.id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3 }, "properties": { "pixelSize": 80 }, "fields": "pixelSize" } }, # C列: 記録日数
            { "updateDimensionProperties": { "range": { "sheetId": sheet.id, "dimension": "COLUMNS", "startIndex": 3, "endIndex": 4 }, "properties": { "pixelSize": 120 }, "fields": "pixelSize" } }  # D列: 平均起床時間
        ]
        sheet.spreadsheet.batch_update({"requests": requests})
            
        print("スプレッドシートの更新が完了しました。")
        
    except Exception as e:
        print(f"スプレッドシート更新中にエラーが発生: {e}")
        raise

# --- Discordボットの本体 ---
intents = discord.Intents.default()
intents.messages = True
intents.message_content = True
bot = commands.Bot(command_prefix='!', intents=intents)

@bot.event
async def on_ready():
    print(f'{bot.user}としてログインしました。')
    # スケジューラを別スレッドで起動
    schedule_thread = Thread(target=run_scheduler)
    schedule_thread.daemon = True
    schedule_thread.start()
    print("自動集計スケジューラを起動しました。")

# --- 自動集計ジョブ ---
def daily_job():
    # botが準備完了するまで待つ
    while not bot.is_ready():
        time.sleep(1)
    # asyncioのイベントをメインスレッドで安全に実行
    asyncio.run_coroutine_threadsafe(scheduled_analysis(), bot.loop)

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

def run_scheduler():
    # 毎日日本時間の深夜2時に実行
    schedule.every().day.at("17:00").do(daily_job) # UTCで17:00 = JSTで2:00
    while True:
        schedule.run_pending()
        time.sleep(60)

# --- テスト用コマンド ---
@bot.command()
async def analyze(ctx):
    if ctx.channel.id != TARGET_CHANNEL_ID:
        return

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
