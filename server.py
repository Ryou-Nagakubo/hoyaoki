import discord
from discord.ext import commands, tasks
from flask import Flask, request
from threading import Thread
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from collections import defaultdict
import datetime
import requests
import json
import asyncio
from functools import wraps
import queue
import time

# --- Renderの環境変数から設定を読み込む ---
BOT_TOKEN = os.environ.get('DISCORD_TOKEN')
TARGET_CHANNEL_ID = int(os.environ.get('DISCORD_CHANNEL_ID'))
SHEET_ID = os.environ.get('SHEET_ID')
TRIGGER_SECRET = os.environ.get('TRIGGER_SECRET')
# 初回実行時のみ使用するメッセージ取得上限
MESSAGE_LIMIT = 20000

# --- グローバルなタスクキュー ---
analysis_queue = queue.Queue()

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
def get_spreadsheet():
    creds_json_str = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if not creds_json_str:
        raise ValueError("環境変数 GOOGLE_CREDENTIALS_JSON が設定されていません。")
    creds_dict = json.loads(creds_json_str)
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SHEET_ID)
    return spreadsheet

def get_or_create_worksheet(spreadsheet, title):
    try:
        worksheet = spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=title, rows="100", cols="20")
    return worksheet

# --- 時刻計算のヘルパー関数 ---
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

# --- データ永続化ロジック ---
def load_historical_data(worksheet):
    print("スプレッドシートから累計データを読み込みます...")
    records = worksheet.get_all_records()
    user_daily_first_post = defaultdict(dict)
    last_timestamp = None
    if not records:
        return user_daily_first_post, None

    for record in records:
        try:
            user_name = record['ユーザー名']
            date_str = record['日付']
            timestamp_str = record['タイムスタンプ']
            # JSTとして解釈
            timestamp_dt = datetime.datetime.fromisoformat(timestamp_str).replace(tzinfo=datetime.timezone(datetime.timedelta(hours=9)))
            user_daily_first_post[user_name][date_str] = timestamp_dt
        except (KeyError, ValueError) as e:
            print(f"レコードの読み込みエラー: {record}, エラー: {e}")
            continue

    # 最後のタイムスタンプを取得 (UTCに変換)
    last_timestamp_iso = records[-1]['タイムスタンプ']
    last_timestamp = datetime.datetime.fromisoformat(last_timestamp_iso)

    print(f"累計データの読み込み完了。最終記録時刻: {last_timestamp}")
    return user_daily_first_post, last_timestamp

def append_new_data(worksheet, new_posts_list):
    if not new_posts_list:
        return
    print(f"{len(new_posts_list)}件の新規データをスプレッドシートに追記します...")
    # ヘッダーがなければ追記
    if not worksheet.get_all_values():
         worksheet.append_row(['ユーザー名', '日付', 'タイムスタンプ'])

    rows_to_append = []
    for post in new_posts_list:
        rows_to_append.append([
            post['user_name'],
            post['date_str'],
            post['timestamp'].isoformat() # タイムゾーン情報なしで保存
        ])
    worksheet.append_rows(rows_to_append, value_input_option='USER_ENTERED')
    print("データの追記が完了しました。")


# --- メインの分析ロジック ---
async def perform_analysis():
    print("分析処理を開始します...")
    
    # 1. スプレッドシートから累計データを読み込む
    spreadsheet = get_spreadsheet()
    db_sheet = get_or_create_worksheet(spreadsheet, "累計データ")
    user_daily_first_post, last_timestamp = load_historical_data(db_sheet)

    # 2. Discordから新しいメッセージを取得
    target_channel = bot.get_channel(TARGET_CHANNEL_ID)
    if not target_channel:
        return None, "指定されたチャンネルが見つかりませんでした。"

    print("Discordから新規メッセージを取得します...")
    new_messages = []
    # last_timestampがNoneなら（初回実行時）、上限まで取得
    fetch_limit = None if last_timestamp else MESSAGE_LIMIT
    # afterはUTCである必要があるため、タイムゾーン情報を削除
    after_utc = last_timestamp.replace(tzinfo=None) if last_timestamp else None

    async for message in target_channel.history(limit=fetch_limit, after=after_utc, oldest_first=True):
        new_messages.append(message)
    
    print(f"{len(new_messages)}件の新規メッセージを取得しました。")

    # 3. 新しいメッセージを処理して、日ごとの最初の投稿を抽出
    newly_found_posts = defaultdict(dict)
    for message in new_messages:
        if message.author.bot: continue
        
        timestamp_jst = message.created_at + datetime.timedelta(hours=9)
        if timestamp_jst.hour >= 17: continue

        date_str = timestamp_jst.strftime("%Y-%m-%d")
        user_name = message.author.global_name or message.author.username

        if date_str not in newly_found_posts[user_name] and date_str not in user_daily_first_post[user_name]:
             newly_found_posts[user_name][date_str] = timestamp_jst

    # 4. 抽出した新しいデータをスプレッドシートに追記
    new_posts_for_sheet = []
    for user_name, daily_posts in newly_found_posts.items():
        for date_str, timestamp in daily_posts.items():
            new_posts_for_sheet.append({
                'user_name': user_name,
                'date_str': date_str,
                'timestamp': timestamp.replace(tzinfo=None) # UTCとして保存
            })
            # メモリ上の累計データも更新
            user_daily_first_post[user_name][date_str] = timestamp

    if new_posts_for_sheet:
        append_new_data(db_sheet, new_posts_for_sheet)

    # 5. 全データを使ってランキングを計算
    now_jst = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    current_month, current_year = now_jst.month, now_jst.year
    prev_month_date = now_jst.replace(day=1) - datetime.timedelta(days=1)
    prev_month, prev_year = prev_month_date.month, prev_month_date.year

    analysis_data = []
    for user_name, daily_posts in user_daily_first_post.items():
        all_times_sec = [time_to_seconds(dt) for dt in daily_posts.values()]
        current_times_sec = [time_to_seconds(dt) for dt in daily_posts.values() if dt.year == current_year and dt.month == current_month]
        previous_times_sec = [time_to_seconds(dt) for dt in daily_posts.values() if dt.year == prev_year and dt.month == prev_month]

        if not current_times_sec: continue # 今月の記録がないユーザーはランキングから除外

        overall_avg = sum(all_times_sec) / len(all_times_sec) if all_times_sec else None
        current_avg = sum(current_times_sec) / len(current_times_sec)
        previous_avg = sum(previous_times_sec) / len(previous_times_sec) if previous_times_sec else None
        delta = current_avg - previous_avg if previous_avg is not None else None

        analysis_data.append({
            'userName': user_name, 'overall_avg_sec': overall_avg, 'overall_count': len(all_times_sec),
            'current_avg_sec': current_avg, 'previous_avg_sec': previous_avg, 'delta_sec': delta
        })

    analysis_data.sort(key=lambda x: x['current_avg_sec'])
    return analysis_data, None


# --- スプレッドシート更新ロジック ---
def update_spreadsheet(analysis_data):
    print("ランキングシートの更新を開始します...")
    try:
        spreadsheet = get_spreadsheet()
        sheet = get_or_create_worksheet(spreadsheet, "起床時刻ランキング")
        sheet.clear()
        
        now_jst = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
        timestamp_str = f"起床時刻ランキング (最終更新: {now_jst.strftime('%Y/%m/%d %H:%M')})"
        
        sheet.update('A1', [[timestamp_str]])
        
        headers = ['順位', 'ユーザー名', '今月の平均', '先月の平均', '変化', '累計平均', '累計日数']
        sheet.update('A2', [headers])
        
        rows = []
        for index, user in enumerate(analysis_data):
            rows.append([
                index + 1,
                user['userName'],
                seconds_to_time_str(user['current_avg_sec']),
                seconds_to_time_str(user['previous_avg_sec']),
                format_delta_seconds(user['delta_sec']),
                seconds_to_time_str(user['overall_avg_sec']),
                user['overall_count']
            ])

        if rows:
            sheet.update('A3', rows)

        # --- シートの書式設定 ---
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
            { "updateDimensionProperties": { "range": { "sheetId": sheet.id, "dimension": "COLUMNS", "startIndex": 5, "endIndex": 7 }, "properties": { "pixelSize": 75 }, "fields": "pixelSize" } },
        ]
        sheet.spreadsheet.batch_update({"requests": requests})
            
        print("ランキングシートの更新が完了しました。")
        
    except Exception as e:
        print(f"スプレッドシート更新中にエラーが発生: {e}")
        raise

# --- Discordボットの本体とWebhook ---
intents = discord.Intents.default()
intents.messages = True
intents.message_content = True
bot = commands.Bot(command_prefix='!', intents=intents)

@bot.event
async def on_ready():
    print(f'{bot.user}としてログインしました。')
    check_queue_task.start()

@tasks.loop(seconds=5)
async def check_queue_task():
    if not analysis_queue.empty():
        print("キューからタスクを検出。分析処理を実行します。")
        analysis_queue.get()
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
    analysis_queue.put(1)
    print("分析タスクをキューに追加しました。")
    return 'Analysis triggered.', 200

@bot.command()
async def analyze(ctx):
    if ctx.channel.id != TARGET_CHANNEL_ID: return
    await ctx.send("分析を開始します。少しお待ちください...")
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
time.sleep(15)
bot.run(BOT_TOKEN)
