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
        worksheet = spreadsheet.add_worksheet(title=title, rows="1000", cols="20")
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

# --- 日付パース関数 (強化版) ---
def parse_timestamp_robust(timestamp_str):
    """
    様々な形式の日付文字列をパースしてJSTのdatetimeオブジェクトを返す
    """
    if not timestamp_str:
        return None
    
    # パターン1: ISOフォーマット (2025-11-25T07:00:00)
    try:
        return datetime.datetime.fromisoformat(timestamp_str).replace(tzinfo=datetime.timezone(datetime.timedelta(hours=9)))
    except ValueError:
        pass
    
    # パターン2: スラッシュ区切り (2025/11/25 7:00:00) - GoogleSheetsの自動変換
    formats = [
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M"
    ]
    
    for fmt in formats:
        try:
            dt = datetime.datetime.strptime(str(timestamp_str), fmt)
            return dt.replace(tzinfo=datetime.timezone(datetime.timedelta(hours=9)))
        except ValueError:
            continue
            
    return None

# --- データ永続化ロジック (修正版) ---
def load_historical_data(worksheet):
    print("スプレッドシートから累計データを読み込みます...")
    
    try:
        header_row = worksheet.row_values(1)
    except Exception:
        header_row = []

    expected_headers = ['ユーザー名', '日付', 'タイムスタンプ']
    
    if not header_row or header_row[0] != expected_headers[0]:
        print("⚠️ ヘッダーが見つからないか破損しています。自動修復を実行します...")
        if worksheet.get_all_values():
            worksheet.insert_row(expected_headers, index=1)
        else:
            worksheet.append_row(expected_headers)
        print("✅ ヘッダーを修復しました。")
    
    records = worksheet.get_all_records()
    user_daily_first_post = defaultdict(dict)
    valid_records_count = 0
    
    if not records:
        return user_daily_first_post, None

    for i, record in enumerate(records):
        try:
            user_name = record.get('ユーザー名')
            date_str = record.get('日付')
            timestamp_str = record.get('タイムスタンプ')
            
            if not user_name or not timestamp_str:
                continue

            # 強化されたパース処理
            timestamp_dt = parse_timestamp_robust(timestamp_str)
            
            if timestamp_dt is None:
                # パース失敗時はログに出してスキップ
                if i < 5: # 大量に出ないように最初の5件だけ
                    print(f"⚠️ 日付読み取り失敗 (行{i+2}): {timestamp_str}")
                continue

            # 【重要】重複データがある場合、より早い時間を採用する
            if date_str not in user_daily_first_post[user_name]:
                user_daily_first_post[user_name][date_str] = timestamp_dt
                valid_records_count += 1
            elif timestamp_dt < user_daily_first_post[user_name][date_str]:
                user_daily_first_post[user_name][date_str] = timestamp_dt

        except Exception as e:
            continue

    print(f"累計データの読み込み完了。有効レコード数: {valid_records_count}/{len(records)}")
    return user_daily_first_post, None # last_timestampは一旦使用しない（全取得で整合性重視）

def append_new_data(worksheet, new_posts_list):
    if not new_posts_list:
        return
    print(f"{len(new_posts_list)}件の新規データをスプレッドシートに追記します...")
    
    header_row = worksheet.row_values(1)
    if not header_row or header_row[0] != 'ユーザー名':
        worksheet.insert_row(['ユーザー名', '日付', 'タイムスタンプ'], index=1)

    rows_to_append = []
    for post in new_posts_list:
        rows_to_append.append([
            post['user_name'],
            post['date_str'],
            post['timestamp'].isoformat()
        ])
    worksheet.append_rows(rows_to_append, value_input_option='USER_ENTERED')
    print("データの追記が完了しました。")


# --- メインの分析ロジック ---
async def perform_analysis():
    print("=== 分析処理を開始します ===")
    
    spreadsheet = get_spreadsheet()
    db_sheet = get_or_create_worksheet(spreadsheet, "累計データ")
    # last_timestampによるフィルタリングを廃止し、常に全データを正しく評価する
    user_daily_first_post, _ = load_historical_data(db_sheet)

    print(f"現在の累計データ保持ユーザー数: {len(user_daily_first_post)}")

    target_channel = bot.get_channel(TARGET_CHANNEL_ID)
    if not target_channel:
        print("エラー: 指定されたチャンネルが見つかりません。")
        return None, "指定されたチャンネルが見つかりませんでした。"

    print("Discordから新規メッセージを取得します...")
    new_messages = []
    # 常に一定期間分を遡ってチェック（漏れ防止）
    fetch_limit = MESSAGE_LIMIT 

    async for message in target_channel.history(limit=fetch_limit, oldest_first=False): # 最新から遡る
        new_messages.append(message)
    
    print(f"取得メッセージ数: {len(new_messages)}件")

    newly_found_posts = defaultdict(dict)
    for message in new_messages:
        if message.author.bot: continue
        
        timestamp_jst = message.created_at + datetime.timedelta(hours=9)
        if timestamp_jst.hour >= 17: continue

        date_str = timestamp_jst.strftime("%Y-%m-%d")
        user_name = message.author.global_name or message.author.username

        # まだデータがない、または既存データより早い時間の場合
        is_new = date_str not in user_daily_first_post[user_name]
        is_earlier = not is_new and timestamp_jst < user_daily_first_post[user_name][date_str]

        if is_new:
             newly_found_posts[user_name][date_str] = timestamp_jst
             user_daily_first_post[user_name][date_str] = timestamp_jst
        elif is_earlier:
             # 既存データより早い投稿が見つかった場合、メモリ上は更新するが
             # シートへの追記は重複になるため、ここでは慎重に扱う（今回は追記対象にする）
             newly_found_posts[user_name][date_str] = timestamp_jst
             user_daily_first_post[user_name][date_str] = timestamp_jst

    new_posts_for_sheet = []
    for user_name, daily_posts in newly_found_posts.items():
        for date_str, timestamp in daily_posts.items():
            new_posts_for_sheet.append({
                'user_name': user_name,
                'date_str': date_str,
                'timestamp': timestamp.replace(tzinfo=None)
            })

    if new_posts_for_sheet:
        print(f"シートへの追記対象: {len(new_posts_for_sheet)}件")
        append_new_data(db_sheet, new_posts_for_sheet)
    else:
        print("シートへの追記データはありません。")

    now_jst = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    current_month, current_year = now_jst.month, now_jst.year
    
    if current_month == 1:
        prev_month, prev_year = 12, current_year - 1
    else:
        prev_month, prev_year = current_month - 1, current_year

    print(f"集計対象: {current_year}年{current_month}月 (比較: {prev_year}年{prev_month}月)")

    analysis_data = []
    for user_name, daily_posts in user_daily_first_post.items():
        all_times_sec = [time_to_seconds(dt) for dt in daily_posts.values()]
        
        current_times_sec = []
        previous_times_sec = []
        
        for dt in daily_posts.values():
            sec = time_to_seconds(dt)
            if dt.year == current_year and dt.month == current_month:
                current_times_sec.append(sec)
            if dt.year == prev_year and dt.month == prev_month:
                previous_times_sec.append(sec)

        overall_avg = sum(all_times_sec) / len(all_times_sec) if all_times_sec else None
        current_avg = sum(current_times_sec) / len(current_times_sec) if current_times_sec else None
        previous_avg = sum(previous_times_sec) / len(previous_times_sec) if previous_times_sec else None
        delta = current_avg - previous_avg if (current_avg is not None and previous_avg is not None) else None

        analysis_data.append({
            'userName': user_name, 
            'overall_avg_sec': overall_avg, 
            'overall_count': len(all_times_sec),
            'current_avg_sec': current_avg, 
            'previous_avg_sec': previous_avg, 
            'delta_sec': delta
        })

    analysis_data.sort(key=lambda x: x['current_avg_sec'] if x['current_avg_sec'] is not None else float('inf'))
    
    print(f"ランキング生成完了: {len(analysis_data)}名")
    return analysis_data, None


# --- スプレッドシート更新ロジック ---
def update_spreadsheet(analysis_data):
    print("ランキングシートの更新を開始します...")
    try:
        spreadsheet = get_spreadsheet()
        sheet = get_or_create_worksheet(spreadsheet, "起床時刻ランキング")
        
        if not analysis_data:
            print("更新するデータがありません。")
            return

        sheet.clear()
        
        now_jst = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
        timestamp_str = f"起床時刻ランキング (最終更新: {now_jst.strftime('%Y/%m/%d %H:%M')})"
        
        sheet.update('A1', [[timestamp_str]])
        
        headers = ['順位', 'ユーザー名', '今月の平均', '先月の平均', '変化', '累計平均', '累計日数']
        sheet.update('A2', [headers])
        
        rows = []
        for index, user in enumerate(analysis_data):
            rank = index + 1
            rows.append([
                rank,
                user['userName'],
                seconds_to_time_str(user['current_avg_sec']),
                seconds_to_time_str(user['previous_avg_sec']),
                format_delta_seconds(user['delta_sec']),
                seconds_to_time_str(user['overall_avg_sec']),
                user['overall_count']
            ])

        if rows:
            sheet.update('A3', rows)
            print(f"{len(rows)}行のデータを書き込みました。")

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
