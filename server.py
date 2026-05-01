import discord
from discord.ext import commands, tasks
from flask import Flask, request
from threading import Thread
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from collections import defaultdict
import datetime
from dateutil import parser
import requests
import json
import asyncio
from functools import wraps
import queue

BOT_TOKEN = os.environ.get('DISCORD_TOKEN')
TARGET_CHANNEL_ID = int(os.environ.get('DISCORD_CHANNEL_ID'))
SHEET_ID = os.environ.get('SHEET_ID')
TRIGGER_SECRET = os.environ.get('TRIGGER_SECRET')
MESSAGE_LIMIT = int(os.environ.get('MESSAGE_LIMIT', 2000))
PORT = int(os.environ.get('PORT', 8080))

analysis_queue = queue.Queue()

app = Flask('')

@app.route('/')
def home():
    return "Bot is running!"

def run_flask():
    app.run(host='0.0.0.0', port=PORT)

def keep_alive():
    t = Thread(target=run_flask)
    t.start()

def get_spreadsheet():
    creds_json_str = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if not creds_json_str:
        raise ValueError("GOOGLE_CREDENTIALS_JSON is not set.")
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

def parse_timestamp_smart(timestamp_str):
    if not timestamp_str:
        return None
    try:
        dt = parser.parse(str(timestamp_str))
        if dt.tzinfo is None:
             dt = dt.replace(tzinfo=datetime.timezone(datetime.timedelta(hours=9)))
        else:
             dt = dt.astimezone(datetime.timezone(datetime.timedelta(hours=9)))
        return dt
    except Exception as e:
        return None

def load_historical_data(worksheet):
    print("load_historical_data...")
    try:
        header_row = worksheet.row_values(1)
    except Exception:
        header_row = []

    expected_headers = ['ユーザー名', '日付', 'タイムスタンプ']
    
    if not header_row or header_row[0] != expected_headers[0]:
        if worksheet.get_all_values():
            worksheet.insert_row(expected_headers, index=1)
        else:
            worksheet.append_row(expected_headers)
    
    records = worksheet.get_all_records()
    user_daily_first_post = defaultdict(dict)
    valid_count = 0
    
    if not records:
        return user_daily_first_post, None

    for i, record in enumerate(records):
        try:
            user_name = record.get('ユーザー名')
            date_str = record.get('日付')
            timestamp_str = record.get('タイムスタンプ')
            
            if not user_name or not timestamp_str:
                continue

            timestamp_dt = parse_timestamp_smart(timestamp_str)
            
            if timestamp_dt is None:
                continue

            if date_str not in user_daily_first_post[user_name]:
                user_daily_first_post[user_name][date_str] = timestamp_dt
                valid_count += 1
            elif timestamp_dt < user_daily_first_post[user_name][date_str]:
                user_daily_first_post[user_name][date_str] = timestamp_dt

        except Exception:
            continue

    return user_daily_first_post, None

def append_new_data(worksheet, new_posts_list):
    if not new_posts_list:
        return
    
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

async def perform_analysis():
    print("perform_analysis...")
    
    spreadsheet = get_spreadsheet()
    db_sheet = get_or_create_worksheet(spreadsheet, "累計データ")
    user_daily_first_post, _ = load_historical_data(db_sheet)

    target_channel = bot.get_channel(TARGET_CHANNEL_ID)
    if not target_channel:
        return None, None, "Channel not found."

    new_messages = []
    fetch_limit = MESSAGE_LIMIT 

    async for message in target_channel.history(limit=fetch_limit, oldest_first=False):
        new_messages.append(message)

    newly_found_posts = defaultdict(dict)
    for message in new_messages:
        if message.author.bot: continue
        
        timestamp_jst = message.created_at + datetime.timedelta(hours=9)
        if timestamp_jst.hour >= 17: continue

        date_str = timestamp_jst.strftime("%Y-%m-%d")
        user_name = message.author.global_name or message.author.username

        is_new = date_str not in user_daily_first_post[user_name]
        is_earlier = not is_new and timestamp_jst < user_daily_first_post[user_name][date_str]

        if is_new or is_earlier:
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
        append_new_data(db_sheet, new_posts_for_sheet)

    now_jst = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    current_month, current_year = now_jst.month, now_jst.year
    
    if current_month == 1:
        prev_month, prev_year = 12, current_year - 1
    else:
        prev_month, prev_year = current_month - 1, current_year

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
    
    return analysis_data, user_daily_first_post, None

def update_spreadsheet(analysis_data):
    print("update_spreadsheet...")
    try:
        spreadsheet = get_spreadsheet()
        sheet = get_or_create_worksheet(spreadsheet, "起床時刻ランキング")
        
        if not analysis_data:
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
        
    except Exception as e:
        print(f"error: {e}")
        raise

def update_monthly_average_sheet(user_daily_first_post):
    print("update_monthly_average_sheet...")
    try:
        spreadsheet = get_spreadsheet()
        sheet = get_or_create_worksheet(spreadsheet, "月別平均推移")
        
        if not user_daily_first_post:
            return

        sheet.clear()

        all_year_months = set()
        user_monthly_data = defaultdict(lambda: defaultdict(list))

        for user, daily_posts in user_daily_first_post.items():
            for date_str, dt in daily_posts.items():
                ym = f"{dt.year:04d}/{dt.month:02d}"
                all_year_months.add(ym)
                user_monthly_data[user][ym].append(time_to_seconds(dt))

        sorted_yms = sorted(list(all_year_months))

        if not sorted_yms:
            return

        headers = ['ユーザー名'] + sorted_yms
        
        rows = []
        for user, monthly_data in user_monthly_data.items():
            row = [user]
            for ym in sorted_yms:
                if ym in monthly_data and monthly_data[ym]:
                    avg_sec = sum(monthly_data[ym]) / len(monthly_data[ym])
                    row.append(seconds_to_time_str(avg_sec))
                else:
                    row.append("--:--")
            rows.append(row)

        rows.sort(key=lambda x: x[0])

        sheet.update('A1', [headers])
        if rows:
            sheet.update('A2', rows)
        
        requests = [
            { "updateSheetProperties": { "properties": { "sheetId": sheet.id, "gridProperties": { "frozenRowCount": 1, "frozenColumnCount": 1 } }, "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount" } },
            { "repeatCell": { "range": { "sheetId": sheet.id, "startRowIndex": 0, "endRowIndex": 1 }, "cell": { "userEnteredFormat": { "backgroundColor": { "red": 0.2, "green": 0.2, "blue": 0.2 }, "textFormat": { "foregroundColor": { "red": 1, "green": 1, "blue": 1 }, "bold": True }, "horizontalAlignment": "CENTER" } }, "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)" } },
            { "repeatCell": { "range": { "sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": 1 }, "cell": { "userEnteredFormat": { "horizontalAlignment": "CENTER" } }, "fields": "userEnteredFormat.horizontalAlignment" } },
        ]
        sheet.spreadsheet.batch_update({"requests": requests})

    except Exception as e:
        print(f"error: {e}")

intents = discord.Intents.default()
intents.messages = True
intents.message_content = True
bot = commands.Bot(command_prefix='!', intents=intents)

@bot.event
async def on_ready():
    check_queue_task.start()

@tasks.loop(seconds=5)
async def check_queue_task():
    if not analysis_queue.empty():
        analysis_queue.get()
        try:
            analysis_data, user_daily_first_post, error = await perform_analysis()
            if error:
                return
            update_spreadsheet(analysis_data)
            update_monthly_average_sheet(user_daily_first_post)
        except Exception as e:
            print(f"error: {e}")

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
    return 'Analysis triggered.', 200

@bot.command()
async def analyze(ctx):
    if ctx.channel.id != TARGET_CHANNEL_ID: return
    await ctx.send("分析を開始します。少しお待ちください...")
    try:
        analysis_data, user_daily_first_post, error = await perform_analysis()
        if error:
            await ctx.send(f"エラー: {error}")
            return
        update_spreadsheet(analysis_data)
        update_monthly_average_sheet(user_daily_first_post)
        sheet_url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
        await ctx.send(f"分析が完了しました！\n結果はこちら: {sheet_url}")
    except Exception as e:
        await ctx.send(f"エラーが発生しました: {e}")

keep_alive()
bot.run(BOT_TOKEN)
