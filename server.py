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
import re

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

def time_to_seconds(dt, is_sleep=False):
    if dt is None: return None
    h = dt.hour
    # 就寝時間 (18:00~03:59) は平均計算のため、0~3時を24~27時として扱う
    if is_sleep and h < 12:
        h += 24
    return h * 3600 + dt.minute * 60 + dt.second

def seconds_to_time_str(seconds, is_sleep=False):
    if seconds is None: return "--:--"
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    if is_sleep and h >= 24:
        h -= 24
    return f"{h:02d}:{m:02d}"

def format_delta_seconds(seconds):
    if seconds is None: return "N/A"
    total_minutes = round(seconds / 60)
    sign = "+" if total_minutes >= 0 else "-"
    return f"{sign}{abs(total_minutes)}分"

def parse_timestamp_smart(timestamp_str):
    if not timestamp_str: return None
    try:
        dt = parser.parse(str(timestamp_str))
        if dt.tzinfo is None:
             dt = dt.replace(tzinfo=datetime.timezone(datetime.timedelta(hours=9)))
        else:
             dt = dt.astimezone(datetime.timezone(datetime.timedelta(hours=9)))
        return dt
    except Exception:
        return None

def extract_time_from_text(text):
    match = re.search(r'(\d{1,2})[:：](\d{2})', text)
    if match: return int(match.group(1)), int(match.group(2))
    match = re.search(r'(\d{1,2})時(?:(\d{1,2})分)?', text)
    if match: return int(match.group(1)), int(match.group(2) or 0)
    match = re.search(r'(\d{1,2})時半', text)
    if match: return int(match.group(1)), 30
    return None, None

def load_historical_data(worksheet):
    try: header_row = worksheet.row_values(1)
    except Exception: header_row = []

    expected_headers = ['ユーザー名', '日付', '起床時刻', '就寝時刻']
    
    # 互換性チェック＆マイグレーション
    if not header_row or header_row[0] != 'ユーザー名':
        worksheet.append_row(expected_headers)
    elif len(header_row) == 3 and header_row[2] == 'タイムスタンプ':
        print("⚠️ 旧形式のデータを検出。起床時刻へのマイグレーションを実行します...")
        all_values = worksheet.get_all_values()
        worksheet.clear()
        new_values = [expected_headers]
        for row in all_values[1:]:
            if len(row) >= 3: new_values.append([row[0], row[1], row[2], ""])
        worksheet.update('A1', new_values)
        print("✅ マイグレーション完了。")
    
    records = worksheet.get_all_records()
    user_daily_data = defaultdict(lambda: defaultdict(dict))
    
    for record in records:
        try:
            user_name = record.get('ユーザー名')
            date_str = record.get('日付')
            wake_dt = parse_timestamp_smart(record.get('起床時刻', ''))
            sleep_dt = parse_timestamp_smart(record.get('就寝時刻', ''))
            
            if not user_name or not date_str: continue
            if wake_dt: user_daily_data[user_name][date_str]['wake'] = wake_dt
            if sleep_dt: user_daily_data[user_name][date_str]['sleep'] = sleep_dt
        except Exception:
            continue

    return user_daily_data

def save_historical_data(worksheet, user_daily_data):
    worksheet.clear()
    headers = ['ユーザー名', '日付', '起床時刻', '就寝時刻']
    rows = [headers]
    for user_name, daily in user_daily_data.items():
        for date_str, times in daily.items():
            w = times.get('wake')
            s = times.get('sleep')
            rows.append([user_name, date_str, w.isoformat() if w else "", s.isoformat() if s else ""])
    worksheet.update('A1', rows)

async def perform_analysis():
    print("perform_analysis...")
    
    spreadsheet = get_spreadsheet()
    db_sheet = get_or_create_worksheet(spreadsheet, "累計データ")
    user_daily_data = load_historical_data(db_sheet)

    target_channel = bot.get_channel(TARGET_CHANNEL_ID)
    if not target_channel:
        return None, None, [], "Channel not found."

    new_messages = []
    async for message in target_channel.history(limit=MESSAGE_LIMIT, oldest_first=False):
        new_messages.append(message)

    user_id_map = {}
    
    for message in new_messages:
        if message.author.bot: continue
        
        timestamp_jst = message.created_at + datetime.timedelta(hours=9)
        user_name = message.author.global_name or message.author.username
        user_id_map[user_name] = message.author.id
        
        # 朝4時を1日の境目として日付を決定
        logical_date_dt = timestamp_jst - datetime.timedelta(hours=4)
        date_str = logical_date_dt.strftime("%Y-%m-%d")
        
        hour = timestamp_jst.hour
        is_wake = 4 <= hour < 18
        record_key = 'wake' if is_wake else 'sleep'
        dt_to_record = timestamp_jst
        
        # 手動指定時刻の抽出 (起床のみ)
        if is_wake:
            ex_h, ex_m = extract_time_from_text(message.content)
            if ex_h is not None and ex_m is not None:
                if 4 <= ex_h < 18:
                    dt_to_record = dt_to_record.replace(hour=ex_h, minute=ex_m, second=0, microsecond=0)
        
        # 既存より古い(早い)データなら採用
        existing_dt = user_daily_data[user_name].get(date_str, {}).get(record_key)
        if existing_dt is None or dt_to_record < existing_dt:
            user_daily_data[user_name][date_str][record_key] = dt_to_record

    save_historical_data(db_sheet, user_daily_data)

    now_jst = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    current_month, current_year = now_jst.month, now_jst.year
    prev_month, prev_year = (12, current_year - 1) if current_month == 1 else (current_month - 1, current_year)

    analysis_data = []
    for user_name, daily_posts in user_daily_data.items():
        all_wakes = []
        cur_wakes, prev_wakes = [], []
        cur_sleeps, prev_sleeps = [], []
        
        for dt_dict in daily_posts.values():
            if 'wake' in dt_dict:
                w_dt = dt_dict['wake']
                sec = time_to_seconds(w_dt)
                all_wakes.append(sec)
                if w_dt.year == current_year and w_dt.month == current_month: cur_wakes.append(sec)
                if w_dt.year == prev_year and w_dt.month == prev_month: prev_wakes.append(sec)
            
            if 'sleep' in dt_dict:
                s_dt = dt_dict['sleep']
                s_sec = time_to_seconds(s_dt, is_sleep=True)
                # 就寝時刻は論理日付の年月でカウント
                s_logical = s_dt - datetime.timedelta(hours=4)
                if s_logical.year == current_year and s_logical.month == current_month: cur_sleeps.append(s_sec)
                if s_logical.year == prev_year and s_logical.month == prev_month: prev_sleeps.append(s_sec)

        cur_w_avg = sum(cur_wakes) / len(cur_wakes) if cur_wakes else None
        prev_w_avg = sum(prev_wakes) / len(prev_wakes) if prev_wakes else None
        
        analysis_data.append({
            'userName': user_name, 
            'overall_wake_avg': sum(all_wakes) / len(all_wakes) if all_wakes else None, 
            'overall_count': len(all_wakes),
            'current_wake_avg': cur_w_avg, 
            'previous_wake_avg': prev_w_avg, 
            'current_sleep_avg': sum(cur_sleeps) / len(cur_sleeps) if cur_sleeps else None,
            'previous_sleep_avg': sum(prev_sleeps) / len(prev_sleeps) if prev_sleeps else None,
            'delta_sec': cur_w_avg - prev_w_avg if (cur_w_avg is not None and prev_w_avg is not None) else None
        })

    # 起床時間の早い順にソート
    analysis_data.sort(key=lambda x: x['current_wake_avg'] if x['current_wake_avg'] is not None else float('inf'))
    
    # === 未投稿者の抽出 ===
    today_str = (now_jst - datetime.timedelta(hours=4)).strftime("%Y-%m-%d")
    active_users = {u for u, d in user_daily_data.items() if any(k.startswith(f"{current_year}-{current_month:02d}") for k in d.keys())}
    
    missing_users = []
    for user in active_users:
        if 'wake' not in user_daily_data[user].get(today_str, {}):
            if user in user_id_map:
                missing_users.append(f"<@{user_id_map[user]}>")

    return analysis_data, user_daily_data, missing_users, None

def update_spreadsheet(analysis_data):
    try:
        spreadsheet = get_spreadsheet()
        sheet = get_or_create_worksheet(spreadsheet, "起床時刻ランキング")
        if not analysis_data: return
        sheet.clear()
        
        now_jst = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
        sheet.update('A1', [[f"起床時刻ランキング (最終更新: {now_jst.strftime('%Y/%m/%d %H:%M')})"]])
        
        headers = ['順位', 'ユーザー名', '今月の起床', '今月の就寝', '先月の起床', '先月の就寝', '変化(起床)', '累計起床', '累計日数']
        sheet.update('A2', [headers])
        
        rows = []
        for index, user in enumerate(analysis_data):
            rows.append([
                index + 1,
                user['userName'],
                seconds_to_time_str(user['current_wake_avg']),
                seconds_to_time_str(user['current_sleep_avg'], is_sleep=True),
                seconds_to_time_str(user['previous_wake_avg']),
                seconds_to_time_str(user['previous_sleep_avg'], is_sleep=True),
                format_delta_seconds(user['delta_sec']),
                seconds_to_time_str(user['overall_wake_avg']),
                user['overall_count']
            ])

        if rows: sheet.update('A3', rows)

        requests = [
            { "updateSheetProperties": { "properties": { "sheetId": sheet.id, "gridProperties": { "frozenRowCount": 2 } }, "fields": "gridProperties.frozenRowCount" } },
            { "mergeCells": { "range": { "sheetId": sheet.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 9 }, "mergeType": "MERGE_ALL" } },
            { "repeatCell": { "range": { "sheetId": sheet.id, "startRowIndex": 0, "endRowIndex": 1 }, "cell": { "userEnteredFormat": { "textFormat": { "bold": True, "fontSize": 12 }, "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE" } }, "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment)" } },
            { "repeatCell": { "range": { "sheetId": sheet.id, "startRowIndex": 1, "endRowIndex": 2 }, "cell": { "userEnteredFormat": { "backgroundColor": { "red": 0.2, "green": 0.2, "blue": 0.2 }, "textFormat": { "foregroundColor": { "red": 1, "green": 1, "blue": 1 }, "bold": True }, "horizontalAlignment": "CENTER" } }, "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)" } },
            { "repeatCell": { "range": { "sheetId": sheet.id, "startRowIndex": 2 }, "cell": { "userEnteredFormat": { "verticalAlignment": "MIDDLE" } }, "fields": "userEnteredFormat.verticalAlignment" } },
            { "repeatCell": { "range": { "sheetId": sheet.id, "startRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 1 }, "cell": { "userEnteredFormat": { "horizontalAlignment": "CENTER" } }, "fields": "userEnteredFormat.horizontalAlignment" } },
            { "repeatCell": { "range": { "sheetId": sheet.id, "startRowIndex": 2, "startColumnIndex": 2, "endColumnIndex": 9 }, "cell": { "userEnteredFormat": { "horizontalAlignment": "CENTER" } }, "fields": "userEnteredFormat.horizontalAlignment" } },
            { "updateDimensionProperties": { "range": { "sheetId": sheet.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1 }, "properties": { "pixelSize": 40 }, "fields": "pixelSize" } },
            { "updateDimensionProperties": { "range": { "sheetId": sheet.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2 }, "properties": { "pixelSize": 120 }, "fields": "pixelSize" } },
            { "updateDimensionProperties": { "range": { "sheetId": sheet.id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 9 }, "properties": { "pixelSize": 75 }, "fields": "pixelSize" } },
        ]
        sheet.spreadsheet.batch_update({"requests": requests})
    except Exception as e:
        print(f"error: {e}")

def update_monthly_average_sheet(user_daily_data):
    try:
        spreadsheet = get_spreadsheet()
        sheet = get_or_create_worksheet(spreadsheet, "月別平均推移")
        if not user_daily_data: return
        sheet.clear()

        all_year_months = sorted(list({k[:7].replace('-','/') for daily in user_daily_data.values() for k in daily.keys()}))
        if not all_year_months: return

        headers = ['ユーザー名']
        for ym in all_year_months:
            headers.extend([f"{ym} 起床", f"{ym} 就寝"])
        
        rows = []
        for user in sorted(user_daily_data.keys()):
            row = [user]
            for ym in all_year_months:
                w_secs, s_secs = [], []
                for d_str, times in user_daily_data[user].items():
                    if d_str.startswith(ym.replace('/','-')):
                        if 'wake' in times: w_secs.append(time_to_seconds(times['wake']))
                        if 'sleep' in times: s_secs.append(time_to_seconds(times['sleep'], is_sleep=True))
                
                row.append(seconds_to_time_str(sum(w_secs)/len(w_secs)) if w_secs else "--:--")
                row.append(seconds_to_time_str(sum(s_secs)/len(s_secs), is_sleep=True) if s_secs else "--:--")
            rows.append(row)

        sheet.update('A1', [headers])
        if rows: sheet.update('A2', rows)
        
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
            analysis_data, user_daily_data, missing_users, error = await perform_analysis()
            if error: return
            update_spreadsheet(analysis_data)
            update_monthly_average_sheet(user_daily_data)
            
            if missing_users:
                channel = bot.get_channel(TARGET_CHANNEL_ID)
                msg = "🌅 おはようございます！ 本日の起床記録がまだのようです。投稿をお願いします！\n" + " ".join(missing_users)
                await channel.send(msg)
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
        analysis_data, user_daily_data, missing_users, error = await perform_analysis()
        if error:
            await ctx.send(f"エラー: {error}")
            return
        update_spreadsheet(analysis_data)
        update_monthly_average_sheet(user_daily_data)
        sheet_url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
        await ctx.send(f"分析が完了しました！\n結果はこちら: {sheet_url}")
        
        # if missing_users:
        #     msg = "🌅 本日の起床記録がまだのようです！忘れずに投稿をお願いします。\n" + " ".join(missing_users)
        #     await ctx.send(msg)
            
    except Exception as e:
        await ctx.send(f"エラーが発生しました: {e}")

keep_alive()
bot.run(BOT_TOKEN)
