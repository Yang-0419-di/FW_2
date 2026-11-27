# 1. Python 標準庫 (Standard Library)
import base64
from datetime import datetime, timezone, timedelta

import io
import json
from io import BytesIO
import math
import os
import re
import sqlite3
import time

# 2. 第三方套件 (Third-Party Packages)
from flask import (
    Flask, 
    render_template, 
    request, 
    jsonify, 
    abort, 
    redirect, 
    url_for, 
    session
)
from flask_login import LoginManager, login_required, current_user
import gspread
from google.oauth2.service_account import Credentials
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import pandas as pd
import requests

# 3. 本地自訂模組 (Local Application Imports)
from modules.gsheet import client, SHEET_ID
from modules.billing import billing_bp, User
from modules import staff as STAFF_CFG
from modules import calendar_people
from modules import personal_calendar
from modules import line_api, line_bind

# 取得當前 app.py 所在的目錄，並指到 billing.db
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, 'billing.db')

# ====== 基本設定 ======
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'your_default_secret_key_here')  # Flask 密鑰設定
# 設定 Session 存活時間為 10 分鐘
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(minutes=10)
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
if os.environ.get('SECRET_KEY') is None:
    print('⚠️ 未設定環境變數 SECRET_KEY，正在使用不安全的預設值，請盡快設定！')
# 可檢視日誌的管理員（逗號分隔），預設維持原本的 yang.di
ADMIN_USERS = {u.strip() for u in os.environ.get('ADMIN_USERS', 'yang.di').split(',') if u.strip()}

GITHUB_XLSX_URL = 'https://raw.githubusercontent.com/Yang-0419-di/FW_2/master/data.xlsx'
cached_xls = None
cached_at = 0
EXCEL_TTL = int(os.environ.get('EXCEL_TTL', '300'))  # 秒
version_time = None
app.config['VERSION_TIME'] = version_time

# ====== 初始化 Flask-Login ======
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'billing.login'  # 未登入時自動重定向之登入頁面路由

_API_PREFIXES = ('/api/', '/sc_disk/', '/calendar/events', '/calendar/assign', '/personal_calendar/', '/billing/get_last_counts', '/billing/save_person_field')

@login_manager.unauthorized_handler
def _unauthorized():
    wants_json = (
        request.path.startswith(_API_PREFIXES)
        or request.is_json
        or request.headers.get('X-Requested-With', '').lower() == 'fetch'
        or request.accept_mimetypes.best == 'application/json'
    )
    if wants_json:
        return jsonify({'status': 'error', 'message': '登入已逾時，請重新登入',
                        'login_url': url_for('billing.login')}), 401
    return redirect(url_for('billing.login', next=request.full_path.rstrip('?')))

@login_manager.user_loader
def load_user(user_id):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('SELECT id, username FROM users WHERE id = ?', (user_id,))
    user_row = cursor.fetchone()
    conn.close()
    
    if user_row:
        return User(id=user_row[0], username=user_row[1])
    return None

# ====== 註冊 billing 藍圖 ======
app.register_blueprint(billing_bp)

# ====== 人員設定提供給所有模板（modules/staff.py 為唯一來源）======
@app.context_processor
def inject_staff():
    return {'STAFF': STAFF_CFG.STAFF, 'STAFF_DEFAULT_COLOR': STAFF_CFG.DEFAULT_COLOR, 'STAFF_PAST_COLOR': STAFF_CFG.PAST_COLOR}

# ====== 字型設定（支援中文） ======
font_path = os.path.join(BASE_DIR, "fonts", "NotoSansCJKtc-Regular.otf")
if os.path.exists(font_path):
    from matplotlib import font_manager
    font_manager.fontManager.addfont(font_path)
    font_prop = FontProperties(fname=font_path)
    matplotlib.rcParams['font.sans-serif'] = [font_prop.get_name(), 'Microsoft JhengHei']
else:
    font_prop = None
    matplotlib.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
matplotlib.rcParams['axes.unicode_minus'] = False

# ====== 載入 Excel（含版本號） ======
def load_excel_from_github(url):
    global cached_xls, version_time, cached_at
    if cached_xls and (time.time() - cached_at) < EXCEL_TTL:
        return cached_xls
    try:
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            excel_bytes = BytesIO(response.content)
            cached_xls = pd.ExcelFile(excel_bytes, engine='openpyxl')
            cached_at = time.time()
            df_version = pd.read_excel(cached_xls, sheet_name='首頁', header=None, usecols="G", nrows=1)
            version_time = str(df_version.iat[0, 0]) if not pd.isna(df_version.iat[0, 0]) else "無版本資訊"
            app.config['VERSION_TIME'] = version_time
            return cached_xls
    except Exception as e:
        print(f"❌ Excel 下載失敗: {e}")
    if cached_xls:
        return cached_xls  # 下載失敗時退回舊快取
    abort(500, description="⚠️ 無法從 GitHub 載入 Excel 檔案")

def clean_df(df):
    df.columns = df.columns.astype(str).str.replace('\n', '', regex=False)
    return df.fillna('')

# ====== 首頁路由 ======
@app.route('/')
@login_required
def home():
    xls = load_excel_from_github(GITHUB_XLSX_URL)

    # 首頁其他表格處理
    df_department = clean_df(pd.read_excel(xls, sheet_name='首頁', usecols="A:E", skiprows=4, nrows=1))
    df_seasons = clean_df(pd.read_excel(xls, sheet_name='首頁', usecols="A:D", skiprows=8, nrows=2))
    df_project1 = clean_df(pd.read_excel(xls, sheet_name='首頁', usecols="A:E", skiprows=12, nrows=5))

    # HUB 前段統計
    df_HUB_top_raw = pd.read_excel(
        xls, sheet_name='首頁', header=None, usecols="A:C", skiprows=20, nrows=2
    )
    df_HUB_top_raw.columns = df_HUB_top_raw.iloc[0].str.strip()
    df_HUB_top = df_HUB_top_raw[1:]

    cols = ['HUB檢查', 'HUB完工', 'HUB進度']
    existing_cols = [c for c in cols if c in df_HUB_top.columns]
    df_HUB_top = df_HUB_top[existing_cols]

    if 'HUB進度' in df_HUB_top.columns:
        df_HUB_top['HUB進度'] = (
            pd.to_numeric(df_HUB_top['HUB進度'], errors='coerce')
            .fillna(0)
            .mul(100)
            .round(0)
            .astype(int)
            .astype(str) + '%'
        )

    df_HUB = clean_df(pd.read_excel(xls, sheet_name='首頁', header=22, usecols="A:E"))
    df_HUB = df_HUB[df_HUB['門市編號'].astype(str).str.strip() != '']
    df_HUB['門市編號'] = df_HUB['門市編號'].astype(str).str.replace(r'\.0$', '', regex=True)
    df_HUB = df_HUB[['門市編號', '門市名稱', 'HUB規格', '異常原因', '完工確認']]

    df = clean_df(pd.read_excel(xls, sheet_name=0, header=20, nrows=500, usecols="A:O"))
    df = df[['門市編號', '門市名稱', 'PMQ_檢核', '專案檢核', 'HUB', '完工檢核']]

    keyword = request.args.get('keyword', '').strip()
    no_data_found = False
    if keyword:
        df = df[df.astype(str).apply(lambda col: col.str.contains(keyword, case=False, regex=False)).any(axis=1)]
        no_data_found = df.empty

    # 區域數量（三段）
    df1 = pd.read_excel(xls, sheet_name='首頁', header=None, usecols="E:K", skiprows=56, nrows=3)
    headers1 = df1.iloc[0].tolist()
    area_table_1 = [dict(zip(headers1, df1.iloc[i].tolist())) for i in range(1, 3)]

    df2 = pd.read_excel(xls, sheet_name='首頁', header=None, usecols="E:P", skiprows=60, nrows=3)
    headers2 = df2.iloc[0].tolist()
    area_table_2 = [dict(zip(headers2, df2.iloc[i].tolist())) for i in range(1, 3)]

    df3 = pd.read_excel(xls, sheet_name='首頁', header=None, usecols="E:L", skiprows=64, nrows=3)
    headers3 = df3.iloc[0].tolist()
    area_table_3 = [dict(zip(headers3, df3.iloc[i].tolist())) for i in range(1, 3)]

    return render_template(
        'home.html',
        version=version_time,
        area_table_1=area_table_1,
        area_table_2=area_table_2,
        area_table_3=area_table_3,
        keyword=keyword,
        tables=df.to_dict(orient='records'),
        department_table=df_department.to_dict(orient='records'),
        seasons_table=df_seasons.to_dict(orient='records'),
        project1_table=df_project1.to_dict(orient='records'),
        HUB_top_table=df_HUB_top.to_dict(orient='records'),
        HUB_table=df_HUB.to_dict(orient='records'),
        no_data_found=no_data_found,
        billing_invoice_log=False,
        home_page=True
    )

# 2. 新增獨立 API 路由：專門提供 SC 硬碟檢測資料
@app.route('/api/sc_disk_data')
@login_required
def get_sc_disk_data():
    xls = load_excel_from_github(GITHUB_XLSX_URL)
    sc_disk_data = []
    try:
        df_im = clean_df(pd.read_excel(xls, sheet_name='IM'))
        
        cond_category = df_im['報修類別'].astype(str).isin(['HL-TM主機', 'HL-SC主機'])
        content_col = '工作內容' if '工作內容' in df_im.columns else df_im.columns[28]
        cond_content = (
            df_im[content_col].astype(str).str.contains('更換', case=False) & 
            df_im[content_col].astype(str).str.contains('硬碟', case=False)
        )

        df_filtered = df_im[cond_category & cond_content].copy()

        if '離場時間' in df_filtered.columns:
            df_filtered['離場時間_dt'] = pd.to_datetime(df_filtered['離場時間'], errors='coerce')
            today = pd.Timestamp.now().normalize()
            start_date = today - pd.Timedelta(days=5)
            end_date = today + pd.Timedelta(days=6) - pd.Timedelta(seconds=1)
            cond_5days = df_filtered['離場時間_dt'].between(start_date, end_date)
            df_filtered = df_filtered[cond_5days].sort_values(by='離場時間_dt', ascending=False)

        try:
            sh = client.open_by_key(SHEET_ID)
            ws_disk = sh.worksheet("硬碟檢測")
            gs_df = pd.DataFrame(ws_disk.get_all_records())
        except Exception:
            gs_df = pd.DataFrame()

        for _, row in df_filtered.iterrows():
            store_id = str(row.get('門店編號', '')).strip()
            # 取得台芝工作案號作為唯一 Key（clean_df 會將 \n 濾掉，所以抓 '台芝工作案號'）
            work_no = str(row.get('台芝工作案號', '')).strip()

            matched = pd.DataFrame()
            if not gs_df.empty and '台芝工作案號' in gs_df.columns:
                matched = gs_df[gs_df['台芝工作案號'].astype(str).str.strip() == work_no]

            if not matched.empty and 'SC(1)' in matched.columns:
                if str(matched.iloc[0]['SC(1)']).strip() == 'DELETED':
                    continue

            sc_disk_data.append({
                '離場時間': str(row.get('離場時間', '')),
                '門店編號': store_id,
                '門店名稱': str(row.get('門店名稱', '')),
                '報修類別': str(row.get('報修類別', '')),
                '工作內容': str(row.get(content_col, '')),
                '台芝工作案號': work_no,  # 👈 確保有加上這一行
                'SC1': matched.iloc[0]['SC(1)'] if not matched.empty and 'SC(1)' in matched.columns else '',
                'SC2': matched.iloc[0]['SC(2)'] if not matched.empty and 'SC(2)' in matched.columns else '',
                'TM1': matched.iloc[0]['TM(1)'] if not matched.empty and 'TM(1)' in matched.columns else '',
                'TM2': matched.iloc[0]['TM(2)'] if not matched.empty and 'TM(2)' in matched.columns else ''
            })

            if len(sc_disk_data) >= 5:
                break

    except Exception as e:
        print(f"⚠️ SC硬碟檢測載入失敗: {e}")

    return jsonify(sc_disk_data)

def _to_int(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return 0

DISK_USERS = {u.strip() for u in os.environ.get('DISK_USERS', '蔡欣璇,狄澤洋,湯家瑋,劉柏均,吳宗鴻').split(',') if u.strip()}

@app.route("/disk", methods=["GET"])
@login_required
def disk_page():
    try:
        sh = client.open_by_key(SHEET_ID)
        sheet = sh.worksheet("硬碟統計")
    except gspread.exceptions.APIError as e:
        return f"⚠️ 無法讀取 Google Sheet: {e}", 500

    all_rows = sheet.get_all_records()

    latest_data = {}
    for row in all_rows:
        user = row.get('user')
        if user:
            latest_data[user] = row

    rows = list(latest_data.values())

    total_keys = [
        'sc_128_new','sc_128_old','sc_240_new','sc_240_old',
        'sc_256_new','sc_256_old','sc_500_new','sc_500_old',
        'sc_1t_new','sc_1t_old','tm_128_new','tm_128_old','tm_256_new','tm_256_old'
    ]
    total = {k: sum(_to_int(r.get(k)) for r in rows) for k in total_keys}

    return render_template("disk.html", page_header="POS 相關", rows=rows, total=total)

@app.route("/disk/save", methods=["POST"])
@login_required
def disk_save():
    data = {
        "user": request.form.get("user"),
        "sc_128_new": request.form.get("sc_128_new") or "0",
        "sc_128_old": request.form.get("sc_128_old") or "0",
        "sc_240_new": request.form.get("sc_240_new") or "0",
        "sc_240_old": request.form.get("sc_240_old") or "0",
        "sc_256_new": request.form.get("sc_256_new") or "0",
        "sc_256_old": request.form.get("sc_256_old") or "0",
        "sc_500_new": request.form.get("sc_500_new") or "0",
        "sc_500_old": request.form.get("sc_500_old") or "0",
        "sc_1t_new": request.form.get("sc_1t_new") or "0",
        "sc_1t_old": request.form.get("sc_1t_old") or "0",
        "tm_128_new": request.form.get("tm_128_new") or "0",
        "tm_128_old": request.form.get("tm_128_old") or "0",
        "tm_256_new": request.form.get("tm_256_new") or "0",
        "tm_256_old": request.form.get("tm_256_old") or "0"
    }

    if data['user'] not in DISK_USERS:
        return "⚠️ 必須選擇有效的使用者", 400
    for _k, _v in data.items():
        if _k != 'user' and not (str(_v).isdigit() and len(str(_v)) <= 6):
            return f"⚠️ 欄位 {_k} 必須是 0 以上的整數", 400

    try:
        sh = client.open_by_key(SHEET_ID)
        sheet = sh.worksheet("硬碟統計")
        row = [
            data["user"], data["sc_128_new"], data["sc_128_old"],
            data["sc_240_new"], data["sc_240_old"],
            data["sc_256_new"], data["sc_256_old"],
            data["sc_500_new"], data["sc_500_old"],
            data["sc_1t_new"], data["sc_1t_old"],
            data["tm_128_new"], data["tm_128_old"],
            data["tm_256_new"], data["tm_256_old"]
        ]
        sheet.append_row(row)
    except gspread.exceptions.APIError as e:
        return f"⚠️ 無法寫入 Google Sheet: {e}", 500

    return redirect(url_for('disk_page'))

@app.route('/countpass')
@login_required
def countpass():
    return render_template('countpass.html', 
                           page_header="POS 相關",
                           version=version_time, 
                           home_page=False, 
                           billing_invoice_log=False)

@app.route('/personal/<name>')
@login_required
def personal(name):
    sheet_map = {n: n for n in STAFF_CFG.SHEET_NAMES}
    sheet_name = sheet_map.get(name)
    if not sheet_name:
        return f"找不到 {name} 的分頁", 404

    xls = load_excel_from_github(GITHUB_XLSX_URL)

    df_top = clean_df(pd.read_excel(xls, sheet_name=sheet_name, usecols="A:G", nrows=5))
    df_project = clean_df(pd.read_excel(xls, sheet_name=sheet_name, usecols="H:L", nrows=5))
    df_bottom = clean_df(pd.read_excel(xls, sheet_name=sheet_name, usecols="A:K", skiprows=6))
    df_ads = clean_df(pd.read_excel(xls, sheet_name=sheet_name, usecols="A,B,S", skiprows=6))

    df_area = pd.read_excel(
        xls,
        sheet_name=sheet_name,
        usecols="W:AE",
        nrows=1,
        header=0
    )

    df_area.columns = df_area.columns.map(
        lambda x: "-" if str(x).strip().startswith("-") else str(x)
    )

    df_unfinished = pd.read_excel(
        xls,
        sheet_name="未完工清單",
        usecols="A:K"
    )
    df_unfinished = clean_df(df_unfinished)
    df_unfinished = df_unfinished[
        df_unfinished.iloc[:, 0].astype(str).str.strip() == name
    ]

    keyword = request.args.get('keyword', '').strip()
    no_data_found = False
    if keyword:
        df_bottom = df_bottom[
            df_bottom.apply(lambda r: r.astype(str).str.contains(keyword, case=False, regex=False).any(), axis=1)
        ]
        no_data_found = df_bottom.empty

    return render_template(
        "personal.html",
        personal_page=name,
        show_top=not df_top.empty,
        show_area=not df_area.empty,
        show_project=not df_project.empty,
        show_ads=not df_ads.empty,
        show_unfinished=not df_unfinished.empty,
        tables_unfinished=df_unfinished.to_dict(orient="records"),
        tables_top=df_top.to_dict(orient="records"),
        tables_project=df_project.to_dict(orient="records"),
        tables_bottom=df_bottom.to_dict(orient="records"),
        tables_ads=df_ads.to_dict(orient="records"),
        tables_area=df_area.to_dict(orient="records"),
        version=version_time,
        billing_invoice_log=False,
        home_page=False
    )

REPORT_PER_PAGE = int(os.environ.get('REPORT_PER_PAGE', '50'))

@app.route('/report')
@login_required
def report():
    xls = load_excel_from_github(GITHUB_XLSX_URL)
    df = clean_df(pd.read_excel(xls, sheet_name='IM'))
    df = df[['案件類別', '門店編號', '門店名稱', '報修時間', '報修類別', '報修項目', '報修說明', '設備號碼', '服務人員', '工作內容']]
    keyword = request.args.get('keyword', '').strip()
    store_id = request.args.get('store_id', '').strip()
    repair_item = request.args.get('repair_item', '').strip()
    
    tables = []
    
    pagination = None
    query_args = {k: v for k, v in (('keyword', keyword), ('store_id', store_id), ('repair_item', repair_item)) if v}
    if keyword or store_id or repair_item:
        if keyword:
            df = df[df.astype(str).apply(lambda col: col.str.contains(keyword, case=False, regex=False)).any(axis=1)]
        if store_id:
            df = df[df['門店編號'].astype(str).str.contains(store_id, case=False, regex=False)]
        if repair_item:
            df = df[df['報修類別'].astype(str).str.strip() == repair_item.strip()]
        total = len(df)
        pages = max(1, math.ceil(total / REPORT_PER_PAGE))
        page = min(max(request.args.get('page', 1, type=int) or 1, 1), pages)
        tables = df.iloc[(page - 1) * REPORT_PER_PAGE: page * REPORT_PER_PAGE].to_dict(orient='records')
        pagination = {'page': page, 'pages': pages, 'total': total, 'per_page': REPORT_PER_PAGE}
        
    return render_template(
        'report.html',
        page_header="POS 相關",
        version=version_time,
        tables=tables,
        pagination=pagination,
        query_args=query_args,
        keyword=keyword,
        store_id=store_id,
        repair_item=repair_item,
        no_data_found=(len(tables) == 0 and (keyword or store_id or repair_item)),
        billing_invoice_log=False,
        home_page=False
    )

@app.route('/sm_web')
@login_required
def sm_web_page():
    return render_template(
        'sm_web.html',
        sm_web=True,
        version=version_time,
        billing_invoice_log=False,
        home_page=False
    )

@app.route('/inspection_log')
@login_required
def inspection_log():
    # 限制僅 yang.di 可存取
    if current_user.username not in ADMIN_USERS:
        abort(403)

    target_username = request.args.get('username')
    logs = []
    
    try:
        sh = client.open_by_key(SHEET_ID)
        ws = sh.worksheet("log")
        raw_logs = ws.get_all_records()
        
        # 確保資料依原本邏輯反轉（最新的排在前面）
        logs = list(raw_logs)
        logs.reverse()  
    except Exception as e:
        print(f"⚠️ 檢視日誌載入失敗: {e}")

    if target_username:
        # === 模式二：特定帳號詳細紀錄與不重複 IP ===
        user_logs = [log for log in logs if str(log.get('username')) == str(target_username)]
        
        # 萃取該帳號使用過的所有 IP 並去重 (保持不重複)
        unique_ips = []
        for log in user_logs:
            ip = log.get('ip_address')
            if ip and ip not in unique_ips:
                unique_ips.append(ip)
                
        return render_template(
            'inspection_log.html',
            page_header="檢視日誌",
            version=version_time,
            view_mode='detail',
            target_username=target_username,
            logs=user_logs,
            unique_ips=unique_ips,
            billing_invoice_log=False,
            home_page=False
        )
    else:
        # === 模式一：帳號總覽首頁（只取每個帳號最後一次登入） ===
        summary_dict = {}
        for log in logs:
            uname = log.get('username')
            if uname and uname not in summary_dict:
                summary_dict[uname] = {
                    'username': uname,
                    'last_ip': log.get('ip_address'),
                    'last_date': log.get('login_date'),
                    'last_time': log.get('login_time')
                }
        
        summary_list = list(summary_dict.values())
        
        return render_template(
            'inspection_log.html',
            page_header="檢視日誌",
            version=version_time,
            view_mode='summary',
            summary_list=summary_list,
            billing_invoice_log=False,
            home_page=False
        )


# save_log v3
@app.route('/api/ping', methods=['POST'])
@login_required
def api_ping():
    """前端使用中定期呼叫，讓後端 session 延長有效期（避免輸入到一半被登出）"""
    session.modified = True
    return jsonify({'status': 'ok'})


@app.route('/api/save_log', methods=['POST'])
@login_required
def save_log():
    """記錄裝置指紋。
    - 帳號一律取自登入者，不信任前端傳入
    - 優先補在該帳號『剛登入』那一筆（避免一次登入出現兩列）
    - 相同指紋 30 分鐘內不重複寫入
    """
    data = request.get_json(silent=True) or {}
    username = current_user.username
    fingerprint = str(data.get('fingerprint') or '').strip()[:200]
    if not fingerprint:
        return jsonify({'status': 'error', 'message': '缺少 fingerprint'}), 400

    ip = (request.headers.get('X-Forwarded-For') or request.remote_addr or '').split(',')[0].strip()
    taiwan_tz = timezone(timedelta(hours=8))
    now = datetime.now(taiwan_tz)

    try:
        sh = client.open_by_key(SHEET_ID)
        ws = sh.worksheet("log")
        headers = ws.row_values(1)

        def _col(name, default):
            return headers.index(name) if name in headers else default

        i_user, i_created, i_fp = _col('username', 1), _col('created_at', 5), _col('fingerprint', 6)

        usernames = ws.col_values(i_user + 1)  # 含標題列；list 索引 i 對應試算表第 i+1 列
        last_row = None
        for i in range(len(usernames) - 1, 0, -1):
            if usernames[i] == username:
                last_row = i + 1
                break

        if last_row is not None:
            row = ws.row_values(last_row)
            last_fp = row[i_fp] if len(row) > i_fp else ''
            try:
                last_at = datetime.strptime(row[i_created], '%Y-%m-%d %H:%M:%S')
                age = (now.replace(tzinfo=None) - last_at).total_seconds()
            except (ValueError, IndexError):
                age = None

            if age is not None and age < 30 * 60:
                if last_fp == fingerprint:
                    return jsonify({'status': 'success', 'message': '指紋已記錄（略過重複）'})
                if not last_fp:
                    ws.update_cell(last_row, i_fp + 1, fingerprint)
                    return jsonify({'status': 'success', 'message': '已補上本次登入的裝置指紋'})

        new_id = len(usernames)  # 標題列 + 既有資料列數 = 新 ID
        ws.append_row([
            new_id, username, ip,
            now.strftime('%Y-%m-%d'), now.strftime('%H:%M:%S'), now.strftime('%Y-%m-%d %H:%M:%S'),
            fingerprint
        ])
        return jsonify({'status': 'success', 'message': '日誌與指紋已記錄'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/time')
@login_required
def time_page():
    xls = pd.ExcelFile(os.path.join(BASE_DIR, "MFP", "MFP.xlsx"))

    df_summary = pd.read_excel(
        xls,
        sheet_name='出勤時間',
        usecols="A:F",
        header=0,
        nrows=1
    )

    detail_1 = pd.read_excel(
        xls,
        sheet_name='出勤時間',
        usecols="A:Q",
        header=3,
        nrows=4
    )

    detail_2 = pd.read_excel(
        xls,
        sheet_name='出勤時間',
        usecols="A:Q",
        header=8,
        nrows=4
    )

    detail_3 = pd.read_excel(
        xls,
        sheet_name='出勤時間',
        usecols="A:Q",
        header=13,
        nrows=4
    )

    df_chart = pd.read_excel(xls, sheet_name='出勤時間', header=None)

    x = [str(v) for v in df_chart.iloc[13, 1:16].tolist()]
    names = df_chart.iloc[14:18, 0].tolist()
    y_data = df_chart.iloc[14:18, 1:16].values.tolist()

    fig, ax = plt.subplots(figsize=(10, 5))

    for i, y in enumerate(y_data):
        ax.plot(x, y, marker='o', label=names[i])

    ax.set_xlabel('日期')
    ax.set_ylabel('時數')
    ax.legend()
    plt.xticks(rotation=45)
    plt.tight_layout()

    img = io.BytesIO()
    plt.savefig(img, format='png')
    img.seek(0)
    plot_url = base64.b64encode(img.read()).decode('utf-8')
    plt.close()

    return render_template(
        'time.html',
        version=version_time,
        summary_table=df_summary.to_html(index=False, classes='dataframe'),
        detail_table_1=detail_1.to_html(index=False, classes='dataframe'),
        detail_table_2=detail_2.to_html(index=False, classes='dataframe'),
        detail_table_3=detail_3.to_html(index=False, classes='dataframe'),
        plot_url=plot_url,
        df_summary=df_summary,
        time_page=True,
        billing_invoice_log=False,
        home_page=False
    )

@app.route('/mfp_parts', methods=['GET', 'POST'])
@login_required
def mfp_parts():
    xls = load_excel_from_github(GITHUB_XLSX_URL)
    df = pd.read_excel(xls, sheet_name='MFP_零件表')
    model = request.form.get('model', '')
    part = request.form.get('part', '')
    message = ""
    table_html = ""
    if request.method == 'POST':
        if not model:
            message = "⚠️ 請選擇機型"
        else:
            filtered = df[df['機型'] == model]
            if part:
                filtered = filtered[filtered['部件'] == part]
            if filtered.empty:
                message = "查無資料"
            else:
                table_html = filtered[['零件名稱', '部件', '料號', '型號', '壽命(張)']].to_html(classes="data-table", index=False, border=0)
                
    return render_template(
        'mfp_parts.html',
        page_header="MFP 相關",
        version=version_time,
        message=message,
        table_html=table_html,
        selected_model=model,
        selected_part=part,
        billing_invoice_log=False,
        home_page=False
    )

@app.route('/calendar')
@login_required
def calendar_page():
    return render_template(
        'calendar.html',
        version=version_time,
        calendar_page=True
    )

TW_TZ = timezone(timedelta(hours=8))
CALENDAR_REQUIRED_COLS = ('date', 'title')


def _calendar_rows():
    """讀 data.xlsx「行事曆」分頁 → (事件清單, 警告清單)。讀不到時丟出 RuntimeError（訊息給使用者看）。"""
    try:
        xls = load_excel_from_github(GITHUB_XLSX_URL)
        df = pd.read_excel(xls, sheet_name='行事曆')
    except Exception as e:
        print(f"⚠️ 行事曆載入失敗: {e}")
        raise RuntimeError("無法讀取排程資料（data.xlsx 下載或「行事曆」分頁讀取失敗），請稍後再試或通知管理員。")
    df.columns = df.columns.astype(str).str.strip()
    missing = [c for c in CALENDAR_REQUIRED_COLS if c not in df.columns]
    if missing:
        raise RuntimeError(f"「行事曆」分頁缺少欄位：{'、'.join(missing)}（請確認第 1 列標題沒有被改名或移動）。")

    rows, no_title = [], 0
    for _, row in df.iterrows():
        date_val = row.get('date')
        if pd.isna(date_val) or str(date_val).strip() == '':
            continue
        title_val = row.get('title')
        if pd.isna(title_val) or str(title_val).strip() == '':
            no_title += 1          # 有日期卻沒標題：多半是公式沒有計算結果
            continue
        try:
            start_date = pd.to_datetime(date_val).date()
        except Exception:
            continue
        owner_raw = row.get('屬性')
        owner_raw = '' if pd.isna(owner_raw) else str(owner_raw).strip()
        special = row.get('特殊')
        special = '' if pd.isna(special) else str(special)
        rows.append({
            'date': start_date,
            'title': str(title_val).strip(),
            'owner_raw': owner_raw,
            'is_special': special.strip().upper() in ['Y', 'YES', '1'],
        })
    warnings = []
    if no_title:
        warnings.append(f"有 {no_title} 筆排程有日期但沒有標題，未顯示（data.xlsx 的標題公式可能沒有計算結果，請用 PM 工具「儲存並重算」後再發佈）。")
    return rows, warnings


@app.route('/calendar/events')
@login_required
def calendar_events():
    try:
        rows, warnings = _calendar_rows()
    except RuntimeError as e:
        return jsonify({'status': 'error', 'message': str(e)}), 503

    fresh = request.args.get('fresh') == '1'
    try:
        people = calendar_people.load_people(fresh=fresh)
    except Exception as e:
        print(f"⚠️ 行事曆人員載入失敗: {e}")
        people = {}
        warnings.append("無法讀取「人員2」資料（Google Sheet），目前只顯示預設人員。")

    today = datetime.now(TW_TZ).date()      # Render 主機是 UTC，改用台灣日期判斷是否過期
    events = []
    for r in rows:
        date_str = r['date'].strftime('%Y-%m-%d')
        key = calendar_people.make_key(date_str, r['title'])
        owner = STAFF_CFG.display_name(r['owner_raw'])
        person2 = people.get(key, '')
        color = STAFF_CFG.PAST_COLOR if r['date'] < today else STAFF_CFG.color_of(r['owner_raw'])
        names = ' & '.join(n for n in (owner, person2) if n)
        events.append({
            'title': r['title'] + (f"\n人員：{names}" if person2 else ''),
            'start': date_str,
            'color': color,
            'extendedProps': {
                'is_special': r['is_special'],
                'base_title': r['title'],
                'owner': owner,
                'person2': person2,
            },
        })
    return jsonify({'status': 'ok', 'events': events, 'warnings': warnings})


PERSON2_MAX = 30


@app.route('/calendar/assign', methods=['POST'])
@login_required
def calendar_assign():
    d = request.get_json(silent=True) or {}
    date_str = str(d.get('date', '')).strip()[:10]
    title = str(d.get('title', '')).strip()
    person2 = re.sub(r'\s+', ' ', str(d.get('person2', ''))).strip()

    try:
        datetime.strptime(date_str, '%Y-%m-%d')
    except ValueError:
        return jsonify({'status': 'error', 'message': '日期格式錯誤'}), 400
    if not title:
        return jsonify({'status': 'error', 'message': '缺少工務標題'}), 400
    if len(person2) > PERSON2_MAX:
        return jsonify({'status': 'error', 'message': f'人員2 最多 {PERSON2_MAX} 個字'}), 400
    if re.search(r'[<>"\\=]', person2):
        return jsonify({'status': 'error', 'message': '人員2 不可包含 < > " \\ = 等符號'}), 400

    # 只接受 data.xlsx 中確實存在的工務，避免寫入無對應的資料
    try:
        rows, _ = _calendar_rows()
    except RuntimeError as e:
        return jsonify({'status': 'error', 'message': str(e)}), 503
    key = calendar_people.make_key(date_str, title)
    if not any(calendar_people.make_key(r['date'].strftime('%Y-%m-%d'), r['title']) == key for r in rows):
        return jsonify({'status': 'error', 'message': '找不到這筆工務（排程可能剛被更新），請重新整理頁面後再試。'}), 409

    try:
        calendar_people.save_person(date_str, title, person2, getattr(current_user, 'username', ''))
    except Exception as e:
        print(f"⚠️ 行事曆人員寫入失敗: {e}")
        return jsonify({'status': 'error', 'message': '寫入 Google Sheet 失敗，請稍後再試。'}), 502
    return jsonify({'status': 'ok', 'person2': person2})

# ====== 個人行事曆（Google Sheet「個人行事曆」分頁）======
_store_cache = {'at': None, 'list': []}


def _store_list():
    """首頁門市資料的 門市編號／門市名稱（跟著 data.xlsx 快取更新）。"""
    xls = load_excel_from_github(GITHUB_XLSX_URL)
    if _store_cache['at'] == cached_at and _store_cache['list']:
        return _store_cache['list']
    out = []
    try:
        df = clean_df(pd.read_excel(xls, sheet_name=0, header=20, nrows=500, usecols="A:O"))
        seen = set()
        for _, r in df.iterrows():
            sid = re.sub(r'\.0$', '', str(r.get('門市編號', '')).strip())
            name = str(r.get('門市名稱', '')).strip()
            if sid and name and sid not in seen:
                seen.add(sid)
                out.append({'id': sid, 'name': name})
    except Exception as e:
        print(f"⚠️ 門市清單讀取失敗: {e}")
    _store_cache.update(at=cached_at, list=out)
    return out


@app.route('/personal_calendar/stores')
@login_required
def personal_calendar_stores():
    return jsonify({'status': 'ok', 'stores': _store_list()})


def _pc_title(e):
    """事件標題：店編_店名_內容_時間（空白的欄位略過），與部門排程表的格式一致。"""
    return '_'.join(x for x in (e['store'], e['name'], e['content'], e['time']) if x)


PC_LIMITS = {'store': 20, 'name': 40, 'content': 60, 'time': 20}


def _pc_clean(d):
    """檢查並整理表單資料；有問題丟出 ValueError（訊息給使用者看）。"""
    data = {}
    date_str = str(d.get('date', '')).strip()[:10]
    try:
        datetime.strptime(date_str, '%Y-%m-%d')
    except ValueError:
        raise ValueError('請選擇日期')
    data['date'] = date_str
    person = str(d.get('person', '')).strip()
    if person not in STAFF_CFG.STAFF_NAMES:
        raise ValueError('請選擇工程師')
    data['person'] = person
    for k, n in PC_LIMITS.items():
        v = re.sub(r'\s+', ' ', str(d.get(k, ''))).strip()
        if len(v) > n:
            raise ValueError(f'{ {"store": "店編", "name": "店名", "content": "內容", "time": "時間"}[k] } 最多 {n} 個字')
        data[k] = v
    sp = str(d.get('special', '')).strip().upper()
    if sp not in ('', 'Y'):
        raise ValueError('特殊欄只能填 Y 或留白')
    data['special'] = sp
    if not (data['store'] or data['content']):
        raise ValueError('店編和內容至少要填一項')
    # 提前通知：空白＝不通知、0＝當日、1~60＝提前幾天
    nt = str(d.get('notify', '')).strip()
    if nt != '':
        if not nt.isdigit() or int(nt) > 60:
            raise ValueError('提前通知天數請填 0～60 的整數（0＝當日）')
        nt = str(int(nt))
    data['notify'] = nt
    return data


@app.route('/personal_calendar/<name>/events')
@login_required
def personal_calendar_events(name):
    if name not in STAFF_CFG.STAFF_NAMES:
        return jsonify({'status': 'error', 'message': f'找不到人員：{name}'}), 404
    try:
        rows = personal_calendar.list_events(name, fresh=request.args.get('fresh') == '1')
    except Exception as e:
        print(f"⚠️ 個人行事曆讀取失敗: {e}")
        return jsonify({'status': 'error', 'message': '無法讀取個人行事曆（Google Sheet），請稍後再試。'}), 503
    today = datetime.now(TW_TZ).date()
    events = []
    for e in rows:
        try:
            d = datetime.strptime(e['date'], '%Y-%m-%d').date()
        except ValueError:
            continue
        events.append({
            'id': e['id'],
            'title': _pc_title(e) or '（未命名）',
            'start': e['date'],
            'color': STAFF_CFG.PAST_COLOR if d < today else STAFF_CFG.color_of(e['person']),
            'extendedProps': {k: e[k] for k in ('person', 'store', 'name', 'content', 'time', 'special', 'notify', 'notified',
                                                'created_by', 'created_at', 'updated_by', 'updated_at')},
        })
    return jsonify({'status': 'ok', 'events': events})


@app.route('/personal_calendar/save', methods=['POST'])
@login_required
def personal_calendar_save():
    d = request.get_json(silent=True) or {}
    try:
        data = _pc_clean(d)
    except ValueError as e:
        return jsonify({'status': 'error', 'message': str(e)}), 400
    user = getattr(current_user, 'username', '')
    event_id = str(d.get('id', '')).strip()
    try:
        if event_id:
            personal_calendar.update_event(event_id, data, user)
        else:
            event_id = personal_calendar.add_event(data, user)
    except LookupError as e:
        return jsonify({'status': 'error', 'message': str(e)}), 404
    except Exception as e:
        print(f"⚠️ 個人行事曆寫入失敗: {e}")
        return jsonify({'status': 'error', 'message': '寫入 Google Sheet 失敗，請稍後再試。'}), 502
    return jsonify({'status': 'ok', 'id': event_id})


@app.route('/personal_calendar/delete', methods=['POST'])
@login_required
def personal_calendar_delete():
    d = request.get_json(silent=True) or {}
    event_id = str(d.get('id', '')).strip()
    if not event_id:
        return jsonify({'status': 'error', 'message': '缺少事件編號'}), 400
    try:
        personal_calendar.delete_event(event_id)
    except LookupError as e:
        return jsonify({'status': 'error', 'message': str(e)}), 404
    except Exception as e:
        print(f"⚠️ 個人行事曆刪除失敗: {e}")
        return jsonify({'status': 'error', 'message': '刪除失敗（Google Sheet），請稍後再試。'}), 502
    return jsonify({'status': 'ok'})


# ====== LINE 通知：綁定（個人頁）與 webhook（LINE 官方帳號）======
def _line_add_friend_url():
    return os.environ.get('LINE_ADD_FRIEND_URL', '').strip()


@app.route('/personal_calendar/<name>/line')
@login_required
def personal_line_status(name):
    if name not in STAFF_CFG.STAFF_NAMES:
        return jsonify({'status': 'error', 'message': f'找不到人員：{name}'}), 404
    try:
        st = line_bind.status(name)
    except Exception as e:
        print(f"⚠️ LINE 綁定狀態讀取失敗: {e}")
        return jsonify({'status': 'error', 'message': '無法讀取 LINE 綁定狀態'}), 503
    return jsonify({'status': 'ok', **st, 'add_friend_url': _line_add_friend_url(),
                    'configured': bool(os.environ.get('LINE_CHANNEL_SECRET') and os.environ.get('LINE_CHANNEL_ACCESS_TOKEN'))})


@app.route('/personal_calendar/<name>/line/code', methods=['POST'])
@login_required
def personal_line_code(name):
    if name not in STAFF_CFG.STAFF_NAMES:
        return jsonify({'status': 'error', 'message': f'找不到人員：{name}'}), 404
    try:
        code, expire = line_bind.new_code(name, getattr(current_user, 'username', ''))
    except Exception as e:
        print(f"⚠️ LINE 綁定碼產生失敗: {e}")
        return jsonify({'status': 'error', 'message': '產生綁定碼失敗（Google Sheet），請稍後再試。'}), 502
    return jsonify({'status': 'ok', 'code': code, 'expire': expire, 'minutes': line_bind.CODE_MINUTES,
                    'add_friend_url': _line_add_friend_url()})


@app.route('/personal_calendar/<name>/line/unbind', methods=['POST'])
@login_required
def personal_line_unbind(name):
    if name not in STAFF_CFG.STAFF_NAMES:
        return jsonify({'status': 'error', 'message': f'找不到人員：{name}'}), 404
    try:
        line_bind.unbind_person(name, getattr(current_user, 'username', ''))
    except Exception as e:
        print(f"⚠️ LINE 解除綁定失敗: {e}")
        return jsonify({'status': 'error', 'message': '解除綁定失敗（Google Sheet），請稍後再試。'}), 502
    return jsonify({'status': 'ok'})


LINE_HELP = ("這是部門行事曆的 LINE 通知帳號。\n"
             "・綁定：到網站自己的個人頁，在「個人行事曆」按「取得綁定碼」，再傳「綁定 123456」到這裡\n"
             "・查詢目前綁定：傳「狀態」\n"
             "・停止接收：傳「解除綁定」")


def _line_handle_text(text, user_id):
    """處理使用者傳給官方帳號的文字，回傳要回覆的內容。"""
    t = text.strip().translate(str.maketrans('０１２３４５６７８９　', '0123456789 '))
    m = re.fullmatch(r'(?:綁定\s*)?(\d{6})', t)
    if m:
        person = line_bind.bind_by_code(m.group(1), user_id)
        if person:
            return f"✅ 已綁定為「{person}」。\n之後個人行事曆設定的提前通知，會在每天早上 8 點左右傳到這裡。"
        return "❌ 綁定碼錯誤或已過期（有效 10 分鐘）。請到個人頁重新按「取得綁定碼」。"
    if t in ('解除綁定', '取消綁定'):
        persons = line_bind.unbind_user(user_id)
        return f"已解除「{'、'.join(persons)}」的綁定，之後不會再收到提醒。" if persons else "這個 LINE 帳號目前沒有綁定任何人員。"
    if t in ('狀態', '查詢', '綁定狀態'):
        persons = line_bind.persons_of(user_id)
        return f"目前綁定：{'、'.join(persons)}" if persons else "目前沒有綁定。\n\n" + LINE_HELP
    return LINE_HELP


@app.route('/line/webhook', methods=['POST'])
def line_webhook():
    """LINE 官方帳號的 webhook（不需登入，改以 LINE 簽章驗證來源）。"""
    body = request.get_data()
    if not line_api.verify_signature(body, request.headers.get('X-Line-Signature', '')):
        return 'invalid signature', 400
    try:
        events = json.loads(body.decode('utf-8')).get('events', [])
    except Exception:
        return 'bad request', 400
    for ev in events:
        try:
            src = ev.get('source', {})
            if src.get('type') != 'user':
                continue                      # 只處理一對一聊天，群組訊息不回應
            token = ev.get('replyToken')
            if ev.get('type') == 'follow' and token:
                line_api.reply(token, "感謝加入！\n\n" + LINE_HELP)
            elif ev.get('type') == 'message' and ev.get('message', {}).get('type') == 'text' and token:
                line_api.reply(token, _line_handle_text(ev['message'].get('text', ''), src.get('userId', '')))
        except Exception as e:
            print(f"⚠️ LINE webhook 處理失敗: {e}")
    return 'OK'


@app.route("/worktime")
@login_required
def worktime():
    path = os.path.join(BASE_DIR, "MFP", "MFP.xlsx")

    df_1 = pd.read_excel(
        path,
        sheet_name="工時計算",
        header=None,
        usecols="A:F",
        skiprows=1,
        nrows=2
    )
    block1_header = df_1.iloc[0].tolist()
    block1_body = df_1.iloc[1:].values.tolist()

    df_2 = pd.read_excel(
        path,
        sheet_name="工時計算",
        header=None,
        usecols="A:J",
        skiprows=4,
        nrows=5
    )
    block2_header = df_2.iloc[0].tolist()
    block2_body = df_2.iloc[1:-1].values.tolist()
    block2_note = df_2.iloc[-1].tolist()

    df_3 = pd.read_excel(
        path,
        sheet_name="工時計算",
        header=None,
        usecols="A:H",
        skiprows=9,
        nrows=4
    )
    block3_header = df_3.iloc[0].tolist()
    block3_body = df_3.iloc[1:].values.tolist()

    df_4 = pd.read_excel(
        path,
        sheet_name="工時計算",
        header=None,
        usecols="A:K",
        skiprows=13,
        nrows=5
    )
    block4_header = df_4.iloc[0].tolist()
    block4_body = df_4.iloc[1:-1].values.tolist()
    block4_note = df_4.iloc[-1].tolist()

    return render_template(
        "worktime.html",
        block1_header=block1_header, block1_body=block1_body,
        block2_header=block2_header, block2_body=block2_body, block2_note=block2_note,
        block3_header=block3_header, block3_body=block3_body,
        block4_header=block4_header, block4_body=block4_body, block4_note=block4_note,
        billing_worktime=True
    )

# 🆕 SC硬碟檢測 儲存/刪除同步 API
@app.route('/sc_disk/update', methods=['POST'])
@login_required
def update_sc_disk():
    data = request.json
    action = data.get('action')
    store_id = str(data.get('store_id', '')).strip()
    work_no = str(data.get('work_no', '')).strip()

    try:
        sh = client.open_by_key(SHEET_ID)
        ws = sh.worksheet("硬碟檢測")
        records = ws.get_all_records()
        
        # 尋找目標列：優先以「台芝工作案號」比對，若無則以「門店編號」輔助比對
        row_idx = None
        for idx, r in enumerate(records, start=2):
            r_work_no = str(r.get('台芝工作案號', '')).strip()
            r_store_id = str(r.get('門店編號', '')).strip()
            
            if work_no and r_work_no == work_no:
                row_idx = idx
                break
            elif not work_no and store_id and r_store_id == store_id:
                row_idx = idx
                break

        if action == 'delete':
            row_data = [
                work_no,                      # A: 台芝工作案號
                data.get('leave_time', ''),   # B: 離場時間
                store_id,                     # C: 門店編號
                data.get('store_name', ''),   # D: 門店名稱
                data.get('repair_cat', ''),   # E: 報修類別
                data.get('work_content', ''), # F: 工作內容
                'DELETED',                    # G: SC(1) 設為 DELETED
                'DELETED',                    # H: SC(2)
                'DELETED',                    # I: TM(1)
                'DELETED'                     # J: TM(2)
            ]
            if row_idx:
                ws.update(f'A{row_idx}:J{row_idx}', [row_data])
            else:
                ws.append_row(row_data)
                
            return jsonify({'status': 'success', 'message': '已標記為刪除'})

        elif action == 'save':
            row_data = [
                work_no,                      # A: 台芝工作案號
                data.get('leave_time', ''),   # B: 離場時間
                store_id,                     # C: 門店編號
                data.get('store_name', ''),   # D: 門店名稱
                data.get('repair_cat', ''),   # E: 報修類別
                data.get('work_content', ''), # F: 工作內容
                data.get('sc1', ''),          # G: SC(1)
                data.get('sc2', ''),          # H: SC(2)
                data.get('tm1', ''),          # I: TM(1)
                data.get('tm2', '')           # J: TM(2)
            ]
            if row_idx:
                ws.update(f'A{row_idx}:J{row_idx}', [row_data])
            else:
                ws.append_row(row_data)
                
            return jsonify({'status': 'success', 'message': '資料已同步至 Google Sheet'})

        return jsonify({'status': 'error', 'message': '未知的 action'}), 400

    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.context_processor
def inject_admin_users():
    return {'ADMIN_USERS': globals().get('ADMIN_USERS', {'yang.di'})}

@app.before_request
def make_session_permanent():
    # 每次使用者發送請求時，重新刷新 10 分鐘的倒數計時
    session.permanent = True

# ====== 啟動 Flask ======
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
