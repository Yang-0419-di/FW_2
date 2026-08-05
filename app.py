from flask import Flask, render_template, request, jsonify, abort, redirect, url_for
from flask_login import LoginManager, UserMixin
import pandas as pd
import requests
import sqlite3
from io import BytesIO
import os, io, base64
from datetime import datetime
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import gspread
from google.oauth2.service_account import Credentials
from modules.gsheet import client, SHEET_ID

# ====== 引入 billing 模組 ======
from modules.billing import billing_bp

# ====== 基本設定 ======
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'your_default_secret_key_here')  # Flask 密鑰設定

GITHUB_XLSX_URL = 'https://raw.githubusercontent.com/Yang-0419-di/FW_2/master/data.xlsx'
cached_xls = None
version_time = None
app.config['VERSION_TIME'] = version_time

# ====== 初始化 Flask-Login ======
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'billing.login'  # 未登入時自動重定向之登入頁面路由

# ====== User 模型類別 ======
class User(UserMixin):
    def __init__(self, id, username):
        self.id = id
        self.username = username

@login_manager.user_loader
def load_user(user_id):
    conn = sqlite3.connect('billing.db')
    cursor = conn.cursor()
    cursor.execute('SELECT id, username FROM users WHERE id = ?', (user_id,))
    user_row = cursor.fetchone()
    conn.close()
    
    if user_row:
        return User(id=user_row[0], username=user_row[1])
    return None

# ====== 註冊 billing 藍圖 ======
app.register_blueprint(billing_bp)

# ====== 字型設定（支援中文） ======
matplotlib.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
matplotlib.rcParams['axes.unicode_minus'] = False
font_path = "./fonts/NotoSansCJKtc-Regular.otf"
font_prop = FontProperties(fname=font_path)

# ====== 載入 Excel（含版本號） ======
def load_excel_from_github(url):
    global cached_xls, version_time
    if cached_xls:
        return cached_xls
    try:
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            excel_bytes = BytesIO(response.content)
            cached_xls = pd.ExcelFile(excel_bytes, engine='openpyxl')
            df_version = pd.read_excel(cached_xls, sheet_name='首頁', header=None, usecols="G", nrows=1)
            version_time = str(df_version.iat[0, 0]) if not pd.isna(df_version.iat[0, 0]) else "無版本資訊"
            app.config['VERSION_TIME'] = version_time
            return cached_xls
    except Exception as e:
        print(f"❌ Excel 下載失敗: {e}")
    abort(500, description="⚠️ 無發從 GitHub 載入 Excel 檔案")

def clean_df(df):
    df.columns = df.columns.astype(str).str.replace('\n', '', regex=False)
    return df.fillna('')

# ====== 首頁路由 ======
@app.route('/')
def home():
    xls = load_excel_from_github(GITHUB_XLSX_URL)

    # SC硬碟檢測：資料篩選與 Google Sheet 串接
    sc_disk_data = []
    try:
        df_im = clean_df(pd.read_excel(xls, sheet_name='IM'))

        # 1. 條件一：報修類別為 HL-TM主機 或 HL-SC主機
        cond_category = df_im['報修類別'].astype(str).isin(['HL-TM主機', 'HL-SC主機'])

        # 2. 條件二：工作內容 (AC欄位) 包含 "更換" 與 "硬碟"
        content_col = '工作內容' if '工作內容' in df_im.columns else df_im.columns[28]
        cond_content = (
            df_im[content_col].astype(str).str.contains('更換', case=False) & 
            df_im[content_col].astype(str).str.contains('硬碟', case=False)
        )

        df_filtered = df_im[cond_category & cond_content].copy()

        # 3. 依據 離場時間 排序並取最新前 5 筆
        if '離場時間' in df_filtered.columns:
            df_filtered['離場時間_dt'] = pd.to_datetime(df_filtered['離場時間'], errors='coerce')
            df_filtered = df_filtered.sort_values(by='離場時間_dt', ascending=False).head(5)
        else:
            df_filtered = df_filtered.head(5)

        # 4. 從 Google Sheet「硬碟檢測」分頁讀取填寫紀錄
        try:
            sh = client.open_by_key(SHEET_ID)
            ws_disk = sh.worksheet("硬碟檢測")
            gs_df = pd.DataFrame(ws_disk.get_all_records())
        except Exception:
            gs_df = pd.DataFrame()

        for _, row in df_filtered.iterrows():
            store_id = str(row.get('門店編號', ''))
            
            matched = pd.DataFrame()
            if not gs_df.empty and '門店編號' in gs_df.columns:
                matched = gs_df[gs_df['門店編號'].astype(str) == store_id]

            sc_disk_data.append({
                '離場時間': str(row.get('離場時間', '')),
                '門店編號': store_id,
                '門店名稱': str(row.get('門店名稱', '')),
                '報修類別': str(row.get('報修類別', '')),
                '工作內容': str(row.get(content_col, '')),
                'SC1': matched.iloc[0]['SC(1)'] if not matched.empty and 'SC(1)' in matched.columns else '',
                'SC2': matched.iloc[0]['SC(2)'] if not matched.empty and 'SC(2)' in matched.columns else '',
                'TM1': matched.iloc[0]['TM(1)'] if not matched.empty and 'TM(1)' in matched.columns else '',
                'TM2': matched.iloc[0]['TM(2)'] if not matched.empty and 'TM(2)' in matched.columns else ''
            })
    except Exception as e:
        print(f"⚠️ SC硬碟檢測載入失敗: {e}")

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
        df = df[df.apply(lambda r: r.astype(str).str.contains(keyword, case=False).any(), axis=1)]
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
        sc_disk_data=sc_disk_data,
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

@app.route("/disk", methods=["GET"])
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
    total = {k: sum(int(r.get(k) or 0) for r in rows) for k in total_keys}

    return render_template("disk.html", page_header="POS 相關", rows=rows, total=total)

@app.route("/disk/save", methods=["POST"])
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

    if not data['user']:
        return "⚠️ 必須選擇使用者", 400

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
def countpass():
    return render_template('countpass.html', 
                           page_header="POS 相關",
                           version=version_time, 
                           home_page=False, 
                           billing_invoice_log=False)

@app.route('/personal/<name>')
def personal(name):
    sheet_map = {'吳宗鴻': '吳宗鴻', '湯家瑋': '湯家瑋', '狄澤洋': '狄澤洋','劉柏均': '劉柏均'}
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
            df_bottom.apply(lambda r: r.astype(str).str.contains(keyword, case=False).any(), axis=1)
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

@app.route('/report')
def report():
    xls = load_excel_from_github(GITHUB_XLSX_URL)
    df = clean_df(pd.read_excel(xls, sheet_name='IM'))
    df = df[['案件類別', '門店編號', '門店名稱', '報修時間', '報修類別', '報修項目', '報修說明', '設備號碼', '服務人員', '工作內容']]
    keyword = request.args.get('keyword', '').strip()
    store_id = request.args.get('store_id', '').strip()
    repair_item = request.args.get('repair_item', '').strip()
    
    tables = []
    
    if keyword or store_id or repair_item:
        if keyword:
            df = df[df.apply(lambda r: r.astype(str).str.contains(keyword, case=False).any(), axis=1)]
        if store_id:
            df = df[df['門店編號'].astype(str).str.contains(store_id, case=False)]
        if repair_item:
            df = df[df['報修類別'].astype(str).str.strip() == repair_item.strip()]
        tables = df.to_dict(orient='records')
        
    return render_template(
        'report.html',
        page_header="POS 相關",
        version=version_time,
        tables=tables,
        keyword=keyword,
        store_id=store_id,
        repair_item=repair_item,
        no_data_found=(len(tables) == 0 and (keyword or store_id or repair_item)),
        billing_invoice_log=False,
        home_page=False
    )

@app.route('/sm_web')
def sm_web_page():
    return render_template(
        'sm_web.html',
        sm_web=True,
        version=version_time,
        billing_invoice_log=False,
        home_page=False
    )

@app.route('/time')
def time_page():
    xls = pd.ExcelFile("MFP/MFP.xlsx")

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
                table_html = filtered[['零件名稱', '部件', '料號', '型號']].to_html(classes="data-table", index=False, border=0)
                
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
def calendar_page():
    return render_template(
        'calendar.html',
        version=version_time,
        calendar_page=True
    )

@app.route('/calendar/events')
def calendar_events():
    try:
        xls = load_excel_from_github(GITHUB_XLSX_URL)
        df = pd.read_excel(xls, sheet_name='行事曆')
    except:
        return jsonify([])
    df.columns = df.columns.str.strip()
    today = datetime.today().date()
    events = []
    for _, row in df.iterrows():
        date_val = row.get('date')
        title_val = row.get('title', '')
        if pd.notna(date_val) and title_val:
            start_date = pd.to_datetime(date_val).date()
            color_map = {"狄澤洋": "red", "V": "red", "湯家瑋": "green", "吳宗鴻": "orange", "劉柏均": "skyblue"}
            color = color_map.get(row.get('屬性'), "blue")
            is_special = str(row.get('特殊', '')).strip().upper() in ['Y', 'YES', '1']
            if start_date < today:
                color = "gray"
            events.append({
                "title": str(title_val),
                "start": start_date.strftime('%Y-%m-%d'),
                "color": color,
                "extendedProps": {
                    "is_special": is_special
                }
            })
    return jsonify(events)

@app.route("/worktime")
def worktime():
    path = "MFP/MFP.xlsx"

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

# SC硬碟檢測 儲存/刪除同步 API
@app.route('/sc_disk/update', methods=['POST'])
def update_sc_disk():
    data = request.json
    action = data.get('action')
    store_id = str(data.get('store_id', '')).strip()

    try:
        sh = client.open_by_key(SHEET_ID)
        ws = sh.worksheet("硬碟檢測")
        records = ws.get_all_records()
        
        row_idx = None
        for idx, r in enumerate(records, start=2):
            if str(r.get('門店編號')).strip() == store_id:
                row_idx = idx
                break

        if action == 'delete':
            if row_idx:
                ws.delete_rows(row_idx)
            return jsonify({'status': 'success', 'message': '已同步刪除 Google Sheet 資料'})

        elif action == 'save':
            row_data = [
                data.get('leave_time', ''),
                store_id,
                data.get('store_name', ''),
                data.get('repair_cat', ''),
                data.get('work_content', ''),
                data.get('sc1', ''),
                data.get('sc2', ''),
                data.get('tm1', ''),
                data.get('tm2', '')
            ]
            
            if row_idx:
                ws.update(f'A{row_idx}:I{row_idx}', [row_data])
            else:
                ws.append_row(row_data)
                
            return jsonify({'status': 'success', 'message': '資料已同步至 Google Sheet'})

    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

# ====== 啟動 Flask ======
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)