from flask import Flask, render_template, request, jsonify, abort
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
from flask import redirect, url_for
import gspread
from google.oauth2.service_account import Credentials
from modules.gsheet import client, SHEET_ID


# ====== 新增：引入 billing 模組 ======
from modules.billing import billing_bp   # ✅ 新增這一行

# ====== 基本設定 ======
app = Flask(__name__)
GITHUB_XLSX_URL = 'https://raw.githubusercontent.com/Yang-0419-di/FW_2/master/data.xlsx'
cached_xls = None
version_time = None
app.config['VERSION_TIME'] = version_time

# ====== 新增：註冊 billing 藍圖 ======
app.register_blueprint(billing_bp)  # ✅ 新增這一行

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
    abort(500, description="⚠️ 無法從 GitHub 載入 Excel 檔案")

def clean_df(df):
    df.columns = df.columns.astype(str).str.replace('\n', '', regex=False)
    return df.fillna('')

# ====== 以下為原有功能（完全不變） ======
@app.route('/')
def home():
    xls = load_excel_from_github(GITHUB_XLSX_URL)

    # ====== 原本首頁資料 ======
    df_department = clean_df(pd.read_excel(xls, sheet_name='首頁', usecols="A:F", skiprows=4, nrows=1))
    df_seasons = clean_df(pd.read_excel(xls, sheet_name='首頁', usecols="A:D", skiprows=8, nrows=2))
    df_project1 = clean_df(pd.read_excel(xls, sheet_name='首頁', usecols="A:E", skiprows=12, nrows=3))

    df_HUB = clean_df(
        pd.read_excel(
            xls,
            sheet_name='首頁',
            header=18,      # 第 19 列當欄位名稱
            usecols="A:E"   # 只抓 A~D
        )
    )
    df_HUB = df_HUB[df_HUB['門市編號'].astype(str).str.strip() != '']
    df_HUB['門市編號'] = df_HUB['門市編號'].astype(str).str.replace(r'\.0$', '', regex=True)
    df_HUB = df_HUB[['門市編號', '門市名稱', 'HUB規格', '異常原因', '完工確認']]


    df = clean_df(pd.read_excel(xls, sheet_name=0, header=21, nrows=500, usecols="A:O"))
    df = df[['門市編號', '門市名稱', 'PMQ_檢核', '專案檢核', 'HUB', '完工檢核']]

    keyword = request.args.get('keyword', '').strip()
    no_data_found = False
    if keyword:
        df = df[df.apply(lambda r: r.astype(str).str.contains(keyword, case=False).any(), axis=1)]
        no_data_found = df.empty


    # ======================================================
    #                 區域數量（三段）- 都在「首頁」
    # ======================================================

    # === 第一段 A55:G55 標題，A56:G57 內容 ===
    df1 = pd.read_excel(
        xls,
        sheet_name='首頁',
        header=None,
        usecols="E:K",
        skiprows=54,
        nrows=3
    )

    headers1 = df1.iloc[0].tolist()
    area_table_1 = []
    for i in range(1, 3):
        area_table_1.append(dict(zip(headers1, df1.iloc[i].tolist())))


    # === 第二段 A59:L59 標題，A60:L61 內容 ===
    df2 = pd.read_excel(
        xls,
        sheet_name='首頁',
        header=None,
        usecols="E:P",
        skiprows=58,
        nrows=3
    )

    headers2 = df2.iloc[0].tolist()
    area_table_2 = []
    for i in range(1, 3):
        area_table_2.append(dict(zip(headers2, df2.iloc[i].tolist())))


    # === 第三段 A63:G63 標題，A64:G65 內容 ===
    df3 = pd.read_excel(
        xls,
        sheet_name='首頁',
        header=None,
        usecols="E:L",
        skiprows=62,
        nrows=3
    )

    headers3 = df3.iloc[0].tolist()
    area_table_3 = []
    for i in range(1, 3):
        area_table_3.append(dict(zip(headers3, df3.iloc[i].tolist())))


    # ======================================================
    #                 回傳到 home.html
    # ======================================================

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
        HUB_table=df_HUB.to_dict(orient='records'),

        no_data_found=no_data_found,
        billing_invoice_log=False,
        home_page=True
    )

from flask import redirect, url_for


@app.route("/disk", methods=["GET"])
def disk_page():
    # 讀取 Google Sheet
    try:
        sh = client.open_by_key(SHEET_ID)
        sheet = sh.worksheet("硬碟統計")  # ← 指定硬碟統計分頁
    except gspread.exceptions.APIError as e:
        return f"⚠️ 無法讀取 Google Sheet: {e}", 500

    # 讀取所有資料
    all_rows = sheet.get_all_records()  # list of dict

    # 只取每個 user 最新一筆資料
    latest_data = {}
    for row in all_rows:
        user = row.get('user')
        if user:
            latest_data[user] = row  # 後面會覆蓋前面，保留最後一筆

    rows = list(latest_data.values())

    # 計算總計
    total_keys = [
        'sc_128_new','sc_128_old','sc_240_new','sc_240_old',
        'sc_256_new','sc_256_old','sc_500_new','sc_500_old',
        'sc_1t_new','sc_1t_old','tm_128_new','tm_128_old','tm_256_new','tm_256_old'
    ]
    total = {k: sum(int(r.get(k) or 0) for r in rows) for k in total_keys}

    return render_template("disk.html", page_header="POS 相關", rows=rows, total=total)


# ====== /disk/save 儲存 ======
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

    # 直接 append 一列到硬碟統計分頁
    try:
        sh = client.open_by_key(SHEET_ID)
        sheet = sh.worksheet("硬碟統計")  # ← 指定硬碟統計分頁
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
    version = version_time,
    sheet_map = {'吳宗鴻': '吳宗鴻', '湯家瑋': '湯家瑋', '狄澤洋': '狄澤洋','劉柏均': '劉柏均'}
    sheet_name = sheet_map.get(name)
    if not sheet_name:
        return f"找不到 {name} 的分頁", 404

    xls = load_excel_from_github(GITHUB_XLSX_URL)

    # --- 其他表格 ---
    df_top = clean_df(pd.read_excel(xls, sheet_name=sheet_name, usecols="A:G", nrows=4))
    df_project = clean_df(pd.read_excel(xls, sheet_name=sheet_name, usecols="H:L", nrows=4))
    df_bottom = clean_df(pd.read_excel(xls, sheet_name=sheet_name, usecols="A:J", skiprows=5))

    # --- 正確讀取區域數量 W1:AE2 ---
    df_area = pd.read_excel(
        xls,
        sheet_name=sheet_name,
        usecols="W:AE",
        nrows=1,        # ← 標題 + 數值
        header=0        # ← 第一列當標題
    )

    # ★★★★★ 強制還原 '-' 欄名 ★★★★★
    df_area.columns = df_area.columns.map(
        lambda x: "-" if str(x).strip().startswith("-") else str(x)
    )

    # ---- 搜尋功能 for 下方門市 ----
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

        tables_top=df_top.to_dict(orient="records"),
        tables_project=df_project.to_dict(orient="records"),
        tables_bottom=df_bottom.to_dict(orient="records"),

        # 區域數量（直接給 dataframe）
        tables_area=df_area.to_dict(orient="records"),

        version=version_time,
        billing_invoice_log=False,
        home_page=False
    )



@app.route('/report')
def report():
    xls = load_excel_from_github(GITHUB_XLSX_URL)
    version=version_time,
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

@app.route('/time')
def time_page():
    xls = load_excel_from_github(GITHUB_XLSX_URL)

    # ======================================================
    # 區塊一：摘要 A1:F1 / A2:F2
    # ======================================================
    df_summary = pd.read_excel(
        xls,
        sheet_name='出勤時間',
        usecols="A:F",
        header=0,   # A1:F1
        nrows=1     # A2:F2
    )

    # ======================================================
    # 區塊二：A4:Q4 / A5:Q8
    # ======================================================
    detail_1 = pd.read_excel(
        xls,
        sheet_name='出勤時間',
        usecols="A:Q",
        header=3,   # A4:Q4
        nrows=4     # A5:Q8
    )

    # ======================================================
    # 區塊三：A9:Q9 / A10:Q13
    # ======================================================
    detail_2 = pd.read_excel(
        xls,
        sheet_name='出勤時間',
        usecols="A:Q",
        header=8,   # A9:Q9
        nrows=4     # A10:Q13
    )

    # ======================================================
    # 區塊四：A14:Q14 / A15:Q18
    # ======================================================
    detail_3 = pd.read_excel(
        xls,
        sheet_name='出勤時間',
        usecols="A:Q",
        header=13,  # A14:Q14
        nrows=4     # A15:Q18
    )

    # ======================================================
    # 圖表資料（4 人）
    # ======================================================
    df_chart = pd.read_excel(xls, sheet_name='出勤時間', header=None)

    # =============================
    # X 軸：B14:P14
    # =============================
    x = [str(v) for v in df_chart.iloc[13, 1:16].tolist()]

    # =============================
    # Y 軸：B15:P18（4 個人）
    # =============================
    names = df_chart.iloc[14:18, 0].tolist()
    y_data = df_chart.iloc[14:18, 1:16].values.tolist()

    # =============================
    # 畫圖
    # =============================
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


    # ======================================================
    # 回傳頁面
    # ======================================================
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
    version=version_time,
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
                table_html = filtered[['零件名稱', '料號', '型號']].to_html(classes="data-table", index=False, border=0)
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
    version=version_time,
    return render_template('calendar.html', version=version_time,)

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
            if start_date < today:
                color = "gray"
            events.append({"title": str(title_val), "start": start_date.strftime('%Y-%m-%d'), "color": color})
    return jsonify(events)
    
@app.route("/worktime")
def worktime():
    import pandas as pd

    path = "MFP/MFP.xlsx"

    # ================================
    # 區塊 1：計算基礎、單位(min)
    # A2:F3 → A2 是標題列
    # ================================
    df_1 = pd.read_excel(
        path,
        sheet_name="工時計算",
        header=None,
        usecols="A:F",
        skiprows=1,     # 從 A2 開始
        nrows=2
    )
    block1_header = df_1.iloc[0].tolist()
    block1_body = df_1.iloc[1:].values.tolist()

    # ================================
    # 區塊 2：跑勤統計（多一列）
    # A5:H9 → A5 標題、A9 說明
    # ================================
    df_2 = pd.read_excel(
        path,
        sheet_name="工時計算",
        header=None,
        usecols="A:H",
        skiprows=4,     # A5
        nrows=5         # A5～A9（比原本多 1 列）
    )

    block2_header = df_2.iloc[0].tolist()
    block2_body = df_2.iloc[1:-1].values.tolist()   # A6～A8（3 列）
    block2_note = df_2.iloc[-1].tolist()            # A9 說明列

    # ================================
    # 區塊 3：維修統計（多一列）
    # A10:H13 → A10 標題
    # ================================
    df_3 = pd.read_excel(
        path,
        sheet_name="工時計算",
        header=None,
        usecols="A:H",
        skiprows=9,     # A10
        nrows=4         # A10～A13（比原本多 1 列）
    )

    block3_header = df_3.iloc[0].tolist()
    block3_body = df_3.iloc[1:].values.tolist()     # A11～A13（3 列）

    # ================================
    # 區塊 4：工時計算（多一列）
    # A14:K18 → A18 說明列
    # ================================
    df_4 = pd.read_excel(
        path,
        sheet_name="工時計算",
        header=None,
        usecols="A:K",
        skiprows=13,    # A14
        nrows=5         # A14～A18（比原本多 1 列）
    )

    block4_header = df_4.iloc[0].tolist()
    block4_body = df_4.iloc[1:-1].values.tolist()   # A15～A17（3 列）
    block4_note = df_4.iloc[-1].tolist()            # A18 說明列

    return render_template(
        "worktime.html",
        block1_header=block1_header, block1_body=block1_body,
        block2_header=block2_header, block2_body=block2_body, block2_note=block2_note,
        block3_header=block3_header, block3_body=block3_body,
        block4_header=block4_header, block4_body=block4_body, block4_note=block4_note,
        billing_worktime=True
    )



# ====== 啟動 Flask ======
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
