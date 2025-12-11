from flask import Flask, render_template, request, jsonify, abort
import pandas as pd
import requests
from io import BytesIO
import os, io, base64
from datetime import datetime
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties

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

    df_HUB = clean_df(pd.read_excel(xls, sheet_name='首頁', header=18, nrows=30, usecols="A:D"))
    df_HUB = df_HUB[['門市編號', '門市名稱', '異常原因', '完工確認']]

    df = clean_df(pd.read_excel(xls, sheet_name=0, header=21, nrows=250, usecols="A:O"))
    df = df[['門市編號', '門市名稱', 'PMQ_檢核', '專案檢核', 'HUB', '完工檢核']]

    keyword = request.args.get('keyword', '').strip()
    no_data_found = False
    if keyword:
        df = df[df.apply(lambda r: r.astype(str).str.contains(keyword, case=False).any(), axis=1)]
        no_data_found = df.empty

    # ======================================================
    #                 區域數量（新增）
    # ======================================================

    # === 第一段 C2:k3 ===
    df1 = pd.read_excel(
        xls,
        sheet_name='門市主檔',
        header=None,
        usecols="C:k",
        skiprows=1,
        nrows=2
    )
    headers1 = df1.iloc[0].tolist()   # C2:k2
    values1 = df1.iloc[1].tolist()    # C3:k3
    area_table_1 = [dict(zip(headers1, values1))]

    # === 第二段 l2:p3 ===
    df2 = pd.read_excel(
        xls,
        sheet_name='門市主檔',
        header=None,
        usecols="l:p",
        skiprows=1,
        nrows=2
    )
    headers2 = df2.iloc[0].tolist()
    values2 = df2.iloc[1].tolist()
    area_table_2 = [dict(zip(headers2, values2))]

    # ======================================================
    #                 回傳到 home.html
    # ======================================================

    return render_template(
        'home.html',
        version=version_time,

        # 新增 - 區域數量三段資料
        area_table_1=area_table_1,
        area_table_2=area_table_2,

        # 原本首頁資料
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


@app.route('/personal/<name>')
def personal(name):
    version = version_time,
    sheet_map = {'吳宗鴻': '吳宗鴻', '湯家瑋': '湯家瑋', '狄澤洋': '狄澤洋'}
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
    version=version_time,
    df_summary = pd.read_excel(xls, sheet_name='出勤時間', usecols="A:E", nrows=2)
    detail_1 = pd.read_excel(xls, sheet_name='出勤時間', usecols="A:Q", skiprows=3, nrows=3)
    detail_2 = pd.read_excel(xls, sheet_name='出勤時間', usecols="A:Q", skiprows=7, nrows=3)
    detail_3 = pd.read_excel(xls, sheet_name='出勤時間', usecols="A:Q", skiprows=11, nrows=3)
    df_chart = pd.read_excel(xls, sheet_name='出勤時間', header=None)
    x = [str(v) for v in df_chart.iloc[11, 1:16].tolist()]
    names = df_chart.iloc[12:15, 0].tolist()
    y_data = df_chart.iloc[12:15, 1:16].values.tolist()
    fig, ax = plt.subplots(figsize=(10, 5))
    for i, y in enumerate(y_data):
        ax.plot(x, y, marker='o', label=names[i])
    plt.xticks(rotation=45)
    plt.legend()
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
            color_map = {"狄澤洋": "red", "V": "red", "湯家瑋": "green", "吳宗鴻": "orange"}
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
    # A2:F3  → A2 是標題列
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
    # 區塊 2：跑勤統計
    # A5:H8 → A5 是標題列、A8:H8 是說明列（要可收合）
    # ================================
    df_2 = pd.read_excel(
        path,
        sheet_name="工時計算",
        header=None,
        usecols="A:H",
        skiprows=4,   # 從 A5 開始
        nrows=4       # A5～A8 共 4 列
    )

    block2_header = df_2.iloc[0].tolist()
    block2_body = df_2.iloc[1:3].values.tolist()      # A6～A7
    block2_note = df_2.iloc[3].tolist()               # A8 說明列

    # ================================
    # 區塊 3：維修統計
    # A10:H12 → A10 標題列
    # ================================
    df_3 = pd.read_excel(
        path,
        sheet_name="工時計算",
        header=None,
        usecols="A:H",
        skiprows=9,   # 從 A10 開始
        nrows=3
    )

    block3_header = df_3.iloc[0].tolist()
    block3_body = df_3.iloc[1:].values.tolist()

    # ================================
    # 區塊 4：工時計算
    # A14:K16 → A16 說明列（要可收合）
    # ================================
    df_4 = pd.read_excel(
        path,
        sheet_name="工時計算",
        header=None,
        usecols="A:K",
        skiprows=13,    # A14
        nrows=4         # A14～A17
    )

    block4_header = df_4.iloc[0].tolist()
    block4_body = df_4.iloc[1:3].values.tolist()     # A15
    block4_note = df_4.iloc[3].tolist()              # A17 說明列

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
