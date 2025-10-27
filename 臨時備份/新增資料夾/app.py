from flask import Flask, render_template
import pandas as pd
import requests
from io import BytesIO
from flask import Flask, render_template, request, abort
import os
import io
import base64
import matplotlib
matplotlib.use('Agg')  # 非 GUI 模式
import matplotlib.pyplot as plt
from matplotlib import rcParams
import matplotlib.font_manager as fm

# 設定中文字體
font_path = 'C:/Windows/Fonts/msjh.ttc'  # 微軟正黑體
font_prop = fm.FontProperties(fname=font_path)
rcParams['font.family'] = font_prop.get_name()

app = Flask(__name__)

GITHUB_XLSX_URL = 'https://raw.githubusercontent.com/Diyn19/flask-excel-website/master/data.xlsx'
cached_xls = None
version_time = None  # 用來儲存 G1 儲存格的版本資訊

def load_excel_from_github(url):
    global cached_xls, version_time
    if cached_xls:
        return cached_xls
    try:
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            content_type = response.headers.get('Content-Type', '')
            if 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' in content_type or url.endswith('.xlsx'):
                excel_bytes = BytesIO(response.content)
                cached_xls = pd.ExcelFile(excel_bytes, engine='openpyxl')

                # 讀取首頁 G1 作為版本資訊
                df_version = pd.read_excel(cached_xls, sheet_name='首頁', header=None, usecols="G", nrows=1)
                version_time = str(df_version.iat[0, 0]) if not pd.isna(df_version.iat[0, 0]) else "無版本資訊"

                return cached_xls
        print(f"❌ Excel 下載失敗：{response.status_code} - {content_type}")
    except Exception as e:
        print(f"❌ 錯誤下載 Excel: {e}")
    abort(500, description="⚠️ 無法從 GitHub 載入 Excel 檔案")

def clean_df(df):
    df.columns = df.columns.astype(str).str.replace('\n', '', regex=False)
    return df.fillna('')

@app.route('/')
def index():
    xls = load_excel_from_github(GITHUB_XLSX_URL)

    df_department = clean_df(pd.read_excel(xls, sheet_name='首頁', usecols="A:F", skiprows=4, nrows=1))
    df_seasons = clean_df(pd.read_excel(xls, sheet_name='首頁', usecols="A:D", skiprows=8, nrows=2))
    df_project1 = clean_df(pd.read_excel(xls, sheet_name='首頁', usecols="A:E", skiprows=12, nrows=3))
    df = clean_df(pd.read_excel(xls, sheet_name=0, header=21, nrows=250, usecols="A:O"))
    df = df[['門市編號', '門市名稱', 'PMQ_檢核', '專案檢核', 'HUB', '完工檢核']]

    keyword = request.args.get('keyword', '').strip()
    no_data_found = False
    if keyword:
        df = df[df.apply(lambda row: row.astype(str).str.contains(keyword, case=False).any(), axis=1)]
        no_data_found = df.empty

    return render_template(
        'index.html',
        tables=df.to_dict(orient='records'),
        keyword=keyword,
        store_id='',
        repair_item='',
        personal_page=False,
        report_page=False,
        department_table=df_department.to_dict(orient='records'),
        seasons_table=df_seasons.to_dict(orient='records'),
        project1_table=df_project1.to_dict(orient='records'),
        no_data_found=no_data_found,
        version=version_time
    )

@app.route('/<name>')
def personal(name):
    sheet_map = {
        '吳宗鴻': '吳宗鴻',
        '湯家瑋': '湯家瑋',
        '狄澤洋': '狄澤洋'
    }
    sheet_name = sheet_map.get(name)
    if not sheet_name:
        return f"找不到{name}的分頁", 404

    xls = load_excel_from_github(GITHUB_XLSX_URL)

    df_top = clean_df(pd.read_excel(xls, sheet_name=sheet_name, usecols="A:G", nrows=4))
    df_top = df_top.applymap(lambda x: int(x) if isinstance(x, (int, float)) and x == int(x) else x)

    df_project = clean_df(pd.read_excel(xls, sheet_name=sheet_name, usecols="H:L", nrows=3))
    df_project = df_project.applymap(lambda x: int(x) if isinstance(x, (int, float)) and x == int(x) else x)

    df_bottom = clean_df(pd.read_excel(xls, sheet_name=sheet_name, usecols="A:J", skiprows=5))

    keyword = request.args.get('keyword', '').strip()
    no_data_found = False
    if keyword:
        df_bottom = df_bottom[df_bottom.apply(lambda row: row.astype(str).str.contains(keyword, case=False).any(), axis=1)]
        no_data_found = df_bottom.empty

    return render_template(
        'index.html',
        tables_top=df_top.to_dict(orient='records'),
        tables_project=df_project.to_dict(orient='records'),
        tables_bottom=df_bottom.to_dict(orient='records'),
        keyword=keyword,
        store_id='',
        repair_item='',
        personal_page=True,
        report_page=False,
        no_data_found=no_data_found,
        show_top=True,
        show_project=True,
        version=version_time
    )

@app.route('/report')
def report():
    keyword = request.args.get('keyword', '').strip()
    store_id = request.args.get('store_id', '').strip()
    repair_item = request.args.get('repair_item', '').strip()
    no_data_found = False
    tables = []

    if keyword or store_id or repair_item:
        xls = load_excel_from_github(GITHUB_XLSX_URL)

        df = clean_df(pd.read_excel(xls, sheet_name='IM'))
        df = df[['案件類別', '門店編號', '門店名稱', '報修時間', '報修類別', '報修項目', '報修說明', '設備號碼', '服務人員', '工作內容']]

        if keyword:
            df = df[df.apply(lambda row: row.astype(str).str.contains(keyword, case=False).any(), axis=1)]

        if store_id:
            df = df[df['門店編號'].astype(str).str.contains(store_id, case=False)]

        if repair_item:
            df = df[df['報修類別'].astype(str).str.strip() == repair_item.strip()]

        if df.empty:
            no_data_found = True
        else:
            tables = df.to_dict(orient='records')

    return render_template(
        'index.html',
        tables=tables,
        keyword='',
        store_id=store_id,
        repair_item=repair_item,
        personal_page=False,
        report_page=True,
        no_data_found=no_data_found,
        version=version_time
    )
    
@app.route('/time')
def time():
    xls = load_excel_from_github(GITHUB_XLSX_URL)

    # 讀取版本號
    try:
        version_df = pd.read_excel(xls, sheet_name='首頁', header=None, usecols="G", nrows=1)
        version = version_df.iloc[0, 0]
    except:
        version = "無法讀取版本號"

    # 讀取摘要與明細資料（保留你原本的）
    df_summary = pd.read_excel(xls, sheet_name='出勤時間', usecols="A:E", nrows=2)
    detail_1 = pd.read_excel(xls, sheet_name='出勤時間', usecols="A:Q", skiprows=3, nrows=3)
    detail_2 = pd.read_excel(xls, sheet_name='出勤時間', usecols="A:Q", skiprows=7, nrows=3)
    detail_3 = pd.read_excel(xls, sheet_name='出勤時間', usecols="A:Q", skiprows=11, nrows=3)

    # 讀取曲線圖資料
    # 時間軸：B12:P12 → index=11，col=1~15（0-based）
    # 姓名：A13:A15 → index=12~14，col=0
    # 出勤次數：B13:P15 → index=12~14，col=1~15

    df_chart = pd.read_excel(xls, sheet_name='出勤時間', header=None)

    # 取得時間軸字串
    x = df_chart.iloc[11, 1:16].tolist()  # B12:P12 (Excel 1-based, DataFrame 0-based)
    # 確認 x 是什麼格式，如果是數字（Excel 時間序列），轉成時間格式字串
    if all(isinstance(v, (int, float)) for v in x):
        # Excel 日期是從 1900-01-01 起算，時間是一天的小數部分
        # 假設這裡是時間欄，直接用 pd.Timedelta 轉小時分鐘
        x = [(pd.Timestamp("1900-01-01") + pd.Timedelta(hours=hour)).strftime("%H:%M") for hour in range(8, 23)]
    else:
        # 否則直接轉成字串（保險用）
        x = [str(v) for v in x]

    # 取得人員姓名
    names = df_chart.iloc[12:15, 0].tolist()

    # 取得次數資料，轉成 list of list，3 行×15 欄
    y_data = df_chart.iloc[12:15, 1:16].values.tolist()

    # 畫圖
    fig, ax = plt.subplots(figsize=(10, 5))
    for i, y in enumerate(y_data):
        ax.plot(x, y, marker='o', label=names[i])

    ax.set_xlabel("時間")
    ax.set_ylabel("出勤次數")
    ax.set_title("出勤時間曲線圖")
    ax.legend()
    plt.xticks(rotation=45)
    plt.tight_layout()

    # 圖片轉 base64
    img = io.BytesIO()
    plt.savefig(img, format='png')
    img.seek(0)
    plot_url = base64.b64encode(img.read()).decode('utf-8')
    plt.close()

    return render_template(
        'index.html',
        summary_table=df_summary.to_html(index=False, classes='dataframe'),
        detail_table_1=detail_1.to_html(index=False, classes='dataframe'),
        detail_table_2=detail_2.to_html(index=False, classes='dataframe'),
        detail_table_3=detail_3.to_html(index=False, classes='dataframe'),
        version=version,
        plot_url=plot_url,
        time_page=True  # 這行很重要
    )
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
