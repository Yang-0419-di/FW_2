# === billing.py ===
from flask import Blueprint, render_template, request, redirect, url_for, abort
import sqlite3
from datetime import datetime
from flask import Blueprint, render_template, request, current_app
import requests
from io import BytesIO
import pandas as pd
from modules.gsheet import get_person_worksheet 
from modules.gsheet import get_customer_worksheet
from modules.gsheet import get_contract_worksheet

billing_bp = Blueprint('billing', __name__, url_prefix='/billing')

GITHUB_XLSX_URL = 'https://raw.githubusercontent.com/Yang-0419-di/FW_2/master/MFP/MFP.xlsx'
_cached_xls = None   # 快取避免多次下載
bp = Blueprint("billing", __name__, url_prefix="/billing")
DB_FILE = "billing.db"

def to_int(val):
    try:
        return int(float(val))
    except (TypeError, ValueError):
        return 0


# --- 計算保養逾期狀態並加上顏色 ---
def color_overdue(val, last_pm, cycle):
    """
    val: 客戶名稱
    last_pm: 最後保養日 (str)
    cycle: 保養週期 (str 或 int)
    """
    # 防呆：空白保養日或空白週期都算逾期
    overdue = False

    if last_pm == "" or cycle == "":
        overdue = True
    else:
        # 特例：合約規範視為30天
        try:
            cycle_days = int(cycle) if str(cycle) != "合約規範" else 30
        except:
            cycle_days = 30  # 防呆

        try:
            last_date = pd.to_datetime(last_pm)
            today = datetime.today()
            delta_days = (today - last_date).days
            if delta_days > cycle_days:
                overdue = True
        except:
            overdue = True  # 無法解析日期也算逾期

    if overdue:
        # HTML 加上淺粉紅背景
        return f'<span style="background-color:#FFC0CB">{val}</span>'
    else:
        return val



# ================================================================
# 讀取 GitHub / 本地 Excel（支援檔名參數）
# ================================================================
_cached_xls = None  # 快取字典 {'filename': '...', 'xls': pd.ExcelFile}

def load_github_excel(filename="MFP.xlsx"):
    """
    安全下載 GitHub RAW EXCEL（含快取與 fallback）
    filename: 可選，本地 fallback 使用的 Excel 檔名
    """
    import requests
    from io import BytesIO
    import pandas as pd

    global _cached_xls

    if _cached_xls and _cached_xls['filename'] == filename:
        return _cached_xls['xls']

    try:
        resp = requests.get(GITHUB_XLSX_URL, timeout=10)
        if resp.status_code != 200:
            raise Exception(f"HTTP {resp.status_code}")

        excel_bytes = BytesIO(resp.content)

        import zipfile
        if not zipfile.is_zipfile(excel_bytes):
            raise Exception("下載內容不是 Excel（不是 zip 格式）")

        xls = pd.ExcelFile(excel_bytes, engine="openpyxl")
        _cached_xls = {'filename': filename, 'xls': xls}
        return xls

    except Exception as e:
        print(f"⚠ GitHub Excel 載入失敗，改用本地 {filename}，原因：{e}")
        local_path = f"MFP/{filename}"  # 本地 fallback
        xls = pd.ExcelFile(local_path, engine="openpyxl")
        _cached_xls = {'filename': filename, 'xls': xls}
        return xls


    except Exception as e:
        print(f"⚠ GitHub Excel '{filename}' 載入失敗，改用本地 fallback，原因：", e)

        # fallback 本地路徑，假設 MFP.xlsx 與 output.xlsx 都在 MFP/資料夾
        local_path = f"MFP/{filename}"
        xls = pd.ExcelFile(local_path, engine="openpyxl")

        _cached_xls = {"filename": filename, "xls": xls}
        return xls

# --- 初始化資料庫（完整，不略） ---
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    # 契約資料表（含稅別欄位與 contra）
    c.execute("""
        CREATE TABLE IF NOT EXISTS contracts (
            device_id TEXT PRIMARY KEY,             -- 設備編號唯一
            monthly_rent REAL,                      -- 月租金（含稅或未稅，REAL 可存小數）
            color_unit_price REAL,                  -- 彩色單價(A4)
            bw_unit_price REAL,                     -- 黑白單價
            color_a3_unit_price REAL DEFAULT 0,     -- 彩色單價(A3)，預設 0
            color_giveaway INTEGER,                 -- 彩色贈送張數
            bw_giveaway INTEGER,                    -- 黑白贈送張數
            color_a3_giveaway INTEGER DEFAULT 0,    -- 彩色A3贈送張數
            color_error_rate REAL,                  -- 彩色誤印率
            bw_error_rate REAL,                     -- 黑白誤印率
            color_a3_error_rate REAL DEFAULT 0,     -- 彩色A3誤印率
            color_basic INTEGER,                     -- 彩色基本張數
            bw_basic INTEGER,                        -- 黑白基本張數
            color_a3_basic INTEGER DEFAULT 0,        -- 彩色A3基本張數
            tax_type TEXT DEFAULT '含稅',           -- 稅別
            contra TEXT DEFAULT '',                  -- 契約說明
            master_device_id TEXT DEFAULT ''         -- 合開主機設備編號
        )
    """)

    # 抄表記錄表
    c.execute("""
        CREATE TABLE IF NOT EXISTS usage (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            device_id TEXT,
            month TEXT,
            color_a3_count INTEGER DEFAULT 0,
            color_count INTEGER,
            bw_count INTEGER,
            timestamp TEXT,
            last_date TEXT DEFAULT '' 
        )
    """)

    # 客戶資料表
    c.execute("""
        CREATE TABLE IF NOT EXISTS customers (
            device_id TEXT PRIMARY KEY,
            customer_name TEXT,
            device_number TEXT,
            machine_model TEXT,
            tax_id TEXT,
            install_address TEXT,
            service_person TEXT,
            contract_number TEXT,
            contract_start TEXT,
            contract_end TEXT
        )
    """)

    # ✅ 發票/計費月結摘要表（每台每月一筆，若已有則覆蓋）
    c.execute("""
        CREATE TABLE IF NOT EXISTS billing_summary (
            device_id TEXT,
            month INTEGER, -- 1~12

            -- ===== 本月抄表總數（合開合計） =====
            color_a3_total INTEGER,   -- 本月抄表 彩色 A3 總張數
            color_total INTEGER,      -- 本月抄表 彩色 總張數
            bw_total INTEGER,         -- 本月抄表 黑白 總張數

            -- ===== 當月使用量（delta） =====
            color_a3_usage INTEGER,   -- 彩色 A3 使用量 = 本月 - 上月
            color_usage INTEGER,      -- 彩色 使用量 = 本月 - 上月
            bw_usage INTEGER,         -- 黑白 使用量 = 本月 - 上月

            -- ===== 實際計費張數 =====
            color_a3_bill_usage INTEGER, -- 彩色 A3 計費張數
            color_bill_usage INTEGER,    -- 彩色 計費張數
            bw_bill_usage INTEGER,       -- 黑白 計費張數

            -- ===== 金額 =====
            color_a3_amount REAL,     -- 彩色 A3 金額
            color_amount REAL,        -- 彩色 金額
            bw_amount REAL,           -- 黑白 金額
            monthly_rent REAL,        -- 月租金

            -- ===== 發票金額 =====
            untaxed_subtotal REAL,    -- 未稅小計（彩色A3 + 彩色 + 黑白 + 月租）
            tax_amount REAL,          -- 稅額
            total_with_tax REAL,      -- 含稅總額

            PRIMARY KEY (device_id, month)
        )
    """)

    conn.commit()
    conn.close()


# 呼叫初始化以確保資料表存在（可在應用啟動時呼叫一次）
init_db()

def safe_int(val):
    try:
        return int(val)
    except:
        return 0


# --- 查詢契約 ---
def get_contract(device_id):
    ws = get_person_worksheet("contracts")
    rows = ws.get_all_records()   # ← 這行是關鍵

    contract_row = next(
        (row for row in rows
         if str(row.get("device_id", "")).strip() == str(device_id).strip()),
        None
    )

    if not contract_row:
        return None, ""

    contra_text = contract_row.get("contra", "")

    # 欄位型別正規化
    float_fields = [
        "monthly_rent",
        "color_unit_price", "bw_unit_price",
        "color_a3_unit_price",
        "color_error_rate", "bw_error_rate", "color_a3_error_rate",
    ]

    int_fields = [
        "color_giveaway", "bw_giveaway", "color_a3_giveaway",
        "color_basic", "bw_basic", "color_a3_basic",
    ]

    for k in float_fields:
        try:
            contract_row[k] = float(contract_row.get(k) or 0)
        except:
            contract_row[k] = 0.0

    for k in int_fields:
        try:
            contract_row[k] = int(float(contract_row.get(k) or 0))
        except:
            contract_row[k] = 0

    return contract_row, contra_text


# --- 查詢客戶資料 ---
def get_customer(device_id):
    ws = get_person_worksheet("customers")
    rows = ws.get_all_records()   # ← 必須加這行

    row = next(
        (r for r in rows
         if str(r.get("device_id", "")).strip() == str(device_id).strip()),
        None
    )

    if not row:
        return None

    customer = {
        "device_id": row.get("device_id", ""),
        "customer_name": row.get("customer_name", ""),
        "pm": row.get("pm", ""),
        "device_number": row.get("device_number", ""),
        "machine_model": row.get("machine_model", ""),
        "tax_id": row.get("tax_id", ""),
        "install_address": row.get("install_address", ""),
        "service_person": row.get("service_person", ""),
        "contract_number": row.get("contract_number", ""),
        "contract_start": row.get("contract_start", ""),
        "contract_end": row.get("contract_end", "")
    }

    # 日期格式化
    for key in ["contract_start", "contract_end"]:
        val = customer.get(key)
        if val:
            try:
                dt = pd.to_datetime(val)
                customer[key] = dt.strftime("%Y/%m/%d")
            except:
                pass

    return customer


# --- 模糊搜尋客戶名稱（Google Sheet 版本） ---
def search_customers_by_name(keyword):
    ws = get_person_worksheet("customers")
    rows = ws.get_all_records()   # ← 必須加這行

    result = []
    keyword_lower = keyword.lower().strip()

    for r in rows:
        if keyword_lower in str(r.get("customer_name", "")).lower():
            result.append({
                "device_id": r.get("device_id", ""),
                "customer_name": r.get("customer_name", "")
            })

    return result

def update_contract(device_id, contract_data):
    ws = get_person_worksheet("contracts")
    rows = ws.get_all_records()

    for idx, row in enumerate(rows, start=2):  # start=2 因為第1列是表頭
        if str(row.get("device_id", "")).strip() == str(device_id).strip():

            for key, value in contract_data.items():
                try:
                    col = list(row.keys()).index(key) + 1
                    ws.update_cell(idx, col, value)
                except:
                    pass

            return True

    return False
    
def update_customer(device_id, customer_data):
    ws = get_person_worksheet("customers")
    rows = ws.get_all_records()

    for idx, row in enumerate(rows, start=2):
        if str(row.get("device_id", "")).strip() == str(device_id).strip():

            for key, value in customer_data.items():
                try:
                    col = list(row.keys()).index(key) + 1
                    ws.update_cell(idx, col, value)
                except:
                    pass

            return True

    return False

#新客戶建檔
def insert_customer(device_id, customer_data):
    """
    將新客戶資料寫入 Google Sheet customers 工作表
    """

    try:
        ws = get_customer_worksheet()

        # 取得表頭
        headers = ws.row_values(1)

        # 建立一整列空值
        new_row = [""] * len(headers)

        # 依表頭名稱填入資料
        for key, value in customer_data.items():
            if key in headers:
                col_index = headers.index(key)
                new_row[col_index] = value

        # device_id 確保一定寫入
        if "device_id" in headers:
            new_row[headers.index("device_id")] = device_id

        ws.append_row(new_row)

        return True

    except Exception as e:
        print("insert_customer error:", e)
        return False

#新合約建檔
def insert_contract(device_id, contract_data):
    """
    將新契約資料寫入 Google Sheet contracts 工作表
    """

    try:
        ws = get_contract_worksheet()

        # 取得表頭
        headers = ws.row_values(1)

        new_row = [""] * len(headers)

        for key, value in contract_data.items():
            if key in headers:
                col_index = headers.index(key)
                new_row[col_index] = value

        # 強制寫入 device_id
        if "device_id" in headers:
            new_row[headers.index("device_id")] = device_id

        ws.append_row(new_row)

        return True

    except Exception as e:
        print("insert_contract error:", e)
        return False


# --- 查詢最後抄表（含跨年） ---
def get_prev_month_year(selected_year, selected_month):
    """
    依照用戶選擇的年月，計算前月年月
    """
    # 移除了原先 request.form.get 的覆蓋程式碼，直接使用傳入的引數
    if selected_month == 1:
        return selected_year - 1, 12
    else:
        return selected_year, selected_month - 1


def get_last_counts(device_id, selected_year, selected_month):
    """
    抓取前月抄表張數，如果沒有資料則回傳 0
    """
    prev_year, prev_month = get_prev_month_year(selected_year, selected_month)

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
        SELECT color_a3_total, color_total, bw_total, last_date
        FROM billing_summary
        WHERE device_id=? AND year=? AND month=?
    """, (device_id, prev_year, prev_month))
    row = c.fetchone()
    conn.close()

    if row:
        return row[0] or 0, row[1] or 0, row[2] or 0, row[3] or ""
    else:
        return 0, 0, 0, ""


# --- 合開群組查詢 ---
def get_related_devices(device_id):
    """根據設備代號找出合開群組（主機 + 子機）"""
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    # 先確認這台設備的 master_device_id
    c.execute("SELECT master_device_id FROM contracts WHERE device_id=?", (device_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return []

    master_id = row[0]

    if not master_id or master_id.strip() == "":
        # 若該設備是主機 → 找所有子機
        c.execute("SELECT device_id FROM contracts WHERE master_device_id=?", (device_id,))
        subs = [r[0] for r in c.fetchall()]
        group = [device_id] + subs
    else:
        # 若該設備是子機 → 找主機與所有兄弟子機
        c.execute("SELECT device_id FROM contracts WHERE master_device_id=?", (master_id,))
        subs = [r[0] for r in c.fetchall()]
        group = [master_id] + subs

    conn.close()
    return group


# --- 紀錄使用量 ---
def insert_usage(device_id, color_a3, color_count, bw_count):
    month = datetime.now().strftime("%Y%m")
    timestamp = datetime.now().strftime("%Y/%m/%d-%H:%M")
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute(
        "INSERT INTO usage (device_id, month, color_a3_count, color_count, bw_count, timestamp) VALUES (?, ?, ?, ?, ?, ?)",
        (device_id, month, color_a3, color_count, bw_count, timestamp)
    )
    conn.commit()
    conn.close()


# --- 計算邏輯（彩色A3 / 彩色(A4) / 黑白 全獨立） ---
def calculate(contract, curr_color_a3, curr_color, curr_bw, last_color_a3, last_color, last_bw):
    if not contract:
        return None

    # --- 1️⃣ contract 欄位轉 float ---
    keys_float = [
        "color_a3_unit_price", "color_unit_price", "bw_unit_price",
        "color_a3_giveaway", "color_giveaway", "bw_giveaway",
        "color_a3_error_rate", "color_error_rate", "bw_error_rate",
        "color_a3_basic", "color_basic", "bw_basic",
        "monthly_rent"
    ]
    for key in keys_float:
        try:
            contract[key] = float(contract.get(key, 0))
        except:
            contract[key] = 0.0

    # --- 2️⃣ 安全轉 int ---
    def safe_int(val):
        try:
            return int(float(val))
        except:
            return 0

    curr_color_a3 = safe_int(curr_color_a3)
    curr_color    = safe_int(curr_color)
    curr_bw       = safe_int(curr_bw)
    last_color_a3 = safe_int(last_color_a3)
    last_color    = safe_int(last_color)
    last_bw       = safe_int(last_bw)

    # =========================
    # 3️⃣ 使用張數（完全獨立）
    # =========================
    used_color_a3 = max(0, curr_color_a3 - last_color_a3)
    used_color    = max(0, curr_color - last_color)
    used_bw       = max(0, curr_bw - last_bw)

    # =========================
    # 4️⃣ 彩色A3計費張數
    # =========================
    bill_color_a3 = max(0, used_color_a3 - contract["color_a3_giveaway"])
    bill_color_a3 = int(round(bill_color_a3 * (1 - contract["color_a3_error_rate"])))
    if contract["color_a3_basic"] > 0:
        bill_color_a3 = max(int(contract["color_a3_basic"]), bill_color_a3)

    # =========================
    # 5️⃣ 彩色(A4)計費張數
    # =========================
    bill_color = max(0, used_color - contract["color_giveaway"])
    bill_color = int(round(bill_color * (1 - contract["color_error_rate"])))
    if contract["color_basic"] > 0:
        bill_color = max(int(contract["color_basic"]), bill_color)

    # =========================
    # 6️⃣ 黑白計費張數
    # =========================
    bill_bw = max(0, used_bw - contract["bw_giveaway"])
    bill_bw = int(round(bill_bw * (1 - contract["bw_error_rate"])))
    if contract["bw_basic"] > 0:
        bill_bw = max(int(contract["bw_basic"]), bill_bw)

    # =========================
    # 7️⃣ 金額計算
    # =========================
    color_a3_amount = bill_color_a3 * contract["color_a3_unit_price"]
    color_amount    = bill_color * contract["color_unit_price"]
    bw_amount       = bill_bw * contract["bw_unit_price"]
    subtotal = contract["monthly_rent"] + color_a3_amount + color_amount + bw_amount

    # =========================
    # 8️⃣ 稅額計算
    # =========================
    tax_rate = 0.05
    if contract.get("tax_type") == "未稅":
        untaxed = subtotal
        tax = subtotal * tax_rate
        total = subtotal + tax
    else:
        total = subtotal
        untaxed = subtotal / (1 + tax_rate)
        tax = total - untaxed

    # =========================
    # 9️⃣ 回傳結果
    # =========================
    return {
        "彩色A3使用張數": used_color_a3,
        "彩色使用張數": used_color,
        "黑白使用張數": used_bw,
        "彩色A3計費張數": bill_color_a3,
        "彩色計費張數": bill_color,
        "黑白計費張數": bill_bw,
        "彩色A3金額": round(color_a3_amount, 2),
        "彩色金額": round(color_amount, 2),
        "黑白金額": round(bw_amount, 2),
        "月租金": round(contract["monthly_rent"], 2),
        "未稅小計": round(untaxed, 2),
        "稅額": round(tax, 2),
        "含稅總額": round(total, 2)
    }



# --- 儲存當月發票紀錄（覆蓋當月） ---
def save_monthly_summary(
    device_id,
    month_int,

    total_curr_color_a3,
    total_curr_color,
    total_curr_bw,

    last_color_a3,
    last_color,
    last_bw,

    calc_result
):
    """
    所有張數與金額皆為「已區分」版本
    """

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    # --- 使用量計算（三種完全獨立） ---
    color_a3_usage = max(0, total_curr_color_a3 - last_color_a3)
    color_usage    = max(0, total_curr_color - last_color)
    bw_usage       = max(0, total_curr_bw - last_bw)

    c.execute("""
        INSERT OR REPLACE INTO billing_summary (
            device_id,
            month,
            year,

            color_a3_total,
            color_total,
            bw_total,

            color_a3_usage,
            color_usage,
            bw_usage,

            color_a3_bill_usage,
            color_bill_usage,
            bw_bill_usage,

            color_a3_amount,
            color_amount,
            bw_amount,

            monthly_rent,
            untaxed_subtotal,
            tax_amount,
            total_with_tax
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        device_id,
        month_int,
        datetime.now().year,
        total_curr_color_a3,
        total_curr_color,
        total_curr_bw,

        color_a3_usage,
        color_usage,
        bw_usage,

        calc_result.get("彩色A3計費張數", 0),
        calc_result.get("彩色計費張數", 0),
        calc_result.get("黑白計費張數", 0),

        calc_result.get("彩色A3金額", 0),
        calc_result.get("彩色金額", 0),
        calc_result.get("黑白金額", 0),

        calc_result.get("月租金", 0),
        calc_result.get("未稅小計", 0),
        calc_result.get("稅額", 0),
        calc_result.get("含稅總額", 0)
    ))

    conn.commit()
    conn.close()


# --- 讀取 billing_summary（回傳 1..12 月陣列） ---
def load_billing_summary(device_id, year):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    c.execute("""
        SELECT
            month,

            color_a3_total,
            color_total,
            bw_total,

            color_a3_usage,
            color_usage,
            bw_usage,

            color_a3_bill_usage,
            color_bill_usage,
            bw_bill_usage,

            color_a3_amount,
            color_amount,
            bw_amount,

            monthly_rent,
            untaxed_subtotal,
            tax_amount,
            total_with_tax
        FROM billing_summary
        WHERE device_id = ? AND year = ?
    """, (device_id,year))

    rows = c.fetchall()
    conn.close()

    # 初始化 12 個月
    months = {
        m: {
            "color_a3_total": "",
            "color_total": "",
            "bw_total": "",

            "color_a3_usage": "",
            "color_usage": "",
            "bw_usage": "",

            "color_a3_bill_usage": "",
            "color_bill_usage": "",
            "bw_bill_usage": "",

            "color_a3_amount": "",
            "color_amount": "",
            "bw_amount": "",

            "monthly_rent": "",
            "untaxed_subtotal": "",
            "tax_amount": "",
            "total_with_tax": ""
        }
        for m in range(1, 13)
    }

    for r in rows:
        m = int(r[0])
        months[m] = {
            "color_a3_total": r[1],
            "color_total": r[2],
            "bw_total": r[3],

            "color_a3_usage": r[4],
            "color_usage": r[5],
            "bw_usage": r[6],

            "color_a3_bill_usage": r[7],
            "color_bill_usage": r[8],
            "bw_bill_usage": r[9],

            "color_a3_amount": r[10],
            "color_amount": r[11],
            "bw_amount": r[12],

            "monthly_rent": r[13],
            "untaxed_subtotal": r[14],
            "tax_amount": r[15],
            "total_with_tax": r[16]
        }

    return months


# --- 主頁面路由 ---
@bp.route("/", methods=["GET", "POST"])
def index():
    message = request.args.get("message", "")
    contract, customer, result = None, None, None
    contra_text = ""
    last_color_a3, last_color, last_bw, last_time = 0, 0, 0, ""
    matches = []
    related_devices = []

    # ✅ 取得選擇的抄表年月（POST 表單或 GET 參數）
    selected_month = int(request.form.get("selected_month") or request.args.get("selected_month") or datetime.now().month)
    selected_year  = int(request.form.get("selected_year")  or request.args.get("selected_year")  or datetime.now().year)
    
    # --- 計算前月年與月 ---
    def get_prev_month_year(year, month):
        if month == 1:
            return year - 1, 12
        return year, month - 1

    prev_year, prev_month = get_prev_month_year(selected_year, selected_month)

    # --- 共用：取得 contract 與 customer ---
    def load_device_data(device_id):
        c, ct = get_contract(device_id)
        cu = get_customer(device_id)
        return c, ct, cu

    if request.method == "POST":
        mode = request.form.get("mode")
        device_id = request.form.get("device_id", "").strip()
        keyword = device_id

        if mode == "query":
            contract, contra_text = get_contract(keyword)
            customer = get_customer(keyword)
            if not contract:
                matches = search_customers_by_name(keyword)
                message = f"🔍 找到 {len(matches)} 筆相符客戶" if matches else f"❌ 找不到設備或客戶：{keyword}"
            else:
                last_color_a3, last_color, last_bw, last_time = get_last_counts(keyword, prev_year, prev_month)
                related_devices = get_related_devices(keyword)

        elif mode == "calculate":
            contract, contra_text = get_contract(device_id)
            customer = get_customer(device_id)
            if contract:
                related_devices = get_related_devices(device_id)
                total_last_color_a3 = total_last_color = total_last_bw = 0
                total_curr_color_a3 = total_curr_color = total_curr_bw = 0

                for dev in related_devices:
                    last_a3, last_c, last_b, _ = get_last_counts(dev, prev_year, prev_month)
                    total_last_color_a3 += last_a3
                    total_last_color += last_c
                    total_last_bw += last_b

                    val_a3 = request.form.get(f"curr_color_a3_{dev}")
                    val_c = request.form.get(f"curr_color_{dev}")
                    val_b = request.form.get(f"curr_bw_{dev}")

                    if val_c is None:
                        total_curr_color_a3 += int(request.form.get("curr_color_a3", 0))
                        total_curr_color += int(request.form.get("curr_color", 0))
                        total_curr_bw += int(request.form.get("curr_bw", 0))
                    else:
                        total_curr_color_a3 += int(val_a3 or 0)
                        total_curr_color += int(val_c or 0)
                        total_curr_bw += int(val_b or 0)

                result = calculate(
                    contract,
                    total_curr_color_a3,
                    total_curr_color,
                    total_curr_bw,
                    total_last_color_a3,
                    total_last_color,
                    total_last_bw
                )

                for dev in related_devices:
                    insert_usage(
                        dev,
                        int(request.form.get(f"curr_color_a3_{dev}", 0)),
                        int(request.form.get(f"curr_color_{dev}", 0)),
                        int(request.form.get(f"curr_bw_{dev}", 0))
                    )

                save_monthly_summary(
                    device_id,
                    selected_month,
                    total_curr_color_a3,
                    total_curr_color,
                    total_curr_bw,
                    total_last_color_a3,
                    total_last_color,
                    total_last_bw,
                    result
                )
                message = f"✅ {device_id} 的抄表與金額已儲存至 {selected_month} 月"
            else:
                message = f"❌ 找不到設備 {device_id}"

        elif mode in ["update_contract", "update_customer", "delete_customer", "new_customer"]:
            # --- Google Sheet 客戶與契約工作表 ---
            customers_ws = get_person_worksheet("customers")
            contracts_ws = get_person_worksheet("contracts")

            if mode == "update_contract":
                contract_data = {
                    "monthly_rent": float(request.form.get("monthly_rent") or 0),
                    "color_unit_price": float(request.form.get("color_unit_price") or 0),
                    "bw_unit_price": float(request.form.get("bw_unit_price") or 0),
                    "color_giveaway": to_int(request.form.get("color_giveaway")),
                    "bw_giveaway": to_int(request.form.get("bw_giveaway")),
                    "color_error_rate": float(request.form.get("color_error_rate") or 0),
                    "bw_error_rate": float(request.form.get("bw_error_rate") or 0),
                    "color_basic": to_int(request.form.get("color_basic")),
                    "bw_basic": to_int(request.form.get("bw_basic")),
                    "color_a3_unit_price": float(request.form.get("color_a3_unit_price") or 0),
                    "color_a3_giveaway": to_int(request.form.get("color_a3_giveaway")),
                    "color_a3_error_rate": float(request.form.get("color_a3_error_rate") or 0),
                    "color_a3_basic": to_int(request.form.get("color_a3_basic")),
                    "tax_type": request.form.get("tax_type", "含稅"),
                    "contra": request.form.get("contra", "").strip()
                }

                update_contract(device_id, contract_data)
                return redirect(url_for("billing.index", device_id=device_id, message="✅ 契約條件已更新"))

            elif mode == "update_customer":
                customer_data = {
                    "customer_name": request.form.get("customer_name", "").strip(),
                    "device_number": request.form.get("device_number", "").strip(),
                    "machine_model": request.form.get("machine_model", "").strip(),
                    "tax_id": request.form.get("tax_id", "").strip(),
                    "install_address": request.form.get("install_address", "").strip(),
                    "service_person": request.form.get("service_person", "").strip(),
                    "contract_number": request.form.get("contract_number", "").strip(),
                    "contract_start": request.form.get("contract_start", "").strip(),
                    "contract_end": request.form.get("contract_end", "").strip(),
                    "pm": request.form.get("pm", "").strip()  # 如果有保養週期
                }
                update_customer(device_id, customer_data)
                return redirect(url_for("billing.index", device_id=device_id, message="✅ 客戶資料已更新"))

            elif mode == "delete_customer":
                delete_customer(device_id)
                message = f"🗑 已刪除客戶（設備編號：{device_id}）"

            elif mode == "new_customer":
                old_id = request.form.get("device_id")
                new_id = request.form.get("device_id_new", "").strip()

                old_customer = get_customer(old_id)
                old_contract, _ = get_contract(old_id)

                # 檢查原始資料
                if not old_customer or not old_contract:
                    message = "❌ 找不到原始客戶或契約資料，無法建檔。"

                # 檢查新設備編號
                elif not new_id:
                    message = "⚠️ 請輸入新設備編號。"

                # 檢查新ID是否已存在
                elif get_customer(new_id):
                    message = "❌ 此設備編號已存在，請使用不同編號。"

                else:
                    # ===== 建立新客戶資料 =====
                    new_customer_data = {
                        "device_id": new_id,
                        "customer_name": request.form.get("customer_name", "").strip(),
                        "device_number": request.form.get("device_number", "").strip(),
                        "machine_model": request.form.get("machine_model", "").strip(),
                        "tax_id": request.form.get("tax_id", "").strip(),
                        "install_address": request.form.get("install_address", "").strip(),
                        "service_person": request.form.get("service_person", "").strip(),
                        "contract_number": request.form.get("contract_number", "").strip(),
                        "contract_start": request.form.get("contract_start", "").strip(),
                        "contract_end": request.form.get("contract_end", "").strip(),
                        "pm": request.form.get("pm", "").strip()
                    }

                    # ===== 建立新契約資料（複製舊契約） =====
                    new_contract_data = {
                        "device_id": new_id,
                        "monthly_rent": old_contract.get("monthly_rent", 0),
                        "color_unit_price": old_contract.get("color_unit_price", 0),
                        "bw_unit_price": old_contract.get("bw_unit_price", 0),
                        "color_giveaway": old_contract.get("color_giveaway", 0),
                        "bw_giveaway": old_contract.get("bw_giveaway", 0),
                        "color_error_rate": old_contract.get("color_error_rate", 0),
                        "bw_error_rate": old_contract.get("bw_error_rate", 0),
                        "color_basic": old_contract.get("color_basic", 0),
                        "bw_basic": old_contract.get("bw_basic", 0),
                        "color_a3_unit_price": old_contract.get("color_a3_unit_price", 0),
                        "color_a3_giveaway": old_contract.get("color_a3_giveaway", 0),
                        "color_a3_error_rate": old_contract.get("color_a3_error_rate", 0),
                        "color_a3_basic": old_contract.get("color_a3_basic", 0),
                        "tax_type": old_contract.get("tax_type", ""),
                        "contra": old_contract.get("contra", "")
                    }

                    # ===== 寫入 Google Sheet =====
                    ok1 = insert_customer(new_id, new_customer_data)
                    ok2 = insert_contract(new_id, new_contract_data)

                    if not ok1 or not ok2:
                        message = "❌ 新客戶建檔失敗（Google Sheet 寫入錯誤）"
                    else:
                        return redirect(
                            url_for(
                                "billing.index",
                                device_id=new_id,
                                message="✅ 新客戶建檔成功！"
                            )
                        )



    # GET 直接帶 device_id
    elif request.args.get("device_id"):
        q_device = request.args.get("device_id")
        contract, contra_text = get_contract(q_device)
        customer = get_customer(q_device)
        if contract:
            prev_year, prev_month = get_prev_month_year(selected_year, selected_month)

            # 前次張數只拿 4 個
            last_color_a3, last_color, last_bw, last_time = get_last_counts(
                q_device, selected_year, selected_month
            )
        else:
            message = f"❌ 找不到設備 {q_device}"

    now = datetime.now()
    return render_template(
        "billing_index.html",
        billing_page=True,
        contract=contract,
        contra_text=contra_text,
        customer=customer,
        last_color=last_color,
        last_color_a3=last_color_a3,
        last_bw=last_bw,
        last_time=last_time,
        result=result,
        matches=matches,
        message=message,
        related_devices=related_devices,
        current_year=now.year,
        selected_year=selected_year,
        selected_month=selected_month
    )


# --- 顯示發票紀錄頁面（12 列，可選年份） ---
@bp.route("/invoice_log/<device_id>")
def invoice_log(device_id):
    selected_year = request.args.get("year", type=int) or datetime.now().year
    months = load_billing_summary(device_id, selected_year)  # dict keyed by 1..12
    return render_template(
        "invoice_log.html",
        device_id=device_id,
        months=months,
        selected_year=selected_year,
        current_year=datetime.now().year,
        billing_invoice_log=True
    )



# ================================================================
# 客戶總表 + 概況（summary）
# ================================================================
@bp.route('/mfp_summary')
def mfp_summary():
    keyword = request.args.get("keyword", "").strip()

    # =====================================
    # ① 改為讀取 GOOGLE SHEET：customers
    # =====================================
    ws = get_person_worksheet("customers")
    rows = ws.get_all_records()   # list of dict

    tables = rows.copy()

    # 🔹 將數字欄位轉整數，避免 round 報錯
    numeric_fields = ['pm', 'device_number', 'tax_id']
    for row in tables:
        for key in numeric_fields:
            val = row.get(key)
            if val not in (None, ""):
                try:
                    row[key] = int(float(val))
                except:
                    pass  # 轉型失敗就保留原值

    # 🔹 日期欄位格式化 YYYY/MM/DD
    for row in tables:
        for key in ['contract_start', 'contract_end']:
            val = row.get(key)
            if val:
                try:
                    dt = pd.to_datetime(val)
                    row[key] = dt.strftime("%Y/%m/%d")
                except:
                    pass

    # 🔹 合約結束距今天小於三個月加標記
    today = pd.Timestamp.today()
    for row in tables:
        val = row.get('contract_end')
        if val:
            try:
                end_date = pd.to_datetime(val)
                delta = (end_date - today).days
                row['_contract_end_alert'] = delta < 90
            except:
                row['_contract_end_alert'] = False
        else:
            row['_contract_end_alert'] = False

    # 🔍 關鍵字搜尋
    if keyword:
        keyword_lower = keyword.lower()
        tables = [
            r for r in tables
            if any(keyword_lower in str(v).lower() for v in r.values())
        ]

    # -------------------------
    # 讀 Excel 概況（區域台數 / 保養週期）
    # -------------------------
    xls = load_github_excel("MFP.xlsx")
    df_overview = pd.read_excel(xls, sheet_name='概況', header=0)

    # 區域台數：A1:R4
    df_area = pd.read_excel(
        xls,
        sheet_name='概況',
        header=0,
        usecols="A:R",
        nrows=4
    ).infer_objects()  # 避免 FutureWarning

    # 保養週期：A6:R12
    df_cycle = pd.read_excel(
        xls,
        sheet_name='概況',
        header=0,
        usecols="A:R",
        skiprows=5,  # 從第6列開始
        nrows=7      # 6~12列
    ).infer_objects()  # 避免 FutureWarning

    version = current_app.config['VERSION_TIME']

    return render_template(
        'billing_mfp_summary.html',
        tables=tables,
        table_area = df_area.to_html(index=False, classes="table table-bordered"),
        table_cycle = df_cycle.to_html(index=False, classes="table table-bordered"),
        version=version,
        keyword=keyword,
        billing_mfp_summary=True
    )



# ================================================================
# 讀取備註（person_page 用）
# ================================================================
def load_person_remarks(sheet_name):
    ws = get_person_worksheet(sheet_name)
    rows = ws.get_all_records()

    return {
        r["設備代號"]: {
            "remark": str(r.get("備註", "") or ""),
            "method": str(r.get("抄表方式", "") or "")
        }
        for r in rows
    }
    
    
# ================================================================
# 寫回備註（AJAX API 用）
# ================================================================
def upsert_person_field(sheet_name, device_id, field, value):
    ws = get_person_worksheet(sheet_name)
    header = ws.row_values(1)

    device_col = header.index("設備代號") + 1
    target_col = {
        "remark": header.index("備註") + 1,
        "method": header.index("抄表方式") + 1
    }[field]

    records = ws.get_all_records()

    for idx, r in enumerate(records, start=2):
        if str(r.get("設備代號")).strip() == str(device_id):
            ws.update_cell(idx, target_col, value)
            return

    ws.append_row([
        str(device_id),
        str(value) if field == "remark" else "",
        str(value) if field == "method" else ""
    ])
    
    
# ================================================================
# 新增 API
# ================================================================
@bp.route("/save_person_field", methods=["POST"])
def save_person_field():
    data = request.json
    upsert_person_field(
        sheet_name=data["sheet"],
        device_id=data["device_id"],
        field=data["field"],
        value=data["value"]
    )
    return {"ok": True}


# ================================================================
# 讀取 GitHub / 本地 Excel（支援檔名參數）
# ================================================================
_cached_xls = None  # 快取字典 {'filename': '...', 'xls': pd.ExcelFile}

def load_github_excel(filename="MFP.xlsx"):
    """
    安全下載 GitHub RAW EXCEL（含快取與 fallback）
    filename: 可選，本地 fallback 使用的 Excel 檔名
    """
    import requests
    from io import BytesIO
    import pandas as pd

    global _cached_xls

    if _cached_xls and _cached_xls['filename'] == filename:
        return _cached_xls['xls']

    try:
        resp = requests.get(GITHUB_XLSX_URL, timeout=10)
        if resp.status_code != 200:
            raise Exception(f"HTTP {resp.status_code}")

        excel_bytes = BytesIO(resp.content)

        import zipfile
        if not zipfile.is_zipfile(excel_bytes):
            raise Exception("下載內容不是 Excel（不是 zip 格式）")

        xls = pd.ExcelFile(excel_bytes, engine="openpyxl")
        _cached_xls = {'filename': filename, 'xls': xls}
        return xls

    except Exception as e:
        print(f"⚠ GitHub Excel 載入失敗，改用本地 {filename}，原因：{e}")
        local_path = f"MFP/{filename}"  # 本地 fallback
        xls = pd.ExcelFile(local_path, engine="openpyxl")
        _cached_xls = {'filename': filename, 'xls': xls}
        return xls

# ================================================================
# 2️⃣ 人員個人資料頁（person）
# ================================================================
@bp.route("/person/<sheet>")
def person_page(sheet):
    keyword = request.args.get("keyword", "").strip()

    # --- 讀 GitHub MFP.xlsx 保留前兩區塊（Accordion） ---
    mfp_xls = load_github_excel("MFP.xlsx")
    df1 = pd.read_excel(mfp_xls, sheet_name=sheet, header=0, usecols="A:R", nrows=4)
    df2 = pd.read_excel(mfp_xls, sheet_name=sheet, header=0, usecols="A:R", skiprows=5, nrows=4)

    # --- 從 Google Sheet 讀取 customers ---
    ws = get_person_worksheet("customers")        # 取得 Worksheet
    rows = ws.get_all_records()                   # 轉成 list of dict
    all_customers = pd.DataFrame(rows)           # 轉成 DataFrame

    # 篩選該負責人的資料
    df3 = all_customers[
        all_customers["service_person"].astype(str).str.strip() == sheet
    ][["customer_name", "pm", "device_id"]].copy()

    # --- 從 output.xlsx 讀取 pm_date ---
    output_xls = load_github_excel("output.xlsx")
    df_pm = pd.read_excel("MFP/output.xlsx", sheet_name="customers", usecols="A:L", engine="openpyxl")
    df_pm["device_id"] = df_pm["device_id"].astype(str).str.strip()  # 確保設備代號一致

    # --- 從 Google Sheet 讀取備註與抄表方式 ---
    gs_data = load_person_remarks(sheet)  # dict keyed by 設備代號

    # --- 表頭重新命名（Google Sheet -> 中文） ---
    df3 = df3.rename(columns={
        "customer_name": "客戶名稱",
        "pm": "保養週期",
        "device_id": "設備代號"
    })

    # --- 新增欄位 備註 / 抄表方式 / 最後保養日 ---
    df3["備註"] = ""
    df3["抄表方式"] = ""
    df3["最後保養日"] = ""

    # --- 合併 Google Sheet 備註資料 ---
    for idx, row in df3.iterrows():
        dev_id = str(row["設備代號"]).strip()
        if dev_id in gs_data:
            df3.at[idx, "備註"] = gs_data[dev_id].get("remark", "")
            df3.at[idx, "抄表方式"] = gs_data[dev_id].get("method", "")

    # --- 合併 output.xlsx pm_date 資料 ---
    for idx, row in df3.iterrows():
        dev_id = str(row["設備代號"]).strip()
        match = df_pm[df_pm["device_id"] == dev_id]
        if not match.empty and pd.notna(match.iloc[0]["pm_date"]):
            df3.at[idx, "最後保養日"] = pd.to_datetime(match.iloc[0]["pm_date"]).strftime("%Y-%m-%d")
        else:
            df3.at[idx, "最後保養日"] = ""

    # --- 對 df3["客戶名稱"] 套用顏色判斷 ---
    df3["客戶名稱"] = df3.apply(
        lambda r: color_overdue(r["客戶名稱"], r["最後保養日"], r["保養週期"]),
        axis=1
    )

    # ✅ 將所有 NaN 轉成空字串
    df3 = df3.fillna("")

    # --- 加上「項次」欄位 ---
    df3.insert(0, "項次", range(1, len(df3) + 1))

    # --- 調整欄位順序 ---
    df3 = df3[["項次", "客戶名稱", "備註", "保養週期", "最後保養日", "設備代號", "抄表方式"]]

    # --- 關鍵字過濾 ---
    if keyword:
        df3 = df3[df3.apply(lambda r: r.astype(str).str.contains(keyword, case=False, na=False).any(), axis=1)]

    # --- 傳給模板 ---
    return render_template(
        "tjw.html",
        table1=df1.to_html(index=False, classes="table table-bordered"),
        table2=df2.to_html(index=False, classes="table table-bordered"),
        df3=df3,  # ← Google Sheet + output.xlsx 資料
        page_name=sheet,
        keyword=keyword,
        billing_person=True
    )


@billing_bp.route('/get_last_counts', methods=['GET'])
def route_get_last_counts():
    device_id = request.args.get('device_id')
    
    # 1. 透過 request.args 正確抓取前端 GET 請求帶來的參數，並給予當前時間作為預設防呆
    selected_year = int(request.args.get('year', datetime.now().year))
    selected_month = int(request.args.get('month', datetime.now().month))
    
    # 2. 呼叫修正後的 get_prev_month_year
    prev_year, prev_month = get_prev_month_year(selected_year, selected_month)
    
    # 3. 接下來從資料庫（billing_summary 或 usage）撈取前次資料
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
        SELECT color_a3_total, color_total, bw_total 
        FROM billing_summary 
        WHERE device_id=? AND year=? AND month=?
    """, (device_id, prev_year, prev_month))
    row = c.fetchone()
    conn.close()
    
    if row:
        return jsonify({
            "color_a3": row[0] or 0,
            "color": row[1] or 0,
            "bw": row[2] or 0,
            "prev_year": prev_year,
            "prev_month": prev_month
        })
    else:
        return jsonify({
            "color_a3": 0,
            "color": 0,
            "bw": 0,
            "prev_year": prev_year,
            "prev_month": prev_month
        })

# ✅ 讓主程式 app.py 可以 import billing_bp
billing_bp = bp
