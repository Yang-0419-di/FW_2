# === billing.py ===
from flask import Blueprint, render_template, request, redirect, url_for, abort
import sqlite3
from datetime import datetime

bp = Blueprint("billing", __name__, url_prefix="/billing")
DB_FILE = "billing.db"


# --- 初始化資料庫（完整，不略） ---
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    # 契約資料表（含稅別欄位與 contra）
    c.execute("""
        CREATE TABLE IF NOT EXISTS contracts (
            device_id TEXT PRIMARY KEY,
            monthly_rent REAL,
            color_unit_price REAL,
            bw_unit_price REAL,
            color_giveaway INTEGER,
            bw_giveaway INTEGER,
            color_error_rate REAL,
            bw_error_rate REAL,
            color_basic INTEGER,
            bw_basic INTEGER,
            tax_type TEXT DEFAULT '含稅',
            contra TEXT DEFAULT '',
            master_device_id TEXT DEFAULT ''
        )
    """)

    # 抄表記錄表
    c.execute("""
        CREATE TABLE IF NOT EXISTS usage (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            device_id TEXT,
            month TEXT,
            color_count INTEGER,
            bw_count INTEGER,
            timestamp TEXT
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
            color_total INTEGER,     -- 本月抄表 彩色總張數（若合開則為合計）
            bw_total INTEGER,        -- 本月抄表 黑白總張數
            color_usage INTEGER,     -- 當月使用彩色 = 本月 - 上月 (delta)
            bw_usage INTEGER,        -- 當月使用黑白 = 本月 - 上月 (delta)
            color_bill_usage INTEGER,-- 彩色計費張數（扣贈送、誤印率、基本張數）
            bw_bill_usage INTEGER,   -- 黑白計費張數
            color_amount REAL,       -- 彩色金額
            bw_amount REAL,          -- 黑白金額
            monthly_rent REAL,       -- 月租金
            untaxed_subtotal REAL,   -- 未稅小計（彩色金額+黑白金額+月租）
            tax_amount REAL,         -- 稅額
            total_with_tax REAL,     -- 含稅總額
            PRIMARY KEY (device_id, month)
        )
    """)

    conn.commit()
    conn.close()


# 呼叫初始化以確保資料表存在（可在應用啟動時呼叫一次）
init_db()


# --- 查詢契約 ---
def get_contract(device_id):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT * FROM contracts WHERE device_id=?", (device_id,))
    contract_row = c.fetchone()
    contra_text = ""

    if contract_row:
        col_names = [desc[0] for desc in c.description]
        contract_dict = dict(zip(col_names, contract_row))
        contra_text = contract_dict.get("contra", "")
    else:
        contract_dict = None

    conn.close()
    return contract_dict, contra_text


# --- 查詢客戶資料 ---
def get_customer(device_id):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT * FROM customers WHERE device_id=?", (device_id,))
    row = c.fetchone()
    conn.close()

    if row:
        # row order matches CREATE TABLE
        return {
            "device_id": row[0],
            "customer_name": row[1],
            "device_number": row[2],
            "machine_model": row[3],
            "tax_id": row[4],
            "install_address": row[5],
            "service_person": row[6],
            "contract_number": row[7],
            "contract_start": row[8],
            "contract_end": row[9]
        }
    return None


# --- 模糊搜尋客戶名稱 ---
def search_customers_by_name(keyword):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
        SELECT device_id, customer_name
        FROM customers
        WHERE customer_name LIKE ?
    """, (f"%{keyword}%",))
    rows = c.fetchall()
    conn.close()
    return [{"device_id": r[0], "customer_name": r[1]} for r in rows]


# --- 查詢最後抄表 ---
def get_last_counts(device_id):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT color_count, bw_count, timestamp FROM usage WHERE device_id=? ORDER BY id DESC LIMIT 1", (device_id,))
    row = c.fetchone()
    conn.close()
    if row:
        # 若為 None，返回 0
        return row[0] or 0, row[1] or 0, row[2] or ""
    return 0, 0, ""


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
def insert_usage(device_id, color_count, bw_count):
    month = datetime.now().strftime("%Y%m")
    timestamp = datetime.now().strftime("%Y/%m/%d-%H:%M")
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("INSERT INTO usage (device_id, month, color_count, bw_count, timestamp) VALUES (?, ?, ?, ?, ?)",
              (device_id, month, color_count, bw_count, timestamp))
    conn.commit()
    conn.close()


# --- 計算邏輯 ---
def calculate(contract, curr_color, curr_bw, last_color, last_bw):
    used_color = max(0, curr_color - last_color)
    used_bw = max(0, curr_bw - last_bw)

    bill_color = max(0, used_color - contract["color_giveaway"])
    bill_bw = max(0, used_bw - contract["bw_giveaway"])

    bill_color = int(round(bill_color * (1 - contract["color_error_rate"])))
    bill_bw = int(round(bill_bw * (1 - contract["bw_error_rate"])))

    if contract["color_basic"] > 0:
        bill_color = max(contract["color_basic"], bill_color)
    if contract["bw_basic"] > 0:
        bill_bw = max(contract["bw_basic"], bill_bw)

    color_amount = bill_color * contract["color_unit_price"]
    bw_amount = bill_bw * contract["bw_unit_price"]
    subtotal = contract["monthly_rent"] + color_amount + bw_amount

    tax_rate = 0.05
    if contract.get("tax_type") == "未稅":
        tax = subtotal * tax_rate
        total = subtotal + tax
        untaxed = subtotal
    else:
        total = subtotal
        untaxed = subtotal / (1 + tax_rate)
        tax = total - untaxed

    # 傳回詳細欄位（中文鍵名與之前一致）
    return {
        "彩色使用張數": used_color,
        "黑白使用張數": used_bw,
        "彩色計費張數": bill_color,
        "黑白計費張數": bill_bw,
        "彩色金額": round(color_amount, 2),
        "黑白金額": round(bw_amount, 2),
        "月租金": round(contract["monthly_rent"], 2),
        "未稅小計": float(round(untaxed, 2)),
        "稅額": float(round(tax, 2)),
        "含稅總額": float(round(total, 2))
    }


# --- 儲存當月發票紀錄（覆蓋當月） ---
def save_monthly_summary(device_id, month_int, total_curr_color, total_curr_bw, last_color, last_bw, calc_result):
    """
    device_id: str
    month_int: 1..12
    total_curr_color/ curr_bw: 本月抄表（合開合計）
    last_color/ last_bw: 上月合計
    calc_result: calculate(...) 回傳的 dict
    """
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    color_usage = max(0, total_curr_color - last_color)
    bw_usage = max(0, total_curr_bw - last_bw)

    c.execute('''
        INSERT OR REPLACE INTO billing_summary (
            device_id, month, color_total, bw_total,
            color_usage, bw_usage,
            color_bill_usage, bw_bill_usage,
            color_amount, bw_amount, monthly_rent,
            untaxed_subtotal, tax_amount, total_with_tax
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        device_id,
        month_int,
        total_curr_color,
        total_curr_bw,
        color_usage,
        bw_usage,
        calc_result.get("彩色計費張數", 0),
        calc_result.get("黑白計費張數", 0),
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
def load_billing_summary(device_id):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('SELECT month, color_total, bw_total, color_usage, bw_usage, color_bill_usage, bw_bill_usage, color_amount, bw_amount, monthly_rent, untaxed_subtotal, tax_amount, total_with_tax FROM billing_summary WHERE device_id=?', (device_id,))
    rows = c.fetchall()
    conn.close()

    # 初始化 12 個月的空值
    months = {m: {
        "color_total": "",
        "bw_total": "",
        "color_usage": "",
        "bw_usage": "",
        "color_bill_usage": "",
        "bw_bill_usage": "",
        "color_amount": "",
        "bw_amount": "",
        "monthly_rent": "",
        "untaxed_subtotal": "",
        "tax_amount": "",
        "total_with_tax": ""
    } for m in range(1, 13)}

    for r in rows:
        m = int(r[0])
        months[m] = {
            "color_total": r[1],
            "bw_total": r[2],
            "color_usage": r[3],
            "bw_usage": r[4],
            "color_bill_usage": r[5],
            "bw_bill_usage": r[6],
            "color_amount": r[7],
            "bw_amount": r[8],
            "monthly_rent": r[9],
            "untaxed_subtotal": r[10],
            "tax_amount": r[11],
            "total_with_tax": r[12]
        }

    return months


# --- 主頁面路由 ---
@bp.route("/", methods=["GET", "POST"])
def index():
    message = request.args.get("message", "")
    contract, customer, result = None, None, None
    contra_text = ""
    last_color, last_bw, last_time = 0, 0, ""
    matches = []
    related_devices = []

    if request.method == "POST":
        mode = request.form.get("mode")
        keyword = request.form.get("device_id", "").strip()

        # 模糊查詢客戶名稱
        if mode == "query":
            contract, contra_text = get_contract(keyword)
            customer = get_customer(keyword)
            if not contract:
                matches = search_customers_by_name(keyword)
                if matches:
                    message = f"🔍 找到 {len(matches)} 筆相符客戶"
                else:
                    message = f"❌ 找不到設備或客戶：{keyword}"
            else:
                last_color, last_bw, last_time = get_last_counts(keyword)
                related_devices = get_related_devices(keyword)

        elif mode == "calculate":
            device_id = keyword
            contract, contra_text = get_contract(device_id)
            customer = get_customer(device_id)
            if contract:
                # 合開群組
                related_devices = get_related_devices(device_id)

                # 合併所有設備的上次讀數 & 當前讀數
                total_last_color = 0
                total_last_bw = 0
                total_curr_color = 0
                total_curr_bw = 0

                # 讀取群組上次/當月數據
                for dev in related_devices:
                    last_c, last_b, _ = get_last_counts(dev)
                    total_last_color += last_c
                    total_last_bw += last_b

                    # 前端表單欄位名稱： curr_color_{device_id}
                    # 若合開群組則表單應有每台的輸入；若沒有則前端單機 input 名稱為 curr_color / curr_bw
                    val_c = request.form.get(f"curr_color_{dev}")
                    val_b = request.form.get(f"curr_bw_{dev}")
                    if val_c is None or val_b is None:
                        # 兼容單機表單欄位
                        total_curr_color += int(request.form.get("curr_color", "0"))
                        total_curr_bw += int(request.form.get("curr_bw", "0"))
                    else:
                        total_curr_color += int(val_c or 0)
                        total_curr_bw += int(val_b or 0)

                # 計算差異
                delta_color = total_curr_color - total_last_color
                delta_bw = total_curr_bw - total_last_bw

                # 套用主機的契約條件計算總金額（使用主機的 contract）
                result = calculate(contract, total_curr_color, total_curr_bw, total_last_color, total_last_bw)

                # 寫入每台機的抄表（保持原行為）
                for dev in related_devices:
                    val_c = request.form.get(f"curr_color_{dev}")
                    val_b = request.form.get(f"curr_bw_{dev}")
                    if val_c is None or val_b is None:
                        curr_c = int(request.form.get("curr_color", "0"))
                        curr_b = int(request.form.get("curr_bw", "0"))
                    else:
                        curr_c = int(val_c or 0)
                        curr_b = int(val_b or 0)
                    insert_usage(dev, curr_c, curr_b)

                # ✅ 將本次計算結果存入 billing_summary（以當前月份為 key，若已有則覆蓋）
                now_month = datetime.now().month
                save_monthly_summary(device_id, now_month, total_curr_color, total_curr_bw, total_last_color, total_last_bw, result)

            else:
                message = f"❌ 找不到設備 {device_id}"


        elif mode == "update_contract":
            device_id = keyword
            fields = {
                "monthly_rent": float(request.form.get("monthly_rent", "0") or 0),
                "color_unit_price": float(request.form.get("color_unit_price", "0") or 0),
                "bw_unit_price": float(request.form.get("bw_unit_price", "0") or 0),
                "color_giveaway": int(request.form.get("color_giveaway", "0") or 0),
                "bw_giveaway": int(request.form.get("bw_giveaway", "0") or 0),
                "color_error_rate": float(request.form.get("color_error_rate", "0") or 0),
                "bw_error_rate": float(request.form.get("bw_error_rate", "0") or 0),
                "color_basic": int(request.form.get("color_basic", "0") or 0),
                "bw_basic": int(request.form.get("bw_basic", "0") or 0),
                "tax_type": request.form.get("tax_type", "含稅"),
            }
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("""
                UPDATE contracts SET
                    monthly_rent=?, color_unit_price=?, bw_unit_price=?,
                    color_giveaway=?, bw_giveaway=?, color_error_rate=?, bw_error_rate=?,
                    color_basic=?, bw_basic=?, tax_type=?
                WHERE device_id=?""",
                (*fields.values(), device_id))
            conn.commit()
            conn.close()
            return redirect(url_for("billing.index", device_id=device_id, message="✅ 契約條件已更新"))

        elif mode == "update_customer":
            device_id = keyword
            fields = {
                "customer_name": request.form.get("customer_name", "").strip(),
                "device_number": request.form.get("device_number", "").strip(),
                "machine_model": request.form.get("machine_model", "").strip(),
                "tax_id": request.form.get("tax_id", "").strip(),
                "install_address": request.form.get("install_address", "").strip(),
                "service_person": request.form.get("service_person", "").strip(),
                "contract_number": request.form.get("contract_number", "").strip(),
                "contract_start": request.form.get("contract_start", "").strip(),
                "contract_end": request.form.get("contract_end", "").strip(),
            }
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("""
                UPDATE customers SET
                    customer_name=?, device_number=?, machine_model=?, tax_id=?,
                    install_address=?, service_person=?, contract_number=?,
                    contract_start=?, contract_end=?
                WHERE device_id=?
            """, (*fields.values(), device_id))
            conn.commit()
            conn.close()
            return redirect(url_for("billing.index", device_id=device_id, message="✅ 客戶資料已更新"))

        elif mode == "delete_customer":
            device_id = request.form.get("device_id")
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            # 刪除客戶資料
            c.execute("DELETE FROM customers WHERE device_id=?", (device_id,))
            # 同時刪除該客戶的契約資料
            c.execute("DELETE FROM contracts WHERE device_id=?", (device_id,))
            # （可選）刪除該客戶的抄表資料
            c.execute("DELETE FROM usage WHERE device_id=?", (device_id,))
            # （可選）刪除該客戶的 billing_summary 紀錄
            c.execute("DELETE FROM billing_summary WHERE device_id=?", (device_id,))
            conn.commit()
            conn.close()
            message = f"🗑 已刪除客戶（設備編號：{device_id}）"

        elif mode == "new_customer":
            old_id = request.form.get("device_id")
            new_id = request.form.get("device_id_new", "").strip()

            # 取得舊客戶資料與契約條件
            old_customer = get_customer(old_id)
            old_contract, _ = get_contract(old_id)

            if not old_customer or not old_contract:
                message = f"❌ 找不到原始客戶或契約資料，無法建檔。"
            elif not new_id:
                message = "⚠️ 請輸入新設備編號。"
            else:
                # 收集新客戶資料
                new_fields = {
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
                }

                conn = sqlite3.connect(DB_FILE)
                c = conn.cursor()

                # 🔹 新增客戶資料
                c.execute("""
                    INSERT INTO customers (
                        device_id, customer_name, device_number, machine_model,
                        tax_id, install_address, service_person,
                        contract_number, contract_start, contract_end
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, tuple(new_fields.values()))

                # 🔹 複製原契約條件
                c.execute("""
                    INSERT INTO contracts (
                        device_id, monthly_rent, color_unit_price, bw_unit_price,
                        color_giveaway, bw_giveaway, color_error_rate, bw_error_rate,
                        color_basic, bw_basic, tax_type, contra
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    new_id,
                    old_contract["monthly_rent"], old_contract["color_unit_price"], old_contract["bw_unit_price"],
                    old_contract["color_giveaway"], old_contract["bw_giveaway"],
                    old_contract["color_error_rate"], old_contract["bw_error_rate"],
                    old_contract["color_basic"], old_contract["bw_basic"],
                    old_contract["tax_type"], old_contract.get("contra", "")
                ))

                conn.commit()
                conn.close()

                return redirect(url_for("billing.index", device_id=new_id, message="✅ 新客戶建檔成功！"))

    elif request.args.get("device_id"):
        q_device = request.args.get("device_id")
        contract, contra_text = get_contract(q_device)
        customer = get_customer(q_device)
        if contract:
            last_color, last_bw, last_time = get_last_counts(q_device)
            related_devices = get_related_devices(q_device)
        else:
            message = f"❌ 找不到設備 {q_device}"

    # ✅ 統一回傳畫面
    return render_template("billing_index.html",
                           billing_page=True,
                           contract=contract,
                           contra_text=contra_text,
                           customer=customer,
                           last_color=last_color,
                           last_bw=last_bw,
                           last_time=last_time,
                           result=result,
                           matches=matches,
                           message=message,
                           related_devices=related_devices)


# --- 顯示發票紀錄頁面（12 列） ---
@bp.route("/invoice_log/<device_id>")
def invoice_log(device_id):
    months = load_billing_summary(device_id)  # dict keyed by 1..12
    # 傳給模板：months 為 dict，模板會用 1..12 月遍歷
    return render_template("invoice_log.html", device_id=device_id, months=months)


# ✅ 讓主程式 app.py 可以 import billing_bp
billing_bp = bp
