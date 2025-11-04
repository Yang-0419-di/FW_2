# === billing.py ===
from flask import Blueprint, render_template, request, redirect, url_for, abort
import sqlite3
from datetime import datetime

bp = Blueprint("billing", __name__, url_prefix="/billing")
DB_FILE = "billing.db"


# --- 初始化資料庫 ---
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()

    # 契約資料表
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

    # 發票紀錄表
    c.execute("""
        CREATE TABLE IF NOT EXISTS billing_summary (
            device_id TEXT,
            month INTEGER,
            color_total INTEGER,
            bw_total INTEGER,
            color_usage INTEGER,
            bw_usage INTEGER,
            color_bill_usage INTEGER,
            bw_bill_usage INTEGER,
            color_amount REAL,
            bw_amount REAL,
            monthly_rent REAL,
            untaxed_subtotal REAL,
            tax_amount REAL,
            total_with_tax REAL,
            PRIMARY KEY (device_id, month)
        )
    """)

    conn.commit()
    conn.close()


init_db()


# --- 公用查詢函數 ---
def get_contract(device_id):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT * FROM contracts WHERE device_id=?", (device_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return None, ""
    col_names = [d[0] for d in c.description]
    contract = dict(zip(col_names, row))
    conn.close()
    return contract, contract.get("contra", "")


def get_customer(device_id):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT * FROM customers WHERE device_id=?", (device_id,))
    row = c.fetchone()
    conn.close()
    if not row:
        return None
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
        "contract_end": row[9],
    }


def get_last_counts(device_id, month_int=None):
    """取得該設備指定月份（或最後一筆）的讀數"""
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    if month_int:
        month_str = f"{datetime.now().year}{month_int:02d}"
        c.execute(
            "SELECT color_count, bw_count, timestamp FROM usage WHERE device_id=? AND month=? ORDER BY id DESC LIMIT 1",
            (device_id, month_str),
        )
    else:
        c.execute(
            "SELECT color_count, bw_count, timestamp FROM usage WHERE device_id=? ORDER BY id DESC LIMIT 1",
            (device_id,),
        )
    row = c.fetchone()
    conn.close()
    if row:
        return row[0] or 0, row[1] or 0, row[2] or ""
    return 0, 0, ""


def get_related_devices(device_id):
    """找出主機與合開子機"""
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT master_device_id FROM contracts WHERE device_id=?", (device_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return [device_id]
    master_id = row[0]
    if not master_id:
        c.execute("SELECT device_id FROM contracts WHERE master_device_id=?", (device_id,))
        subs = [r[0] for r in c.fetchall()]
        conn.close()
        return [device_id] + subs
    else:
        c.execute("SELECT device_id FROM contracts WHERE master_device_id=?", (master_id,))
        subs = [r[0] for r in c.fetchall()]
        conn.close()
        return [master_id] + subs


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
        "含稅總額": float(round(total, 2)),
    }


def insert_usage(device_id, month_int, color_count, bw_count):
    month_str = f"{datetime.now().year}{month_int:02d}"
    timestamp = datetime.now().strftime("%Y/%m/%d-%H:%M")
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute(
        "INSERT INTO usage (device_id, month, color_count, bw_count, timestamp) VALUES (?, ?, ?, ?, ?)",
        (device_id, month_str, color_count, bw_count, timestamp),
    )
    conn.commit()
    conn.close()


def save_monthly_summary(device_id, month_int, total_curr_color, total_curr_bw, last_color, last_bw, calc_result):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    color_usage = max(0, total_curr_color - last_color)
    bw_usage = max(0, total_curr_bw - last_bw)
    c.execute(
        """
        INSERT OR REPLACE INTO billing_summary (
            device_id, month, color_total, bw_total,
            color_usage, bw_usage,
            color_bill_usage, bw_bill_usage,
            color_amount, bw_amount, monthly_rent,
            untaxed_subtotal, tax_amount, total_with_tax
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """,
        (
            device_id,
            month_int,
            total_curr_color,
            total_curr_bw,
            color_usage,
            bw_usage,
            calc_result["彩色計費張數"],
            calc_result["黑白計費張數"],
            calc_result["彩色金額"],
            calc_result["黑白金額"],
            calc_result["月租金"],
            calc_result["未稅小計"],
            calc_result["稅額"],
            calc_result["含稅總額"],
        ),
    )
    conn.commit()
    conn.close()


# --- 主頁面路由 ---
@bp.route("/", methods=["GET", "POST"])
def index():
    message = request.args.get("message", "")
    contract, customer, result = None, None, None
    contra_text = ""
    last_color, last_bw, last_time = 0, 0, ""
    related_devices = []

    if request.method == "POST":
        mode = request.form.get("mode")
        keyword = request.form.get("device_id", "").strip()

        # 查詢客戶
        if mode == "search":
            contract, contra_text = get_contract(keyword)
            customer = get_customer(keyword)
            if contract:
                related_devices = get_related_devices(keyword)
                last_color, last_bw, last_time = get_last_counts(keyword)
            else:
                message = "查無此客戶或設備代號"

        # 計算並儲存抄表
        elif mode == "calculate":
            selected_month = int(request.form.get("selected_month", datetime.now().month))
            prev_month = 12 if selected_month == 1 else selected_month - 1

            contract, contra_text = get_contract(keyword)
            customer = get_customer(keyword)
            if contract:
                related_devices = get_related_devices(keyword)
                total_last_color, total_last_bw = 0, 0
                total_curr_color, total_curr_bw = 0, 0

                for dev in related_devices:
                    last_c, last_b, _ = get_last_counts(dev, prev_month)
                    total_last_color += last_c
                    total_last_bw += last_b
                    val_c = request.form.get(f"curr_color_{dev}")
                    val_b = request.form.get(f"curr_bw_{dev}")
                    if val_c is None or val_b is None:
                        total_curr_color += int(request.form.get("curr_color", "0"))
                        total_curr_bw += int(request.form.get("curr_bw", "0"))
                    else:
                        total_curr_color += int(val_c or 0)
                        total_curr_bw += int(val_b or 0)

                result = calculate(contract, total_curr_color, total_curr_bw, total_last_color, total_last_bw)
                for dev in related_devices:
                    val_c = request.form.get(f"curr_color_{dev}")
                    val_b = request.form.get(f"curr_bw_{dev}")
                    if val_c is None or val_b is None:
                        curr_c = int(request.form.get("curr_color", "0"))
                        curr_b = int(request.form.get("curr_bw", "0"))
                    else:
                        curr_c = int(val_c or 0)
                        curr_b = int(val_b or 0)
                    insert_usage(dev, selected_month, curr_c, curr_b)

                save_monthly_summary(keyword, selected_month, total_curr_color, total_curr_bw, total_last_color, total_last_bw, result)
                message = f"✅ 已計算 {selected_month} 月抄表"

    return render_template(
        "billing_index.html",
        billing_page=True,
        contract=contract,
        contra_text=contra_text,
        customer=customer,
        last_color=last_color,
        last_bw=last_bw,
        last_time=last_time,
        result=result,
        message=message,
        related_devices=related_devices,
    )


@bp.route("/invoice_log/<device_id>")
def invoice_log(device_id):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute(
        "SELECT month, color_total, bw_total, color_usage, bw_usage, color_bill_usage, bw_bill_usage, color_amount, bw_amount, monthly_rent, untaxed_subtotal, tax_amount, total_with_tax FROM billing_summary WHERE device_id=? ORDER BY month",
        (device_id,),
    )
    rows = c.fetchall()
    conn.close()
    months = {m: {} for m in range(1, 13)}
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
            "total_with_tax": r[12],
        }
    return render_template("invoice_log.html", device_id=device_id, months=months)


billing_bp = bp
