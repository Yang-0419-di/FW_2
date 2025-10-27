# === billing.py ===
from flask import Blueprint, render_template, request, redirect, url_for
import sqlite3
from datetime import datetime

bp = Blueprint("billing", __name__, url_prefix="/billing")
DB_FILE = "billing.db"


# --- 初始化資料庫 ---
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

    conn.commit()
    conn.close()


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
    
    def fmt_date(val):
        if not val:
            return ""
        try:
            return datetime.strptime(val, "%Y-%m-%d %H:%M:%S").strftime("%Y/%m/%d")
        except:
            return val  # 若不是這種格式就原樣回傳
    
    if row:
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
        return row[0] or 0, row[1] or 0, row[2] or ""
    return 0, 0, ""


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

    return {
        "彩色使用張數": used_color,
        "黑白使用張數": used_bw,
        "彩色計費張數": bill_color,
        "黑白計費張數": bill_bw,
        "彩色金額": round(color_amount, 2),
        "黑白金額": round(bw_amount, 2),
        "月租金": round(contract["monthly_rent"], 2),
        "未稅小計": int(round(untaxed)),
        "稅額": int(round(tax)),
        "含稅總額": int(round(total))
    }


# --- 主頁面路由 ---
@bp.route("/", methods=["GET", "POST"])
def index():
    message = request.args.get("message", "")
    contract, customer, result = None, None, None
    contra_text = ""
    last_color, last_bw, last_time = 0, 0, ""
    matches = []

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

        elif mode == "calculate":
            device_id = keyword
            contract, contra_text = get_contract(device_id)
            customer = get_customer(device_id)
            if contract:
                last_color, last_bw, last_time = get_last_counts(device_id)
                curr_color = int(request.form.get("curr_color", "0"))
                curr_bw = int(request.form.get("curr_bw", "0"))
                result = calculate(contract, curr_color, curr_bw, last_color, last_bw)
                insert_usage(device_id, curr_color, curr_bw)
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
        else:
            message = f"❌ 找不到設備 {q_device}"

    # ✅ 移出 return，讓所有情況都會回傳畫面
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
                           message=message)
                           
# ✅ 讓主程式 app.py 可以 import billing_bp
billing_bp = bp
