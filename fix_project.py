#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_project.py (v3) — Flask 計費系統 / Excel 自動化腳本 / 前端 templates 的缺陷修正工具

用法（在專案根目錄，也就是 app.py 所在資料夾執行）：
    python fix_project.py              # 套用所有修正（會先備份）
    python fix_project.py --dry-run    # 只預覽，不寫檔
    python fix_project.py --add-sri    # （另外執行，需連網）替 CDN 的 Bootstrap/FullCalendar 補 integrity
    python fix_project.py --restore _fix_backup/20261027_153000   # 還原備份
    python fix_project.py D:\\flask2    # 指定專案路徑

特性：
  * 可重複執行：已套用的修正會顯示「已套用」並略過
  * 套用前先備份到 _fix_backup/<時間>/，保留原檔 CRLF / BOM
  * 每個檔案寫入前先做語法檢查，語法錯誤則不寫入該檔
  * 找不到比對字串（檔案已被你改過）時只警告，不會硬改
"""
import argparse
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

PATCHES = []  # dict(file, desc, kind, old/new 或 pattern/repl)


def patch(file, desc, old, new):
    PATCHES.append(dict(file=file, desc=desc, kind="str", old=old, new=new))


def rpatch(file, desc, pattern, repl):
    PATCHES.append(dict(file=file, desc=desc, kind="re", pattern=pattern, repl=repl))


# =====================================================================
# app.py
# =====================================================================
patch("app.py", "匯入 time（Excel 快取 TTL 用）",
      "import os\nimport sqlite3\n",
      "import os\nimport sqlite3\nimport time\n")

patch("app.py", "Session Cookie 加固 + 管理員帳號改由環境變數設定 + SECRET_KEY 警告",
      "app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(minutes=10)\n",
      "app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(minutes=10)\n"
      "app.config['SESSION_COOKIE_HTTPONLY'] = True\n"
      "app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'\n"
      "if os.environ.get('SECRET_KEY') is None:\n"
      "    print('⚠️ 未設定環境變數 SECRET_KEY，正在使用不安全的預設值，請盡快設定！')\n"
      "# 可檢視日誌的管理員（逗號分隔），預設維持原本的 yang.di\n"
      "ADMIN_USERS = {u.strip() for u in os.environ.get('ADMIN_USERS', 'yang.di').split(',') if u.strip()}\n")

patch("app.py", "load_user 改用絕對路徑 DB_FILE（避免工作目錄不同讀到別的 DB）",
      "conn = sqlite3.connect('billing.db')",
      "conn = sqlite3.connect(DB_FILE)")

patch("app.py", "Excel 快取加入 TTL 變數",
      "cached_xls = None\nversion_time = None\n",
      "cached_xls = None\ncached_at = 0\nEXCEL_TTL = int(os.environ.get('EXCEL_TTL', '300'))  # 秒\nversion_time = None\n")

patch("app.py", "Excel 快取過期機制（原本永不更新，GitHub 更新後必須重啟才看得到）",
      "def load_excel_from_github(url):\n    global cached_xls, version_time\n    if cached_xls:\n        return cached_xls\n",
      "def load_excel_from_github(url):\n    global cached_xls, version_time, cached_at\n"
      "    if cached_xls and (time.time() - cached_at) < EXCEL_TTL:\n        return cached_xls\n")

patch("app.py", "記錄快取時間",
      "            cached_xls = pd.ExcelFile(excel_bytes, engine='openpyxl')\n",
      "            cached_xls = pd.ExcelFile(excel_bytes, engine='openpyxl')\n            cached_at = time.time()\n")

patch("app.py", "下載失敗時退回舊快取，並修正錯字「無發」",
      '    abort(500, description="⚠️ 無發從 GitHub 載入 Excel 檔案")',
      '    if cached_xls:\n        return cached_xls  # 下載失敗時退回舊快取\n'
      '    abort(500, description="⚠️ 無法從 GitHub 載入 Excel 檔案")')

patch("app.py", "中文字型：實際載入專案內 Noto 字型（Linux/Render 沒有微軟正黑體，圖表會變方框）",
      "matplotlib.rcParams['font.sans-serif'] = ['Microsoft JhengHei']\n"
      "matplotlib.rcParams['axes.unicode_minus'] = False\n"
      "font_path = \"./fonts/NotoSansCJKtc-Regular.otf\"\n"
      "font_prop = FontProperties(fname=font_path)\n",
      "font_path = os.path.join(BASE_DIR, \"fonts\", \"NotoSansCJKtc-Regular.otf\")\n"
      "if os.path.exists(font_path):\n"
      "    from matplotlib import font_manager\n"
      "    font_manager.fontManager.addfont(font_path)\n"
      "    font_prop = FontProperties(fname=font_path)\n"
      "    matplotlib.rcParams['font.sans-serif'] = [font_prop.get_name(), 'Microsoft JhengHei']\n"
      "else:\n"
      "    font_prop = None\n"
      "    matplotlib.rcParams['font.sans-serif'] = ['Microsoft JhengHei']\n"
      "matplotlib.rcParams['axes.unicode_minus'] = False\n")

rpatch("app.py", "關鍵字搜尋改為純文字比對（輸入 ( [ * 等字元會觸發 regex 錯誤造成 500）",
       r"contains\((keyword|store_id), case=False(?!, regex=False)",
       r"contains(\1, case=False, regex=False")

patch("app.py", "inspection_log 管理員改讀 ADMIN_USERS",
      "if current_user.username != 'yang.di':",
      "if current_user.username not in ADMIN_USERS:")

patch("app.py", "/sc_disk/update 補上 @login_required（原本未登入也能改 Google Sheet）",
      "@app.route('/sc_disk/update', methods=['POST'])\ndef update_sc_disk():",
      "@app.route('/sc_disk/update', methods=['POST'])\n@login_required\ndef update_sc_disk():")

patch("app.py", "/sc_disk/update 未知 action 時回傳 400（原本回傳 None 會 500）",
      "            return jsonify({'status': 'success', 'message': '資料已同步至 Google Sheet'})\n\n    except Exception as e:",
      "            return jsonify({'status': 'success', 'message': '資料已同步至 Google Sheet'})\n\n"
      "        return jsonify({'status': 'error', 'message': '未知的 action'}), 400\n\n    except Exception as e:")

patch("app.py", "/time 改用絕對路徑讀 MFP.xlsx",
      'xls = pd.ExcelFile("MFP/MFP.xlsx")',
      'xls = pd.ExcelFile(os.path.join(BASE_DIR, "MFP", "MFP.xlsx"))')

patch("app.py", "/worktime 改用絕對路徑讀 MFP.xlsx",
      'path = "MFP/MFP.xlsx"',
      'path = os.path.join(BASE_DIR, "MFP", "MFP.xlsx")')

patch("app.py", "calendar_events 不再吞掉所有例外",
      "    except:\n        return jsonify([])",
      "    except Exception as e:\n        print(f\"⚠️ 行事曆載入失敗: {e}\")\n        return jsonify([])")

# =====================================================================
# modules/billing.py
# =====================================================================
patch("modules/billing.py", "補必要 import",
      "import threading  # 🚀 新增：匯入背景線程模組\n",
      "import threading  # 🚀 新增：匯入背景線程模組\n"
      "import os\n"
      "from html import escape as html_escape\n"
      "from urllib.parse import urlparse\n"
      "from gspread.utils import rowcol_to_a1\n")

patch("modules/billing.py", "DB_FILE 改為專案根目錄絕對路徑（與 app.py 一致）",
      'DB_FILE = "billing.db"',
      'DB_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "billing.db")')

patch("modules/billing.py", "人員分頁白名單（防止任意讀寫 Google Sheet 其他分頁）",
      "bp = billing_bp  # 相容性別名\n",
      "bp = billing_bp  # 相容性別名\n\n"
      "# 允許讀寫的人員分頁白名單（可用環境變數 PERSON_SHEETS 以逗號分隔覆寫）\n"
      "ALLOWED_PERSON_SHEETS = {\n"
      "    s.strip() for s in os.environ.get('PERSON_SHEETS', '吳宗鴻,湯家瑋,狄澤洋,劉柏均').split(',') if s.strip()\n"
      "}\n")

patch("modules/billing.py", "登入後導向：阻擋 Open Redirect（next=https://evil.com）",
      "                    # 嚴格檢查 next_page 是否為無效的空值/空字串\n"
      "                    next_page = request.args.get('next')\n"
      "                    if not next_page or next_page.strip() == '':\n"
      "                        next_page = url_for('home')\n\n"
      "                    return redirect(next_page)\n",
      "                    # 只允許站內相對路徑，避免 Open Redirect\n"
      "                    next_page = (request.args.get('next') or '').strip()\n"
      "                    parsed_next = urlparse(next_page)\n"
      "                    if (not next_page.startswith('/') or next_page.startswith('//')\n"
      "                            or next_page.startswith('/\\\\') or parsed_next.scheme or parsed_next.netloc):\n"
      "                        next_page = url_for('home')\n\n"
      "                    return redirect(next_page)\n")

patch("modules/billing.py", "color_overdue：客戶名稱 HTML escape（避免 Google Sheet 內容造成 XSS）",
      "    if overdue:\n        return f'<span style=\"background-color:#FFC0CB\">{val}</span>'\n    else:\n        return val\n",
      "    safe_val = html_escape(str(val))  # 客戶名稱來自可編輯的 Google Sheet，必須 escape\n"
      "    if overdue:\n        return f'<span style=\"background-color:#FFC0CB\">{safe_val}</span>'\n    else:\n        return safe_val\n")

# ---- 計費路由補登入保護（除 /login 外全部都沒有 @login_required）----
for _route, _def in [
    ('@billing_bp.route("/", methods=["GET", "POST"])', "def index():"),
    ('@billing_bp.route("/invoice_log/<device_id>")', "def invoice_log(device_id):"),
    ("@billing_bp.route('/mfp_summary')", "def mfp_summary():"),
    ('@billing_bp.route("/save_person_field", methods=["POST"])', "def save_person_field():"),
    ('@billing_bp.route("/person/<sheet>")', "def person_page(sheet):"),
    ("@billing_bp.route('/get_last_counts', methods=['GET'])", "def route_get_last_counts():"),
]:
    patch("modules/billing.py", f"補 @login_required：{_def}",
          f"{_route}\n{_def}\n",
          f"{_route}\n@login_required\n{_def}\n")

patch("modules/billing.py", "person_page：分頁名稱白名單",
      'def person_page(sheet):\n    keyword = request.args.get("keyword", "").strip()\n',
      'def person_page(sheet):\n    if sheet not in ALLOWED_PERSON_SHEETS:\n        abort(404)\n'
      '    keyword = request.args.get("keyword", "").strip()\n')

patch("modules/billing.py", "save_person_field：驗證 sheet / field（原本可寫入任意分頁）",
      "    data = request.json\n    upsert_person_field(\n",
      "    data = request.get_json(silent=True) or {}\n"
      "    if (data.get(\"sheet\") not in ALLOWED_PERSON_SHEETS\n"
      "            or data.get(\"field\") not in (\"remark\", \"method\")\n"
      "            or not data.get(\"device_id\")):\n"
      "        return {\"ok\": False, \"error\": \"invalid request\"}, 400\n"
      "    upsert_person_field(\n")

rpatch("modules/billing.py", "關鍵字搜尋改為純文字比對",
       r"contains\((keyword|store_id), case=False(?!, regex=False)",
       r"contains(\1, case=False, regex=False")

# ---- 核心計費缺陷 ----
patch("modules/billing.py", "查詢模式：上月讀數重複減一個月（選 10 月會讀到 8 月）",
      "get_last_counts(keyword, prev_year, prev_month)",
      "get_last_counts(keyword, selected_year, selected_month)")

patch("modules/billing.py", "insert_usage 支援指定月份（原本固定寫入「今天」的月份）",
      'def insert_usage(device_id, color_a3, color_count, bw_count):\n    month = datetime.now().strftime("%Y%m")\n',
      'def insert_usage(device_id, color_a3, color_count, bw_count, month=None):\n'
      '    month = month or datetime.now().strftime("%Y%m")\n')

patch("modules/billing.py", "計算模式：修正上月重複減月、空群組致金額歸零、空白欄位 int('') 當機、單機欄位被群組重複累加",
      '            if contract:\n'
      '                related_devices = get_related_devices(device_id)\n'
      '                total_last_color_a3 = total_last_color = total_last_bw = 0\n'
      '                total_curr_color_a3 = total_curr_color = total_curr_bw = 0\n\n'
      '                for dev in related_devices:\n'
      '                    last_a3, last_c, last_b, _ = get_last_counts(dev, prev_year, prev_month)\n'
      '                    total_last_color_a3 += last_a3\n'
      '                    total_last_color += last_c\n'
      '                    total_last_bw += last_b\n\n'
      '                    val_a3 = request.form.get(f"curr_color_a3_{dev}")\n'
      '                    val_c = request.form.get(f"curr_color_{dev}")\n'
      '                    val_b = request.form.get(f"curr_bw_{dev}")\n\n'
      '                    if val_c is None:\n'
      '                        total_curr_color_a3 += int(request.form.get("curr_color_a3", 0))\n'
      '                        total_curr_color += int(request.form.get("curr_color", 0))\n'
      '                        total_curr_bw += int(request.form.get("curr_bw", 0))\n'
      '                    else:\n'
      '                        total_curr_color_a3 += int(val_a3 or 0)\n'
      '                        total_curr_color += int(val_c or 0)\n'
      '                        total_curr_bw += int(val_b or 0)\n',
      '            if contract:\n'
      '                # SQLite 查不到合開群組時，至少要計算本機（否則總數歸零並覆蓋帳單）\n'
      '                related_devices = get_related_devices(device_id) or [device_id]\n'
      '                total_last_color_a3 = total_last_color = total_last_bw = 0\n'
      '                total_curr_color_a3 = total_curr_color = total_curr_bw = 0\n'
      '                dev_counts = {}\n\n'
      '                for dev in related_devices:\n'
      '                    last_a3, last_c, last_b, _ = get_last_counts(dev, selected_year, selected_month)\n'
      '                    total_last_color_a3 += last_a3\n'
      '                    total_last_color += last_c\n'
      '                    total_last_bw += last_b\n\n'
      '                    val_a3 = request.form.get(f"curr_color_a3_{dev}")\n'
      '                    val_c = request.form.get(f"curr_color_{dev}")\n'
      '                    val_b = request.form.get(f"curr_bw_{dev}")\n\n'
      '                    # 單機且表單用未帶設備代號的欄位名稱時才使用備援欄位\n'
      '                    if val_c is None and len(related_devices) == 1:\n'
      '                        val_a3 = request.form.get("curr_color_a3")\n'
      '                        val_c = request.form.get("curr_color")\n'
      '                        val_b = request.form.get("curr_bw")\n\n'
      '                    cnt_a3, cnt_c, cnt_b = to_int(val_a3), to_int(val_c), to_int(val_b)\n'
      '                    total_curr_color_a3 += cnt_a3\n'
      '                    total_curr_color += cnt_c\n'
      '                    total_curr_bw += cnt_b\n'
      '                    dev_counts[dev] = (cnt_a3, cnt_c, cnt_b)\n')

patch("modules/billing.py", "使用量紀錄：沿用已解析的讀數（原本單機備援欄位時會寫入 0）並寫入所選月份",
      '                for dev in related_devices:\n'
      '                    insert_usage(\n'
      '                        dev,\n'
      '                        int(request.form.get(f"curr_color_a3_{dev}", 0)),\n'
      '                        int(request.form.get(f"curr_color_{dev}", 0)),\n'
      '                        int(request.form.get(f"curr_bw_{dev}", 0))\n'
      '                    )\n',
      '                usage_month = f"{selected_year}{selected_month:02d}"\n'
      '                for dev, (cnt_a3, cnt_c, cnt_b) in dev_counts.items():\n'
      '                    insert_usage(dev, cnt_a3, cnt_c, cnt_b, month=usage_month)\n')

patch("modules/billing.py", "save_monthly_summary 增加 year 參數（原本一律用今年，1 月補登 12 月會存成隔年）",
      "    last_bw,\n    calc_result\n):\n",
      "    last_bw,\n    calc_result,\n    year=None\n):\n")

patch("modules/billing.py", "save_monthly_summary 寫入所選年份",
      "        month_int,\n        datetime.now().year,\n",
      "        month_int,\n        year or datetime.now().year,\n")

patch("modules/billing.py", "呼叫 save_monthly_summary 時傳入 selected_year",
      "                    total_last_bw,\n                    result\n                )\n",
      "                    total_last_bw,\n                    result,\n                    year=selected_year\n                )\n")

patch("modules/billing.py", "init_db 自動補 billing_summary.last_date 欄位（新資料庫查上月讀數會 OperationalError，原本需手動跑 db_updata.py）",
      "    conn.commit()\n    conn.close()\n\ninit_db()\n",
      "    # 舊資料庫自動補欄位（取代手動執行 db_updata.py）\n"
      "    cols = [r[1] for r in c.execute(\"PRAGMA table_info(billing_summary)\").fetchall()]\n"
      "    if \"last_date\" not in cols:\n"
      "        c.execute(\"ALTER TABLE billing_summary ADD COLUMN last_date TEXT DEFAULT ''\")\n\n"
      "    conn.commit()\n    conn.close()\n\ninit_db()\n")

patch("modules/billing.py", "儲存帳單時寫入 last_date（原本沒有任何程式寫入，『上次抄表日』永遠是空白）",
      "            total_with_tax\n        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
      "            total_with_tax,\n            last_date\n        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)")

patch("modules/billing.py", "last_date 參數值",
      '        calc_result.get("含稅總額", 0)\n    ))',
      '        calc_result.get("含稅總額", 0),\n        datetime.now().strftime("%Y/%m/%d")\n    ))')

# ---- Google Sheet 逐格更新 → 批次更新 ----
patch("modules/billing.py", "新增批次更新 helper（取代逐格 update_cell，避免 60 req/min 配額被打爆）",
      "def update_contract(device_id, contract_data):\n",
      '''def _update_row_fields(ws, row_idx, data):
    """一次 batch 更新同一列的多個欄位（取代逐格 update_cell）"""
    headers = ws.row_values(1)
    payload = []
    for key, value in data.items():
        if key in headers:
            payload.append({
                "range": rowcol_to_a1(row_idx, headers.index(key) + 1),
                "values": [[value]],
            })
        else:
            print(f"⚠️ 工作表缺少欄位 {key}，已略過")
    if payload:
        ws.batch_update(payload, value_input_option="USER_ENTERED")


def update_contract(device_id, contract_data):
''')

patch("modules/billing.py", "update_contract 改用批次更新",
      "            for key, value in contract_data.items():\n"
      "                try:\n"
      "                    col = list(row.keys()).index(key) + 1\n"
      "                    ws.update_cell(idx, col, value)\n"
      "                except Exception:\n"
      "                    pass\n"
      "            return True\n",
      "            _update_row_fields(ws, idx, contract_data)\n"
      "            return True\n")

patch("modules/billing.py", "update_customer 改用批次更新",
      "            for key, value in customer_data.items():\n"
      "                try:\n"
      "                    col = list(row.keys()).index(key) + 1\n"
      "                    ws.update_cell(idx, col, value)\n"
      "                except Exception:\n"
      "                    pass\n"
      "            return True\n",
      "            _update_row_fields(ws, idx, customer_data)\n"
      "            return True\n")

# =====================================================================
# modules/sc_check.py
# =====================================================================
patch("modules/sc_check.py", "str.replace 第三參數誤傳 ''（TypeError，同步一律失敗回傳 []）",
      r"c.replace('\n', '', '').replace(' ', '')",
      r"c.replace('\n', '').replace(' ', '')")

patch("modules/sc_check.py", "匯入 login_required",
      "from flask import Blueprint, request, jsonify\n",
      "from flask import Blueprint, request, jsonify\nfrom flask_login import login_required\n")

patch("modules/sc_check.py", "/sc_check/update 補登入保護",
      "@sc_check_bp.route('/sc_check/update', methods=['POST'])\ndef sc_check_update():",
      "@sc_check_bp.route('/sc_check/update', methods=['POST'])\n@login_required\ndef sc_check_update():")

patch("modules/sc_check.py", "/sc_check/delete 補登入保護",
      "@sc_check_bp.route('/sc_check/delete', methods=['POST'])\ndef sc_check_delete():",
      "@sc_check_bp.route('/sc_check/delete', methods=['POST'])\n@login_required\ndef sc_check_delete():")

# =====================================================================
# Excel_Edge.py
# =====================================================================
patch("Excel_Edge.py", "帳密改由環境變數讀取（原本明碼寫在程式內）",
      "from datetime import datetime\n",
      "from datetime import datetime\nimport sys\n\n"
      "EIP_USER = os.environ.get(\"EIP_USER\", \"\")\n"
      "EIP_PASS = os.environ.get(\"EIP_PASS\", \"\")\n"
      "if not EIP_USER or not EIP_PASS:\n"
      "    sys.exit(\"❌ 請先設定環境變數 EIP_USER / EIP_PASS（例：setx EIP_USER 帳號、setx EIP_PASS 密碼，重開終端機生效）\")\n")

rpatch("Excel_Edge.py", "移除明碼帳號（AccountID）",
       r'(By\.NAME, "AccountID"\)\)\)\.send_keys\()"[^"]*"\)', r"\1EIP_USER)")
rpatch("Excel_Edge.py", "移除明碼密碼（PassWord）",
       r'(By\.NAME, "PassWord"\)\.send_keys\()"[^"]*"\)', r"\1EIP_PASS)")

patch("Excel_Edge.py", "新增失敗清單",
      "try:\n    # 使用 webdriver_manager 自動管理驅動版本，避免未來手動更新麻煩\n",
      "failures = []\n\ntry:\n    # 使用 webdriver_manager 自動管理驅動版本，避免未來手動更新麻煩\n")

patch("Excel_Edge.py", "POS 下載失敗記錄",
      '        print("❌ POS 報表未下載完成")',
      '        print("❌ POS 報表未下載完成")\n        failures.append("POS")')

patch("Excel_Edge.py", "MFP 下載失敗記錄",
      '        print("❌ MFP 報表未下載完成")',
      '        print("❌ MFP 報表未下載完成")\n        failures.append("MFP")')

patch("Excel_Edge.py", "失敗時回傳非 0 結束碼（Auto_Run 的 check=True 才抓得到）",
      "    driver.quit()\n\nexcept Exception as e:\n    print(f\"❌ 發生錯誤：{e}\")\n    if 'driver' in locals():\n        driver.quit()",
      "    driver.quit()\n    if failures:\n        print(f\"❌ 下載失敗項目：{', '.join(failures)}\")\n        sys.exit(1)\n\n"
      "except Exception as e:\n    print(f\"❌ 發生錯誤：{e}\")\n    if 'driver' in locals():\n        driver.quit()\n    sys.exit(1)")

# =====================================================================
# add_ver.py / Auto_Run.py / data_updw.py
# =====================================================================
patch("add_ver.py", "get_version 接受外部傳入版本號",
      "def get_version(timeout=10, auto_only=False):\n    version = None\n",
      "def get_version(timeout=10, auto_only=False, version=None):\n    if version:\n        auto_only = True  # 已有版本號就不再等待輸入\n")

patch("add_ver.py", "支援位置參數（Auto_Run 以 `add_ver.py <版本號>` 呼叫，原本 argparse 直接 exit 2）",
      '    parser.add_argument("--auto", action="store_true")\n    args = parser.parse_args()\n',
      '    parser.add_argument("--auto", action="store_true")\n'
      '    parser.add_argument("version", nargs="?", default=None)\n'
      '    args = parser.parse_args()\n')

patch("add_ver.py", "把版本號傳入 get_version",
      "        get_version(auto_only=args.auto)",
      "        get_version(auto_only=args.auto, version=args.version)")

patch("Auto_Run.py", "固定工作目錄與直譯器（避免從別處啟動時找不到腳本、用到錯誤的 python）",
      "import tkinter as tk\nfrom tkinter import messagebox\nimport subprocess\n",
      "import os\nimport sys\nimport tkinter as tk\nfrom tkinter import messagebox\nimport subprocess\n\n"
      "BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))\n"
      "os.chdir(BASE_DIR)\n"
      "PY = \"python\" if getattr(sys, \"frozen\", False) else sys.executable\n")

rpatch("Auto_Run.py", "subprocess 改用 PY", r'\["python", ', r"[PY, ")

patch("Auto_Run.py", "版本號在執行任何步驟前先驗證（原本跑完步驟 1~3 才發現沒填）",
      '    if selected[2] and not tags:\n        messagebox.showerror("錯誤", "請輸入步驟 3 的年月")\n        return\n',
      '    if selected[2] and not tags:\n        messagebox.showerror("錯誤", "請輸入步驟 3 的年月")\n        return\n'
      '    if selected[3] and not ver_entry.get().strip():\n        messagebox.showerror("錯誤", "請輸入版本號")\n        return\n')

patch("data_updw.py", "匯入 sys",
      "import win32com.client\n",
      "import sys\nimport win32com.client\n")

patch("data_updw.py", "重大錯誤時回傳非 0（原本印訊息後仍視為成功）",
      '    except Exception as e:\n        print(f"[重大錯誤] {e}")\n',
      '    except Exception as e:\n        print(f"[重大錯誤] {e}")\n        sys.exit(1)\n')

# =====================================================================
# run_update*.py / run_data_all.py：公式列號替換
# =====================================================================
for _f in ["run_update.py", "run_update2.py", "run_update3.py", "run_data_all.py"]:
    patch(_f, "匯入 openpyxl Translator",
          "import copy\n",
          "import copy\nfrom openpyxl.formula.translate import Translator\n")
    rpatch(_f, "公式複製改用 Translator（原本 str.replace(列號) 會誤改 1500、$A$500 等字串）",
           r'(\w+)\.value\.replace\(str\(ref_row\), str\(new_row\)\) if \1\.data_type == "f" else \1\.value',
           r'Translator(\1.value, origin=\1.coordinate).translate_formula(f"{\1.column_letter}{new_row}") '
           r'if \1.data_type == "f" else \1.value')

# =====================================================================
# 第二批：前端 templates 與新發現的缺陷
# =====================================================================
def fpatch(file, desc, fn):
    """fn(text) -> (new_text, 變更處數)；回傳 0 代表已套用或無需修改"""
    PATCHES.append(dict(file=file, desc=desc, kind="func", fn=fn))


# ---- app.py：/disk 硬碟統計 ----
patch("app.py", "新增 _to_int 與可用人員清單（/disk 頁面不再因文字欄位髒資料而 500）",
      '@app.route("/disk", methods=["GET"])\n',
      "def _to_int(v):\n"
      "    try:\n"
      "        return int(float(v))\n"
      "    except (TypeError, ValueError):\n"
      "        return 0\n\n"
      "DISK_USERS = {u.strip() for u in os.environ.get('DISK_USERS', '蔡欣璇,狄澤洋,湯家瑋,劉柏均,吳宗鴻').split(',') if u.strip()}\n\n"
      '@app.route("/disk", methods=["GET"])\n')

patch("app.py", "/disk 合計改用 _to_int（Google Sheet 只要有一格非數字，整頁就會 ValueError）",
      "total = {k: sum(int(r.get(k) or 0) for r in rows) for k in total_keys}",
      "total = {k: sum(_to_int(r.get(k)) for r in rows) for k in total_keys}")

patch("app.py", "disk_save 伺服器端驗證（原本前端是文字欄位、後端完全不驗證，可寫入任意字串）",
      "    if not data['user']:\n        return \"⚠️ 必須選擇使用者\", 400\n",
      "    if data['user'] not in DISK_USERS:\n        return \"⚠️ 必須選擇有效的使用者\", 400\n"
      "    for _k, _v in data.items():\n"
      "        if _k != 'user' and not (str(_v).isdigit() and len(str(_v)) <= 6):\n"
      "            return f\"⚠️ 欄位 {_k} 必須是 0 以上的整數\", 400\n")

patch("app.py", "提供模板 ADMIN_USERS（取代模板內寫死的 yang.di）",
      "@app.before_request\ndef make_session_permanent():",
      "@app.context_processor\ndef inject_admin_users():\n"
      "    return {'ADMIN_USERS': globals().get('ADMIN_USERS', {'yang.di'})}\n\n"
      "@app.before_request\ndef make_session_permanent():")

# ---- Excel_Edge.py：登入走明文 HTTP ----
patch("Excel_Edge.py", "EIP 登入改走 HTTPS（原本帳密以 http:// 明文傳輸；layout.html 的 EIP 連結本來就是 https）",
      'driver.get("http://eip.toshibatec.com.tw/Main.aspx")\n    wait.until',
      'driver.get("https://eip.toshibatec.com.tw/Main.aspx")\n    wait.until')

patch("Excel_Edge.py", "備援導向同改 HTTPS",
      'driver.get("http://eip.toshibatec.com.tw/Main.aspx") # 或者是',
      'driver.get("https://eip.toshibatec.com.tw/Main.aspx") # 或者是')

# ---- home.html：DOM XSS ----
patch("templates/home.html", "加入 esc() HTML 跳脫函式",
      "<script>\n    // ===== 頁面載入後，非同步向後端 API 取得硬碟檢測資料 =====\n",
      "<script>\n    // HTML 跳脫：後端資料來自 Excel / Google Sheet，不可直接塞進 innerHTML\n"
      "    function esc(v) {\n"
      "        return String(v == null ? '' : v).replace(/[&<>\"']/g, function (c) {\n"
      "            return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '\"': '&quot;', \"'\": '&#39;'}[c];\n"
      "        });\n"
      "    }\n\n"
      "    // ===== 頁面載入後，非同步向後端 API 取得硬碟檢測資料 =====\n")

rpatch("templates/home.html", "innerHTML 內所有欄位套用 esc()（修正 DOM XSS：SC/TM 欄位與工作內容可夾帶 <img onerror>）",
       r"\$\{row\['([^']+)'\] \|\| ''\}",
       r"${esc(row['\1'])}")

patch("templates/home.html", "刪除：伺服器回報失敗時不再把該列從畫面移除",
      "        alert(data.message);\n        tr.remove();\n",
      "        alert(data.message || data.status);\n        if (data.status === 'success') tr.remove();\n")

# ---- billing_index.html ----
patch("templates/billing_index.html", "刪除確認：改用 data 屬性（原本 customer_name 含單引號可跳出 JS 字串執行任意程式）",
      "onclick=\"return confirm('確定要刪除 {{ m.customer_name }}（設備編號：{{ m.device_id }}）嗎？');\">",
      "data-msg=\"{{ '確定要刪除 ' ~ m.customer_name ~ '（設備編號：' ~ m.device_id ~ '）嗎？' }}\"\n"
      "                                onclick=\"return confirm(this.dataset.msg);\">")

patch("templates/billing_index.html", "發票紀錄按鈕：改用 url_for + data 屬性（避免 device_id 注入 JS）",
      "onclick=\"location.href='/billing/invoice_log/{{ contract.device_id }}'\"",
      "data-url=\"{{ url_for('billing.invoice_log', device_id=contract.device_id) }}\" onclick=\"location.href=this.dataset.url\"")

patch("templates/billing_index.html", "搜尋結果連結：device_id 以 url_for 正確編碼",
      "{{ url_for('billing.index') }}?device_id={{m.device_id}}",
      "{{ url_for('billing.index', device_id=m.device_id) }}")

patch("templates/billing_index.html", "JS 內 device_id 改用 tojson",
      "const device_id = \"{{ contract.device_id if contract else '' }}\";",
      "const device_id = {{ (contract.device_id if contract else '') | tojson }};")

patch("templates/billing_index.html", "fetch 參數 encodeURIComponent",
      "fetch(`/billing/get_last_counts?device_id=${device_id}&year=${year}&month=${month}`)",
      "fetch(`/billing/get_last_counts?device_id=${encodeURIComponent(device_id)}&year=${year}&month=${month}`)")

patch("templates/billing_index.html", "客戶表單：設備編號欄位改唯讀（原本表單送出兩個 device_id，後端只取第一個，修改該欄位完全無效）",
      '<tr><td>設備編號</td><td><input type="text" name="device_id" value="{{ customer.device_id }}"></td></tr>',
      '<tr><td>設備編號</td><td><input type="text" value="{{ customer.device_id }}" readonly style="background-color:#f1f5f9;"></td></tr>')


def _move_period_selector(text):
    """把「抄表年份/月份」選單移到合開群組判斷之前，群組模式才能選月份"""
    if "<!-- period-selector (moved) -->" in text:
        return text, 0
    start_marker = '                <div class="form-row-flex">'
    end_marker = '\n\n                <table style="margin-top: 15px;">'
    anchor = "            {% if related_devices and related_devices|length > 1 %}"
    try:
        s = text.index(start_marker)
        e = text.index(end_marker, s)
        a = text.index(anchor)
    except ValueError:
        return text, 0
    if not a < s:
        return text, 0
    block = text[s:e]
    text = text[:s] + text[e:]          # 先移除原位置（位置在 anchor 之後，不影響 anchor 索引）
    text = text[:a] + "            <!-- period-selector (moved) -->\n" + block + "\n\n" + text[a:]
    return text, 1


fpatch("templates/billing_index.html",
       "合開群組表單補上年份/月份選單（原本只有單機有，群組永遠記在『今天』的月份）",
       _move_period_selector)

# ---- tjw.html ----
patch("templates/tjw.html", "設備連結改用 url_for 編碼",
      "{{ url_for('billing.index') }}?device_id={{ row[col] }}",
      "{{ url_for('billing.index', device_id=row[col]) }}")

patch("templates/tjw.html", "sheet 名稱改用 tojson 輸出到 JS",
      'sheet: "{{ page_name }}",',
      "sheet: {{ page_name | tojson }},")

patch("templates/tjw.html", "備註/抄表方式儲存失敗時提示（原本 fetch 失敗完全沒有任何回饋，使用者以為已存檔）",
      "                value: el.value\n            })\n        });\n",
      "                value: el.value\n            })\n"
      "        }).then(r => { if (!r.ok) alert('儲存失敗（' + r.status + '），請重新整理後再試'); })\n"
      "          .catch(() => alert('網路錯誤，這筆變更尚未儲存'));\n")

# ---- personal.html ----
patch("templates/personal.html", "移除 |safe（門市資料來自 Excel/EIP 匯出，含店員輸入文字，屬儲存型 XSS 入口）",
      "{{ row[key] | safe }}",
      "{{ row[key] }}")

# ---- mfp_parts.html ----
patch("templates/mfp_parts.html", "機型選項 value 寫錯：選『25AC』實際送出『20AC』",
      '<option value="20AC" {% if selected_model=="20AC" %}selected{% endif %}>25AC</option>',
      '<option value="25AC" {% if selected_model=="25AC" %}selected{% endif %}>25AC</option>')

# ---- disk.html ----
rpatch("templates/disk.html", "數量欄位改 type=number（原本 type=text 可輸入任意文字）",
       r'<input type="text" name="((?:sc|tm)_[a-z0-9]+_(?:new|old))" value="0">',
       r'<input type="number" min="0" step="1" name="\1" value="0">')

# ---- layout.html ----
patch("templates/layout.html", "管理員選單改用 ADMIN_USERS（原本寫死 yang.di）",
      "current_user.username == 'yang.di'",
      "current_user.username in ADMIN_USERS")


def _fix_sm_web(text):
    if "<!DOCTYPE" not in text:
        return text, 0
    return ('{% extends "layout.html" %}\n{% block content %}\n\n'
            '<h2>新北一報修刊版</h2>\n'
            '<iframe src="https://bot.toshibatec.com.tw/hl0800/showjob31.php"\n'
            '        width="100%" height="600" loading="lazy" referrerpolicy="no-referrer"></iframe>\n\n'
            '{% endblock %}\n'), 1


fpatch("templates/sm_web.html", "移除 block 內巢狀的 <html><head><body>（無效 HTML，瀏覽器靠容錯才能顯示）", _fix_sm_web)


# =====================================================================
# 第三批：session 同步 / 401 JSON / 共用 JS / 裝置指紋 / 分頁 / 刪契約 / POST 登出
# =====================================================================
def newfile(file, desc, content):
    PATCHES.append(dict(file=file, desc=desc, kind="newfile", content=content))


COMMON_JS = r"""/*!
 * common.js — 全站共用前端邏輯（由 fix_project.py 產生；請改這一份，不要再複製到各模板）
 *  1. 全域載入遮罩
 *  2. fetch 回 401（登入逾時）時自動導向登入頁
 *  3. 閒置自動登出 + 使用中定期 ping（前端計時與後端 session 同步）
 *  4. 裝置指紋（每次登入只送一次）
 */
(function () {
  'use strict';

  var script = document.currentScript;
  var cfg = (script && script.dataset) || {};
  var isAuth = cfg.auth === '1';
  var loginUrl = cfg.loginUrl || '/billing/login';
  var logoutUrl = cfg.logoutUrl || '/billing/logout';
  var pingUrl = cfg.pingUrl || '/api/ping';
  var logUrl = cfg.logUrl || '/api/save_log';
  var idleMs = (parseInt(cfg.idleMinutes, 10) || 10) * 60 * 1000;
  var pingEveryMs = Math.min(3 * 60 * 1000, Math.floor(idleMs / 2));
  var warnBeforeMs = Math.min(60 * 1000, Math.floor(idleMs / 5));
  var nativeFetch = window.fetch ? window.fetch.bind(window) : null;

  /* ---------- 1. 載入遮罩 ---------- */
  function ensureOverlay() {
    var el = document.getElementById('global-loading');
    if (el) return el;
    el = document.createElement('div');
    el.id = 'global-loading';
    el.style.cssText = 'display:none;position:fixed;top:0;left:0;width:100%;height:100%;' +
      'background:rgba(0,0,0,0.4);z-index:9999;justify-content:center;align-items:center;';
    var box = document.createElement('div');
    box.style.cssText = 'background:white;padding:30px 50px;border-radius:12px;font-size:22px;' +
      'font-weight:bold;box-shadow:0 4px 20px rgba(0,0,0,0.3);';
    box.textContent = '\u23F3 加載中...請稍後';
    el.appendChild(box);
    document.body.appendChild(el);
    return el;
  }
  function showGlobalLoading() { ensureOverlay().style.display = 'flex'; }
  function hideGlobalLoading() {
    var el = document.getElementById('global-loading');
    if (el) el.style.display = 'none';
  }
  window.showGlobalLoading = showGlobalLoading;
  window.hideGlobalLoading = hideGlobalLoading;

  document.addEventListener('click', function (e) {
    var link = e.target && e.target.closest ? e.target.closest('a') : null;
    if (!link) return;
    var rawHref = link.getAttribute('href');
    if (!rawHref || rawHref.indexOf('javascript') === 0 || rawHref.charAt(0) === '#') return;
    if (link.target === '_blank' || link.hasAttribute('download')) return;
    if (e.ctrlKey || e.metaKey || e.shiftKey || e.button === 1) return;
    showGlobalLoading();
  });
  document.addEventListener('submit', function () { showGlobalLoading(); });
  window.addEventListener('pageshow', hideGlobalLoading);

  /* ---------- 2. fetch 遇到 401 → 導向登入頁 ---------- */
  var redirecting = false;
  function goLogin() {
    if (redirecting) return;
    redirecting = true;
    var next = location.pathname + location.search;
    location.replace(loginUrl + (loginUrl.indexOf('?') < 0 ? '?' : '&') + 'next=' + encodeURIComponent(next));
  }
  window.goLogin = goLogin;

  if (nativeFetch) {
    window.fetch = function (input, init) {
      return nativeFetch(input, init).then(function (res) {
        if (res.status === 401) {
          var url = typeof input === 'string' ? input : ((input && input.url) || '');
          if (url.charAt(0) === '/' || url.indexOf(location.origin) === 0) goLogin();
        }
        return res;
      });
    };
  }

  /* ---------- 3. 閒置登出 + 定期 ping ---------- */
  var sessKey = 'fp_sent:' + (cfg.sess || '');

  if (isAuth) {
    var lastActivity = Date.now();
    var lastPing = Date.now();
    var banner = null;

    var showBanner = function (sec) {
      if (!banner) {
        banner = document.createElement('div');
        banner.style.cssText = 'position:fixed;left:50%;bottom:20px;transform:translateX(-50%);z-index:10000;' +
          'background:#b45309;color:#fff;padding:12px 22px;border-radius:8px;font-size:16px;' +
          'box-shadow:0 4px 16px rgba(0,0,0,0.3);';
        document.body.appendChild(banner);
      }
      banner.textContent = '\u26A0\uFE0F 長時間未操作，將於 ' + sec + ' 秒後自動登出。移動滑鼠或按任意鍵即可繼續使用。';
    };
    var hideBanner = function () {
      if (banner && banner.parentNode) banner.parentNode.removeChild(banner);
      banner = null;
    };
    var onActivity = function () { lastActivity = Date.now(); if (banner) hideBanner(); };
    ['mousemove', 'keydown', 'click', 'scroll', 'touchstart', 'input'].forEach(function (ev) {
      window.addEventListener(ev, onActivity, { passive: true, capture: true });
    });

    var doLogout = function () {
      try { sessionStorage.removeItem(sessKey); } catch (e) {}
      var done = function () { location.replace(loginUrl); };
      if (!nativeFetch) { done(); return; }
      nativeFetch(logoutUrl, { method: 'POST', credentials: 'same-origin', keepalive: true }).then(done, done);
    };
    var ping = function () {
      lastPing = Date.now();
      return fetch(pingUrl, {
        method: 'POST',
        credentials: 'same-origin',
        headers: { 'X-Requested-With': 'fetch' }
      }).catch(function () {});
    };
    var tick = function () {
      var now = Date.now();
      var idle = now - lastActivity;
      if (idle >= idleMs) { doLogout(); return; }
      if (idle >= idleMs - warnBeforeMs) { showBanner(Math.ceil((idleMs - idle) / 1000)); return; }
      // 使用中：每隔 pingEveryMs 通知後端延長 session，避免輸入到一半被登出
      if (lastActivity > lastPing && now - lastPing >= pingEveryMs) ping();
    };
    setInterval(tick, 5000);
    document.addEventListener('visibilitychange', function () { if (!document.hidden) tick(); });

    // 手動按登出時，清掉指紋旗標讓下次登入重新記錄
    document.addEventListener('submit', function (e) {
      var action = e.target && e.target.getAttribute ? (e.target.getAttribute('action') || '') : '';
      if (action && action.indexOf(logoutUrl) !== -1) {
        try { sessionStorage.removeItem(sessKey); } catch (err) {}
      }
    });
  }

  /* ---------- 4. 裝置指紋（每次登入只送一次） ---------- */
  function b64(str) {
    var bytes = new TextEncoder().encode(str);
    var bin = '';
    for (var i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
    return btoa(bin);
  }

  function getReadableDeviceFingerprint() {
    var ua = navigator.userAgent;
    var deviceName = 'Unknown Device';
    var browserName = 'Unknown Browser';

    if (/iphone/i.test(ua)) {
      deviceName = 'Apple iPhone';
    } else if (/ipad/i.test(ua)) {
      deviceName = 'Apple iPad';
    } else if (/macintosh|mac os x/i.test(ua)) {
      deviceName = 'Mac';
    } else if (/android/i.test(ua)) {
      deviceName = 'Android Device';
      var match = ua.match(/\b([A-Z0-9]+-[A-Z0-9]+)\b/);
      if (match) deviceName = 'Android (' + match[1] + ')';
    } else if (/win/i.test(ua)) {
      deviceName = 'Windows PC';
    } else if (/linux/i.test(ua)) {
      deviceName = 'Linux PC';
    }

    if (/edg/i.test(ua)) {
      browserName = 'Edge';
    } else if (/chrome/i.test(ua) && !/chromium/i.test(ua)) {
      browserName = 'Chrome';
    } else if (/safari/i.test(ua) && !/chrome/i.test(ua)) {
      browserName = 'Safari';
    } else if (/firefox/i.test(ua)) {
      browserName = 'Firefox';
    } else if (/line/i.test(ua)) {
      browserName = 'LINE App';
    }

    var res = screen.width + 'x' + screen.height;
    var shortHash = 'fingerprint_error';
    try {
      var canvas = document.createElement('canvas');
      var gl = canvas.getContext('webgl') || canvas.getContext('experimental-webgl');
      var renderer = gl ? gl.getParameter(gl.RENDERER) : 'unknown';
      var data = [
        ua,
        navigator.language || '',
        res,
        (Intl.DateTimeFormat().resolvedOptions().timeZone || ''),
        renderer
      ].join('|');
      shortHash = b64(data).substring(0, 32);
    } catch (e) {
      try { shortHash = b64(ua).substring(0, 32); } catch (e2) {}
    }
    return deviceName + ' / ' + browserName + ' (' + res + ')-' + shortHash;
  }

  function sendFingerprintOnce() {
    try { if (sessionStorage.getItem(sessKey) === '1') return; } catch (e) {}
    fetch(logUrl, {
      method: 'POST',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fingerprint: getReadableDeviceFingerprint() })
    }).then(function (res) {
      if (res.ok) { try { sessionStorage.setItem(sessKey, '1'); } catch (e) {} }
    }).catch(function () {});
  }

  if (isAuth) {
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', sendFingerprintOnce);
    else sendFingerprintOnce();
  }
})();
"""

newfile("static/js/common.js", "建立共用前端腳本 static/js/common.js（載入遮罩、401 導向、閒置登出+ping、裝置指紋）", COMMON_JS)

COMMON_INCLUDE = (
    '<script src="{{ url_for(\'static\', filename=\'js/common.js\', v=3) }}"\n'
    '        data-auth="{{ \'1\' if current_user.is_authenticated else \'0\' }}"\n'
    '        data-login-url="{{ url_for(\'billing.login\') }}"\n'
    '        data-logout-url="{{ url_for(\'billing.logout\') }}"\n'
    '        data-ping-url="{{ url_for(\'api_ping\') }}"\n'
    '        data-log-url="{{ url_for(\'save_log\') }}"\n'
    '        data-sess="{{ session.get(\'login_ts\', \'\') }}"\n'
    '        data-idle-minutes="{{ (config.PERMANENT_SESSION_LIFETIME.total_seconds() // 60) | int }}"></script>\n'
)

_RE_OVERLAY = r'(?:<!-- 全域載入動畫 -->\s*)?<div id="global-loading".*?</div>\s*</div>\s*'
_RE_LOADING_JS = r'<script>\s*function showGlobalLoading\(\) \{.*?</script>\s*'
_RE_PAGESHOW_JS = r'<script>\s*window\.addEventListener\("pageshow".*?</script>\s*'
_RE_IDLE = r'\{% if current_user\.is_authenticated %\}\s*<script>[^<]*?let idleTimeout;.*?</script>\s*\{% endif %\}\s*'


def _dedupe_page_scripts(text, patterns, include_at_idle):
    """移除各模板複製的載入遮罩/閒置登出，改引用 common.js"""
    total = 0
    for pat in patterns:
        text, n = re.subn(pat, "", text, flags=re.S)
        total += n
    text, n = re.subn(_RE_IDLE, lambda m: COMMON_INCLUDE + "\n" if include_at_idle else "", text, flags=re.S)
    return text, total + n


def _fix_layout(text):
    return _dedupe_page_scripts(text, [_RE_OVERLAY, _RE_LOADING_JS], True)


def _fix_billing_index(text):
    return _dedupe_page_scripts(text, [_RE_OVERLAY, _RE_LOADING_JS], True)


def _fix_inspection_log(text):
    return _dedupe_page_scripts(text, [_RE_OVERLAY, _RE_LOADING_JS], True)


def _fix_tjw(text):
    # tjw.html 繼承 layout，layout 已載入 common.js；這裡只需移除重複的遮罩與監聽（否則每次點擊會觸發兩次）
    return _dedupe_page_scripts(text, [_RE_OVERLAY, _RE_LOADING_JS, _RE_PAGESHOW_JS], False)


fpatch("templates/layout.html", "layout.html：重複的載入遮罩/閒置登出改引用 common.js", _fix_layout)
fpatch("templates/billing_index.html", "billing_index.html：重複的載入遮罩/閒置登出改引用 common.js", _fix_billing_index)
fpatch("templates/inspection_log.html", "inspection_log.html：重複的載入遮罩/閒置登出/指紋改引用 common.js", _fix_inspection_log)
fpatch("templates/tjw.html", "tjw.html：移除與 layout 重複的載入遮罩與監聽", _fix_tjw)

# ---- 登出改 POST（模板三處）----
_LOGOUT_A = '<a href="{{ url_for(\'billing.logout\') }}" class="logout-btn">登出</a>'
_LOGOUT_FORM = ('<form method="POST" action="{{ url_for(\'billing.logout\') }}" style="display:inline;margin:0;">'
                '<button type="submit" class="logout-btn" '
                'style="margin:0;border:none;cursor:pointer;font-family:inherit;">登出</button></form>')
for _t in ("layout.html", "billing_index.html", "inspection_log.html"):
    patch(f"templates/{_t}", "登出按鈕改為 POST 表單（GET 登出可被 <img src=/billing/logout> 強制登出）", _LOGOUT_A, _LOGOUT_FORM)

patch("modules/billing.py", "/logout 只在 POST 時登出（GET 一律導回首頁，不再登出）",
      "@billing_bp.route('/logout')\n@login_required\ndef logout():\n    logout_user()\n",
      "@billing_bp.route('/logout', methods=['GET', 'POST'])\n@login_required\ndef logout():\n"
      "    if request.method != 'POST':\n"
      "        # GET 不登出，避免 <img src=\"/billing/logout\"> 之類的跨站強制登出\n"
      "        return redirect(url_for('home'))\n"
      "    logout_user()\n")

patch("modules/billing.py", "匯入 session（記錄本次登入時間，供前端判斷是否需重送裝置指紋）",
      "from flask import Blueprint, render_template, request, redirect, url_for, abort, jsonify, current_app, flash\n",
      "from flask import Blueprint, render_template, request, redirect, url_for, abort, jsonify, current_app, flash, session\n")

patch("modules/billing.py", "登入成功時寫入 session['login_ts']",
      "                    login_user(user)\n",
      "                    login_user(user)\n"
      "                    session['login_ts'] = datetime.now(ZoneInfo(\"Asia/Taipei\")).strftime('%Y%m%d%H%M%S')\n")

# ---- app.py：401 JSON / ping / save_log / 分頁 / 向量化搜尋 ----
patch("app.py", "未登入時：API 類請求回 401 JSON，一般頁面才導向登入頁（原本 fetch 會拿到登入頁 HTML）",
      "login_manager.login_view = 'billing.login'  # 未登入時自動重定向之登入頁面路由\n",
      "login_manager.login_view = 'billing.login'  # 未登入時自動重定向之登入頁面路由\n\n"
      "_API_PREFIXES = ('/api/', '/sc_disk/', '/calendar/events', '/billing/get_last_counts', '/billing/save_person_field')\n\n"
      "@login_manager.unauthorized_handler\n"
      "def _unauthorized():\n"
      "    wants_json = (\n"
      "        request.path.startswith(_API_PREFIXES)\n"
      "        or request.is_json\n"
      "        or request.headers.get('X-Requested-With', '').lower() == 'fetch'\n"
      "        or request.accept_mimetypes.best == 'application/json'\n"
      "    )\n"
      "    if wants_json:\n"
      "        return jsonify({'status': 'error', 'message': '登入已逾時，請重新登入',\n"
      "                        'login_url': url_for('billing.login')}), 401\n"
      "    return redirect(url_for('billing.login', next=request.full_path.rstrip('?')))\n")

_SAVE_LOG_V3 = '''# save_log v3
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
'''


def _replace_save_log(text):
    if "# save_log v3" in text:
        return text, 0
    pat = re.compile(r"@app\.route\('/api/save_log', methods=\['POST'\]\)\n@login_required\ndef save_log\(\):.*?(?=\n@app\.route\('/time'\))", re.S)
    new_text, n = pat.subn(lambda m: _SAVE_LOG_V3.rstrip("\n") + "\n", text, count=1)
    return new_text, n


fpatch("app.py", "save_log 重寫：新增 /api/ping；補在登入那一列、30 分鐘內同指紋不重複（原本每次重整新增一筆）", _replace_save_log)

patch("app.py", "報表每頁筆數設定",
      "@app.route('/report')\n@login_required\ndef report():\n",
      "REPORT_PER_PAGE = int(os.environ.get('REPORT_PER_PAGE', '50'))\n\n"
      "@app.route('/report')\n@login_required\ndef report():\n")

patch("app.py", "report：建立分頁預設值與查詢參數（換頁時保留篩選條件）",
      "    if keyword or store_id or repair_item:\n        if keyword:\n",
      "    pagination = None\n"
      "    query_args = {k: v for k, v in (('keyword', keyword), ('store_id', store_id), ('repair_item', repair_item)) if v}\n"
      "    if keyword or store_id or repair_item:\n        if keyword:\n")

patch("app.py", "report：後端分頁（原本把符合條件的資料一次全部送到瀏覽器）",
      "        tables = df.to_dict(orient='records')\n",
      "        total = len(df)\n"
      "        pages = max(1, math.ceil(total / REPORT_PER_PAGE))\n"
      "        page = min(max(request.args.get('page', 1, type=int) or 1, 1), pages)\n"
      "        tables = df.iloc[(page - 1) * REPORT_PER_PAGE: page * REPORT_PER_PAGE].to_dict(orient='records')\n"
      "        pagination = {'page': page, 'pages': pages, 'total': total, 'per_page': REPORT_PER_PAGE}\n")

patch("app.py", "report：把分頁資訊傳給模板",
      "        tables=tables,\n        keyword=keyword,\n        store_id=store_id,\n",
      "        tables=tables,\n        pagination=pagination,\n        query_args=query_args,\n        keyword=keyword,\n        store_id=store_id,\n")

rpatch("app.py", "關鍵字搜尋向量化（原本逐列 df.apply，資料多時很慢）",
       r"df\[df\.apply\(lambda r: r\.astype\(str\)\.str\.contains\(keyword, case=False(?:, regex=False)?\)\.any\(\), axis=1\)\]",
       r"df[df.astype(str).apply(lambda col: col.str.contains(keyword, case=False, regex=False)).any(axis=1)]")

patch("templates/report.html", "報表分頁導覽列",
      "        </table>\n    </div>\n    {% else %}\n        <p class=\"no-data\">查無符合條件的報修資料。</p>",
      "        </table>\n    </div>\n"
      "    {% if pagination %}\n"
      "    <div class=\"pager\">\n"
      "        共 {{ pagination.total }} 筆，第 {{ pagination.page }} / {{ pagination.pages }} 頁\n"
      "        {% if pagination.page > 1 %}\n"
      "            <a href=\"{{ url_for('report', page=1, **query_args) }}\">« 第一頁</a>\n"
      "            <a href=\"{{ url_for('report', page=pagination.page - 1, **query_args) }}\">‹ 上一頁</a>\n"
      "        {% endif %}\n"
      "        {% if pagination.page < pagination.pages %}\n"
      "            <a href=\"{{ url_for('report', page=pagination.page + 1, **query_args) }}\">下一頁 ›</a>\n"
      "            <a href=\"{{ url_for('report', page=pagination.pages, **query_args) }}\">最後一頁 »</a>\n"
      "        {% endif %}\n"
      "    </div>\n"
      "    {% endif %}\n"
      "    {% else %}\n        <p class=\"no-data\">查無符合條件的報修資料。</p>")

patch("templates/report.html", "分頁導覽樣式",
      "        .no-data {\n            color: #777;\n",
      "        .pager {\n            margin: 14px 0;\n            font-size: 15px;\n        }\n\n"
      "        .pager a {\n            margin-left: 10px;\n            color: #007BFF;\n            text-decoration: none;\n        }\n\n"
      "        .no-data {\n            color: #777;\n")

# ---- 刪除客戶時一併刪除契約 ----
patch("modules/billing.py", "新增 delete_contract（刪除 contracts 分頁與 SQLite 契約，並回報仍指向它的子機）",
      "# 新客戶建檔\ndef insert_customer(",
      "def delete_contract(device_id):\n"
      "    \"\"\"刪除契約。回傳 (是否刪除, 仍以此設備為主機的子機清單)；帳單歷史(billing_summary)保留\"\"\"\n"
      "    ws = get_person_worksheet(\"contracts\")\n"
      "    rows = ws.get_all_records()\n"
      "    target = str(device_id).strip()\n"
      "    orphans = [str(r.get(\"device_id\", \"\")).strip() for r in rows\n"
      "               if str(r.get(\"master_device_id\", \"\")).strip() == target\n"
      "               and str(r.get(\"device_id\", \"\")).strip() != target]\n"
      "    deleted = False\n"
      "    for idx, row in enumerate(rows, start=2):\n"
      "        if str(row.get(\"device_id\", \"\")).strip() == target:\n"
      "            ws.delete_rows(idx)\n"
      "            deleted = True\n"
      "            break\n"
      "    try:\n"
      "        conn = sqlite3.connect(DB_FILE)\n"
      "        conn.execute(\"DELETE FROM contracts WHERE device_id = ?\", (target,))\n"
      "        conn.commit()\n"
      "        conn.close()\n"
      "    except Exception as e:\n"
      "        print(f\"⚠️ 刪除本機契約失敗: {e}\")\n"
      "    return deleted, orphans\n\n"
      "# 新客戶建檔\ndef insert_customer(")

patch("modules/billing.py", "刪除客戶時一併刪除契約，避免 contracts 留下孤兒資料",
      "            delete_customer(device_id)\n            message = f\"🗑 已刪除客戶（設備編號：{device_id}）\"\n",
      "            delete_customer(device_id)\n"
      "            _deleted, _orphans = delete_contract(device_id)\n"
      "            message = f\"🗑 已刪除客戶與契約（設備編號：{device_id}）\"\n"
      "            if _orphans:\n"
      "                message += f\"；⚠️ 以下子機仍指向此主機，請重新指定：{', '.join(_orphans)}\"\n")


# =====================================================================
# requirements
# =====================================================================
patch("requirements.txt", "補上實際用到卻未列出的套件",
      "flask-login",
      "flask-login\nrequests")

# =====================================================================
# 引擎
# =====================================================================
ENCODINGS = ["utf-8-sig", "cp950"]


def read_text(path):
    raw = path.read_bytes()
    bom = raw.startswith(b"\xef\xbb\xbf")
    for enc in ENCODINGS:
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("無法解碼（非 UTF-8 / CP950）")
    crlf = "\r\n" in text
    return text.replace("\r\n", "\n"), dict(bom=bom, crlf=crlf, enc="utf-8" if enc == "utf-8-sig" else enc)


def write_text(path, text, meta):
    if meta["crlf"]:
        text = text.replace("\n", "\r\n")
    data = text.encode(meta["enc"])
    if meta["bom"]:
        data = b"\xef\xbb\xbf" + data
    path.write_bytes(data)


def apply_patches(root, dry):
    by_file = {}
    for p in PATCHES:
        if p["kind"] == "newfile":
            continue
        by_file.setdefault(p["file"], []).append(p)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = root / "_fix_backup" / stamp
    stats = dict(applied=0, skipped=0, missing=0, nofile=0)
    changed_files = []

    # ---- 新增檔案（例如 static/js/common.js）----
    for p in [x for x in PATCHES if x["kind"] == "newfile"]:
        path = root / p["file"]
        print(f"\n📄 {p['file']}")
        content = p["content"]
        if path.exists() and path.read_text(encoding="utf-8").replace("\r\n", "\n") == content:
            print(f"   ⏭  已是最新：{p['desc']}")
            stats["skipped"] += 1
            continue
        print(f"   ✅ {p['desc']}")
        stats["applied"] += 1
        changed_files.append(p["file"])
        if not dry:
            if path.exists():
                dst = backup_dir / p["file"]
                dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(path, dst)
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(content, encoding="utf-8", newline="\n")

    for rel, plist in by_file.items():
        path = root / rel
        print(f"\n📄 {rel}")
        if not path.exists():
            print("   ⏭  檔案不存在，略過")
            stats["nofile"] += len(plist)
            continue
        try:
            text, meta = read_text(path)
        except Exception as e:
            print(f"   ❌ 讀取失敗：{e}")
            continue

        original = text
        for p in plist:
            if p["kind"] == "str":
                # 已套用判斷：new 包含 old 時只看 new；否則要 new 存在且 old 已消失
                # （避免 new 只是檔案中別處本來就有的字串而被誤判為已套用）
                if p["old"] in p["new"]:
                    already = p["new"] in text
                else:
                    already = (p["new"] in text) and (p["old"] not in text)
                if already:
                    print(f"   ⏭  已套用：{p['desc']}")
                    stats["skipped"] += 1
                elif p["old"] in text:
                    n = text.count(p["old"])
                    text = text.replace(p["old"], p["new"])
                    print(f"   ✅ {p['desc']}" + (f"（{n} 處）" if n > 1 else ""))
                    stats["applied"] += 1
                else:
                    print(f"   ⚠️  找不到比對內容（可能已被修改）：{p['desc']}")
                    stats["missing"] += 1
            elif p["kind"] == "func":
                new_text, n = p["fn"](text)
                if n:
                    text = new_text
                    print(f"   ✅ {p['desc']}")
                    stats["applied"] += 1
                else:
                    print(f"   ⏭  已套用/無需修改：{p['desc']}")
                    stats["skipped"] += 1
            else:
                new_text, n = re.subn(p["pattern"], p["repl"], text)
                if n:
                    text = new_text
                    print(f"   ✅ {p['desc']}（{n} 處）")
                    stats["applied"] += 1
                elif re.search(p["pattern"].split("(?!")[0], text):
                    print(f"   ⏭  已套用：{p['desc']}")
                    stats["skipped"] += 1
                else:
                    # 沒有可替換的內容：視為已套用或本來就沒有此問題
                    print(f"   ⏭  無需修改：{p['desc']}")
                    stats["skipped"] += 1

        if text == original:
            continue

        if rel.endswith(".py"):
            try:
                compile(text, rel, "exec")
            except SyntaxError as e:
                print(f"   ❌ 套用後語法錯誤（{e}），此檔案不會被寫入")
                continue

        changed_files.append(rel)
        if not dry:
            dst = backup_dir / rel
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, dst)
            write_text(path, text, meta)

    print("\n" + "=" * 60)
    tag = "（預覽模式，未寫入）" if dry else ""
    print(f"完成{tag}：新套用 {stats['applied']}、已套用/無需修改 {stats['skipped']}、"
          f"比對失敗 {stats['missing']}、檔案不存在 {stats['nofile']}")
    if changed_files and not dry:
        print(f"已修改 {len(changed_files)} 個檔案，備份於：{backup_dir}")
        print(f"還原指令：python {Path(sys.argv[0]).name} --restore \"{backup_dir}\"")
    print("\n⚠️ 後續必做：")
    print("  1. 在 Windows 設定環境變數 EIP_USER / EIP_PASS，並『立即更改 EIP 密碼』（舊密碼已明碼存在原始碼）")
    print("  2. Render 設定環境變數 SECRET_KEY（隨機長字串）、ADMIN_USERS（可選）")
    print("  3. 用一筆已知月份的實際資料，驗證『上月讀數』與帳單金額是否正確")
    print("  4. 瀏覽器按 Ctrl+F5 強制重新整理；到 /disk、/billing/、首頁硬碟檢測表各操作一次確認")
    print("  5. 新增了 static/js/common.js：部署到 Render 前務必 git add static/（否則頁面沒有載入遮罩與自動登出）")
    print("  6. 有網路時另外執行：python fix_project.py --add-sri  （替 CDN 套件補 integrity）")



def fetch_bytes(url):
    import urllib.request
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (fix_project SRI)"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read()


_SRI_TAG = re.compile(
    r'<(script|link)\b([^>]*?)\b(src|href)="(https://(?:cdn\.jsdelivr\.net|cdnjs\.cloudflare\.com|unpkg\.com)/[^"]+)"([^>]*)>',
    re.I)


def add_sri(root, dry, fetch=fetch_bytes):
    """替 templates/*.html 內的 CDN <script>/<link rel=stylesheet> 補上 integrity（需要連網；只下載、不執行）"""
    import base64
    import hashlib
    cache, done, failed = {}, 0, 0
    tdir = root / "templates"
    if not tdir.is_dir():
        sys.exit("找不到 templates 資料夾")
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    for path in sorted(tdir.rglob("*.html")):
        rel = path.relative_to(root)
        if any(part in ("備份", "臨時備份", "_fix_backup") for part in rel.parts):
            continue
        text, meta = read_text(path)

        def repl(m):
            nonlocal done, failed
            whole, tag, url = m.group(0), m.group(1).lower(), m.group(4)
            if "integrity=" in whole:
                return whole
            if tag == "link" and "stylesheet" not in whole.lower():
                return whole
            if url not in cache:
                try:
                    cache[url] = "sha384-" + base64.b64encode(hashlib.sha384(fetch(url)).digest()).decode()
                except Exception as e:
                    cache[url] = None
                    print(f"   ❌ 下載失敗 {url}：{e}")
            sri = cache[url]
            if not sri:
                failed += 1
                return whole
            done += 1
            print(f"   ✅ {rel}: {url.split('/')[-1]} → {sri[:24]}…")
            return whole[:-1].rstrip() + f' integrity="{sri}" crossorigin="anonymous">'

        new_text = _SRI_TAG.sub(repl, text)
        if new_text != text and not dry:
            dst = root / "_fix_backup" / stamp / rel
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, dst)
            write_text(path, new_text, meta)
    print(f"\nSRI 完成：新增 {done} 處、失敗 {failed} 處" + ("（預覽，未寫入）" if dry else ""))


def restore(root, backup):
    backup = Path(backup)
    if not backup.is_dir():
        sys.exit(f"找不到備份資料夾：{backup}")
    n = 0
    for f in backup.rglob("*"):
        if f.is_file():
            dst = root / f.relative_to(backup)
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(f, dst)
            print(f"↩️  已還原 {dst.relative_to(root)}")
            n += 1
    print(f"共還原 {n} 個檔案")


def main():
    ap = argparse.ArgumentParser(description="專案缺陷自動修正")
    ap.add_argument("project", nargs="?", default=".", help="專案根目錄（預設為目前資料夾）")
    ap.add_argument("--dry-run", action="store_true", help="只預覽，不寫檔")
    ap.add_argument("--restore", metavar="BACKUP_DIR", help="從備份資料夾還原")
    ap.add_argument("--add-sri", action="store_true", help="替 CDN 的 script/link 補 integrity（需連網，另外執行）")
    args = ap.parse_args()

    root = Path(args.project).resolve()
    if not (root / "app.py").exists():
        sys.exit(f"❌ {root} 下找不到 app.py，請在專案根目錄執行或指定路徑")

    if args.restore:
        restore(root, args.restore)
    elif args.add_sri:
        add_sri(root, args.dry_run)
    else:
        apply_patches(root, args.dry_run)


if __name__ == "__main__":
    main()