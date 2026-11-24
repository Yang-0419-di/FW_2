# modules/calendar_people.py
# ===== 排程表「人員2」儲存（Google Sheet）=====
# 為什麼不寫回 data.xlsx：
#   網站在 Render 上只「讀」GitHub 的 data.xlsx；Render 的磁碟重新部署就會清空，
#   而且 data.xlsx 由本機 pm_app.py 負責改寫，網站若也去改會互相覆蓋。
#   所以網頁填寫的人員2 另外存在既有的 Google Sheet（新分頁「行事曆人員」），
#   顯示時再依「日期 + 工務標題」對回 data.xlsx 的事件。
import re
import threading
import time
from datetime import datetime, timezone, timedelta

import gspread

from modules.gsheet import client, SHEET_ID

WS_NAME = "行事曆人員"
HEADER = ["key", "日期", "工務", "人員2", "修改者", "修改時間"]
CACHE_TTL = 30          # 秒；儲存後會立即清除，畫面儲存後以 fresh=1 重新讀取
TW = timezone(timedelta(hours=8))

_lock = threading.Lock()
_cache = {"at": 0.0, "data": None}


def make_key(date_str, title):
    """事件識別：日期(YYYY-MM-DD) + 工務標題（去頭尾空白、連續空白視為一個）。"""
    t = re.sub(r"\s+", " ", str(title or "")).strip()
    return f"{str(date_str)[:10]}|{t}"


def _worksheet():
    sh = client.open_by_key(SHEET_ID)
    try:
        return sh.worksheet(WS_NAME)
    except gspread.exceptions.WorksheetNotFound:
        # 第一次使用時自動建立分頁與標題列
        ws = sh.add_worksheet(title=WS_NAME, rows=500, cols=len(HEADER))
        ws.update(range_name="A1:F1", values=[HEADER])
        return ws


def load_people(fresh=False):
    """回傳 {key: 人員2}。fresh=True 時略過快取（儲存後立即刷新用）。"""
    with _lock:
        if not fresh and _cache["data"] is not None and time.time() - _cache["at"] < CACHE_TTL:
            return dict(_cache["data"])
    rows = _worksheet().get_all_values()
    data = {}
    for r in rows[1:]:
        if not r or not r[0].strip():
            continue
        person = r[3].strip() if len(r) > 3 else ""
        if person:
            data[r[0].strip()] = person      # 同 key 重複時以後面的為準
        else:
            data.pop(r[0].strip(), None)
    with _lock:
        _cache["data"] = data
        _cache["at"] = time.time()
    return dict(data)


def save_person(date_str, title, person2, user):
    """寫入或清除某事件的人員2。person2 為空字串代表清除。"""
    key = make_key(date_str, title)
    now = datetime.now(TW).strftime("%Y-%m-%d %H:%M:%S")
    ws = _worksheet()
    keys = ws.col_values(1)                  # 含標題列
    row_vals = [key, str(date_str)[:10], re.sub(r"\s+", " ", str(title)).strip(), person2, user, now]
    hit = [i + 1 for i, k in enumerate(keys) if i > 0 and k.strip() == key]
    if hit:
        # 更新第一筆；其餘重複列清掉人員2（避免舊值殘留）
        ws.update(range_name=f"A{hit[0]}:F{hit[0]}", values=[row_vals])
        for r in hit[1:]:
            ws.update(range_name=f"D{r}", values=[[""]])
    elif person2:
        ws.append_row(row_vals, value_input_option="RAW")
    with _lock:
        _cache["data"] = None
    return key
