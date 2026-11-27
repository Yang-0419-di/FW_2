# modules/personal_calendar.py
# ===== 個人行事曆（Google Sheet）=====
# 每位人員自己的行事曆事件，存在既有 Google Sheet 的「個人行事曆」分頁（第一次使用時自動建立）。
# 與 data.xlsx 的部門「行事曆」分開，網站可以直接新增、編輯、刪除，不需要經過 PM 工具與 Git 發佈。
import threading
import time
import uuid
from datetime import datetime, timezone, timedelta

import gspread

from modules.gsheet import client, SHEET_ID

WS_NAME = "個人行事曆"
HEADER = ["id", "人員", "日期", "店編", "店名", "內容", "時間", "特殊", "建立者", "建立時間", "修改者", "修改時間", "提前通知", "已通知"]
# 提前通知（M 欄）：空白＝不通知；0＝當日；N＝提前 N 天。
# 已通知（N 欄）：LINE 提醒送出後由排程程式寫入「日期|天數」；修改事件時若日期或天數改變就清空，會再通知一次。
FIELDS = ["person", "date", "store", "name", "content", "time", "special"]   # 對應 B~H 欄
CACHE_TTL = 30
TW = timezone(timedelta(hours=8))

_lock = threading.Lock()
_cache = {"at": 0.0, "rows": None}


def _now():
    return datetime.now(TW).strftime("%Y-%m-%d %H:%M:%S")


def _worksheet():
    sh = client.open_by_key(SHEET_ID)
    try:
        return sh.worksheet(WS_NAME)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=WS_NAME, rows=1000, cols=len(HEADER))
        ws.update(range_name="A1:N1", values=[HEADER])
        return ws


def _invalidate():
    with _lock:
        _cache["rows"] = None


def _all_rows(fresh=False):
    with _lock:
        if not fresh and _cache["rows"] is not None and time.time() - _cache["at"] < CACHE_TTL:
            return list(_cache["rows"])
    ws = _worksheet()
    values = ws.get_all_values()
    if values and values[0][:len(HEADER)] != HEADER:
        ws.update(range_name="A1:N1", values=[HEADER])      # 舊版分頁補上新欄位標題
    rows = []
    for r in values[1:]:
        r = list(r) + [""] * (len(HEADER) - len(r))
        if not r[0].strip():
            continue
        rows.append({
            "id": r[0].strip(), "person": r[1].strip(), "date": r[2].strip()[:10],
            "store": r[3].strip(), "name": r[4].strip(), "content": r[5].strip(),
            "time": r[6].strip(), "special": r[7].strip(),
            "created_by": r[8].strip(), "created_at": r[9].strip(),
            "updated_by": r[10].strip(), "updated_at": r[11].strip(),
            "notify": r[12].strip(), "notified": r[13].strip(),
        })
    with _lock:
        _cache["rows"] = rows
        _cache["at"] = time.time()
    return list(rows)


def list_events(person, fresh=False):
    return [r for r in _all_rows(fresh) if r["person"] == person]


def get_event(event_id, fresh=True):
    for r in _all_rows(fresh):
        if r["id"] == event_id:
            return r
    return None


def _find_row(ws, event_id):
    ids = ws.col_values(1)
    for i, v in enumerate(ids):
        if i > 0 and v.strip() == event_id:
            return i + 1          # 試算表列號（從 1 起算）
    return None


def add_event(data, user):
    event_id = uuid.uuid4().hex[:12]
    now = _now()
    row = [event_id] + [data[f] for f in FIELDS] + [user, now, user, now, data.get("notify", ""), ""]
    _worksheet().append_row(row, value_input_option="RAW")
    _invalidate()
    return event_id


def update_event(event_id, data, user):
    ws = _worksheet()
    r = _find_row(ws, event_id)
    if not r:
        raise LookupError("找不到這筆事件（可能已被刪除）")
    old = ws.row_values(r) + [""] * len(HEADER)
    # 日期與提前天數都沒變 → 保留「已通知」，避免只改內容就重複提醒
    same = old[2].strip()[:10] == data["date"] and old[12].strip() == data.get("notify", "")
    notified = old[13].strip() if same else ""
    ws.update(range_name=f"B{r}:H{r}", values=[[data[f] for f in FIELDS]], value_input_option="RAW")
    ws.update(range_name=f"K{r}:N{r}", values=[[user, _now(), data.get("notify", ""), notified]], value_input_option="RAW")
    _invalidate()


def delete_event(event_id):
    ws = _worksheet()
    r = _find_row(ws, event_id)
    if not r:
        raise LookupError("找不到這筆事件（可能已被刪除）")
    ws.delete_rows(r)
    _invalidate()
