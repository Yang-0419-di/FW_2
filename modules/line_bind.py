# modules/line_bind.py
# ===== 人員 ↔ LINE 帳號綁定（Google Sheet「LINE綁定」分頁）=====
# 流程：個人頁按「取得綁定碼」→ 用自己的 LINE 傳「綁定 123456」給官方帳號（10 分鐘內有效）→ 完成。
# 一個 LINE 帳號只對應一位人員；重新綁定會取代舊的。
import secrets
from datetime import datetime, timezone, timedelta

import gspread

from modules.gsheet import client, SHEET_ID

WS_NAME = "LINE綁定"
HEADER = ["人員", "LINE userId", "綁定碼", "綁定碼期限", "綁定時間", "設定者"]
CODE_MINUTES = 10
TW = timezone(timedelta(hours=8))
FMT = "%Y-%m-%d %H:%M:%S"


def _now():
    return datetime.now(TW)


def _worksheet():
    sh = client.open_by_key(SHEET_ID)
    try:
        return sh.worksheet(WS_NAME)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=WS_NAME, rows=100, cols=len(HEADER))
        ws.update(range_name="A1:F1", values=[HEADER])
        return ws


def _rows(ws):
    """回傳 [(列號, [6 欄]), ...]"""
    out = []
    for i, r in enumerate(ws.get_all_values()[1:], start=2):
        r = [x.strip() for x in r] + [""] * (len(HEADER) - len(r))
        if r[0]:
            out.append((i, r[:len(HEADER)]))
    return out


def _write(ws, row_no, r):
    ws.update(range_name=f"A{row_no}:F{row_no}", values=[r], value_input_option="RAW")


def bindings():
    """{人員: userId}（只列已綁定的）"""
    return {r[0]: r[1] for _, r in _rows(_worksheet()) if r[1]}


def status(person):
    for _, r in _rows(_worksheet()):
        if r[0] == person:
            return {"bound": bool(r[1]), "bound_at": r[4] if r[1] else ""}
    return {"bound": False, "bound_at": ""}


def new_code(person, user):
    """產生 6 位數綁定碼（10 分鐘內有效）。"""
    ws = _worksheet()
    rows = _rows(ws)
    used = {r[2] for _, r in rows if r[2]}
    code = f"{secrets.randbelow(1000000):06d}"
    while code in used:
        code = f"{secrets.randbelow(1000000):06d}"
    expire = (_now() + timedelta(minutes=CODE_MINUTES)).strftime(FMT)
    for row_no, r in rows:
        if r[0] == person:
            r[2], r[3], r[5] = code, expire, user
            _write(ws, row_no, r)
            return code, expire
    ws.append_row([person, "", code, expire, "", user], value_input_option="RAW")
    return code, expire


def bind_by_code(code, user_id):
    """LINE 使用者傳來綁定碼 → 綁定。成功回傳人員姓名；碼錯誤或過期回傳 None。"""
    ws = _worksheet()
    rows = _rows(ws)
    now = _now()
    target = None
    for row_no, r in rows:
        if r[2] and r[2] == code:
            try:
                ok = datetime.strptime(r[3], FMT).replace(tzinfo=TW) >= now
            except ValueError:
                ok = False
            if ok:
                target = (row_no, r)
            break
    if not target:
        return None
    # 同一個 LINE 帳號原本綁在別人身上 → 先解除
    for row_no, r in rows:
        if r[1] == user_id and row_no != target[0]:
            r[1] = ""; r[4] = ""
            _write(ws, row_no, r)
    row_no, r = target
    r[1], r[2], r[3], r[4] = user_id, "", "", now.strftime(FMT)
    _write(ws, row_no, r)
    return r[0]


def unbind_person(person, user=""):
    ws = _worksheet()
    for row_no, r in _rows(ws):
        if r[0] == person:
            r[1], r[2], r[3], r[4] = "", "", "", ""
            if user:
                r[5] = user
            _write(ws, row_no, r)
            return True
    return False


def unbind_user(user_id):
    """LINE 端傳「解除綁定」。回傳被解除的人員清單。"""
    ws = _worksheet()
    out = []
    for row_no, r in _rows(ws):
        if r[1] == user_id:
            out.append(r[0])
            r[1], r[4] = "", ""
            _write(ws, row_no, r)
    return out


def persons_of(user_id):
    return [r[0] for _, r in _rows(_worksheet()) if r[1] == user_id]
