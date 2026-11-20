#!/usr/bin/env python3
"""PM 保養管理（本機版）：直接讀寫 data.xlsx，不需要開 Excel，也不需要安裝任何套件。
用法：python pm_app.py        → 瀏覽器自動開啟 http://127.0.0.1:8765/
設定：pm_config.json（第一次啟動自動建立）。"""
import os, sys, re, io, copy, collections, json, glob, time, zipfile, threading, datetime, traceback, webbrowser, urllib.parse, subprocess, secrets, shutil
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

if sys.stdout is None or sys.stderr is None:      # 用 pythonw 啟動（沒有黑視窗）時，訊息改寫到 pm_log.txt
    _lg = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "pm_log.txt"), "w", encoding="utf-8", buffering=1)
    if sys.stdout is None: sys.stdout = _lg
    if sys.stderr is None: sys.stderr = _lg
try: sys.stdout.reconfigure(errors="replace")
except Exception: pass
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pm_core import *
from pm_core import _find_row, _last_data_row
from xl_engine import from_serial, serial, shift_formula
from xl_writer import escape, unescape
import pm_merge as PMG
import git_publish as GIT

BASE = os.path.dirname(os.path.abspath(__file__))
CFG_PATH = os.path.join(BASE, "pm_config.json")
DEFAULT_CFG = {"data": "data.xlsx", "backup_dir": "臨時備份\\自動備份", "port": 8765, "keep_backups": 20,
               "report_dir": "IM", "mfp_files": ["MFP\\MFP.xlsx", "MFP\\output.xlsx"], "mfp_backup_dir": "MFP\\自動備份",
               "eip_script": "Excel_Edge.py", "eip_python": "python", "eip_timeout_min": 15,
               "git_enabled": False, "git_dir": "", "git_expect_remote": "", "git_expect_branch": "", "git_warn_remote": "FW_2"}
EV = "行事曆"; ST = "門市狀態"; HP = "首頁"
EV_COLS = {"date": 1, "engineer": 4, "store": 5, "name": 6, "content": 7, "time": 8, "special": 10}
HOME_COLS = {"seq": "A", "dept_type": "D", "dept": "F", "area_code": "G", "area_mgr": "H", "phone": "I", "area_mgr_phone": "K",
             "dept_mgr": "L", "dept_mgr_phone": "M", "director": "N", "div": "X", "attr": "Z", "hub": "AC", "edc": "AD", "stb": "AE",
             "prj1": "AF", "prj2": "AG", "prj3": "AH", "iot": "AK", "pos1": "AM", "pos2": "AN", "safe1": "AO", "safe2": "AP"}


def load_cfg():
    cfg = dict(DEFAULT_CFG)
    if os.path.exists(CFG_PATH):
        try: cfg.update(json.load(open(CFG_PATH, encoding="utf-8")))
        except Exception: pass
    else: json.dump(cfg, open(CFG_PATH, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return cfg


CFG = load_cfg()
DATA = CFG["data"] if os.path.isabs(CFG["data"]) else os.path.join(BASE, CFG["data"])
BACKUP = CFG["backup_dir"] if os.path.isabs(CFG["backup_dir"]) else os.path.join(os.path.dirname(os.path.abspath(DATA)), CFG["backup_dir"])   # 相對於資料檔所在資料夾

DATA_DIR = os.path.dirname(os.path.abspath(DATA))
def _p(x): return x if os.path.isabs(x) else os.path.join(DATA_DIR, x)
REPORT_DIR = _p(CFG["report_dir"]); MFP_FILES = [_p(x) for x in CFG["mfp_files"]]; MFP_BACKUP = _p(CFG["mfp_backup_dir"])
EIP_SCRIPT = _p(CFG["eip_script"]); GIT_DIR = _p(CFG["git_dir"]) if CFG.get("git_dir") else DATA_DIR


class S:                                   # 全域狀態
    lock = threading.RLock(); status = "啟動中…"; ready = False; book = None; wb = None; base = None; base_ready = False
    pending = []; job = None; calc_date = None; err = ""


def file_calc_date(raw):
    try:
        t = zipfile.ZipFile(io.BytesIO(raw)).read("docProps/core.xml").decode("utf-8")
        d = datetime.datetime.strptime(re.search(r"<dcterms:modified[^>]*>([^<]+)<", t).group(1), "%Y-%m-%dT%H:%M:%SZ") + datetime.timedelta(hours=8)
        return d.date()
    except Exception: return datetime.date.today()


def load_all():
    S.ready = False; S.base_ready = False; S.status = "讀取 data.xlsx…"
    try:
        raw = open(DATA, "rb").read()
        S.wb = Workbook(raw); S.book = Book(raw); S.calc_date = file_calc_date(raw); S.pending = []; S.ready = True
        S.status = "已就緒（背景準備重算基準中…）"
        threading.Thread(target=prep_base, args=(S.wb, S.calc_date), daemon=True).start()
    except Exception as e:
        S.status = "讀取失敗：" + str(e); S.err = traceback.format_exc()


def prep_base(wb, cd):
    try:
        v = Engine(wb, today=cd).recalc_all()
        with S.lock:
            if S.wb is wb: S.base = v; S.base_ready = True; S.status = "已就緒"
    except Exception: S.err = traceback.format_exc()


# ---------- 從 XML 模型讀「手動輸入值」（含尚未儲存的修改）----------
def cxv(book, cx):
    if cx is None or cx.has_f: return None
    t = cx.attrs.get("t")
    if t == "inlineStr": return unescape(re.sub(r"<[^>]+>", "", cx.is_ or ""))
    if t == "s": return book.wb.ss[int(cx.v)] if cx.v not in (None, "") else None
    if t == "b": return cx.v == "1"
    if t in ("str", "e"): return unescape(cx.v) if cx.v else ""
    if cx.v in (None, ""): return None
    try: n = float(cx.v); return int(n) if n.is_integer() else n
    except ValueError: return unescape(cx.v)


def bv(book, sh, r, c):
    x = book.sheet(sh).cell(r, c); return cxv(book, x)


def cached(sh, r, c):
    cl = S.wb.sheets[sh].cells.get((r, c)); return None if cl is None else cl.v


def iso(v):
    if is_num(v): return from_serial(v).isoformat()
    return None if v in (None, "") else str(v)[:10]


def num(v):
    v = str(v).strip(); return int(v) if re.fullmatch(r"\d+", v) else v


def sid(v): return "" if v is None else (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v))


# ---------- 讀取檢視 ----------
def view_overview():
    g = lambda sh, r, c: cached(sh, r, c); out = dict(file=DATA, version=g(HP, 1, 7), pending=len(S.pending), base_ready=S.base_ready, status=S.status)
    out["home"] = dict(quarter=g(HP, 2, 1), stores=g(HP, 6, 1), closed=g(HP, 6, 2), new=g(HP, 6, 3), deco=g(HP, 6, 4), pause=g(HP, 6, 5),
                       pct=g(HP, 10, 1), done=g(HP, 10, 2), left=g(HP, 10, 3), reports=g(HP, 10, 4),
                       range=[iso(g(HP, 2, 15)) if is_num(g(HP, 2, 15)) else "", iso(g(HP, 2, 16)) if is_num(g(HP, 2, 16)) else ""],
                       projects=[sid(g(HP, 2, c)) for c in range(2, 7)])
    out["engineers"] = []
    for n in ENG_SHEETS:
        if n in S.wb.sheets: out["engineers"].append(dict(name=n, total=g(n, 6, 1), quarter=g(n, 6, 4), done=g(n, 6, 5), left=g(n, 6, 6), pct=g(n, 6, 7)))
    return out


def home_map():
    book = S.book; sx = book.sheet(HOME); m = {}
    for r, row in sx.rows.items():
        if r >= 2 and 2 in row.cells:
            v = cxv(book, row.cells[2])
            if v not in (None, ""): m[sid(v)] = r
    return m


def view_stores():
    book = S.book; mrow = {}
    msx = book.sheet(MAIN)
    for r, row in msx.rows.items():
        if r >= 22 and 1 in row.cells:
            v = cxv(book, row.cells[1])
            if v not in (None, ""): mrow[sid(v)] = r
    cache = {}
    for r in range(22, 22 + 400):
        v = cached(MAIN, r, 1)
        if v in (None, ""): break
        cache[sid(v)] = r
    hm = home_map(); out = []
    for i, r in mrow.items():
        cr = cache.get(i); pend = cr is None
        row = dict(id=i, pending=pend, name=sid(bv(book, HOME, hm[i], 3)) if i in hm else "", addr=sid(bv(book, HOME, hm[i], 10)) if i in hm else "")
        if cr:
            g = lambda c: cached(MAIN, cr, c); done = g(15)
            row.update(city=g(3), district=g(4), engineer=g(5), status=g(6) or "", pm=g(9) or "", date=iso(g(10)) if is_num(g(10)) else "", done=done, hub=g(14) or "")
        out.append(row)
    return out


def view_store(i):
    book = S.book; hm = home_map()
    if i not in hm: return {}
    r = hm[i]; sx = book.sheet(HOME); head = S.wb.sheets[HOME]; fields = []
    for c in range(1, 43):
        h = head.cells.get((1, c)); x = sx.cell(r, c)
        val = cxv(book, x) if x is not None and not x.has_f else cached(HOME, r, c)
        label = str(h.v).replace("\n", "") if h else ""
        if c in (32, 33, 34): label = f"專案{c - 31}旗標（首頁目前專案名稱：{sid(bv(book, HP, 2, c - 30)) or '--'}；填入與專案名稱完全相同的文字才算參與）"
        if colname(c) in ("AC", "AD", "AF", "AG", "AH"): continue
        if h and h.v not in (None, ""): fields.append(dict(col=colname(c), label=label, value=sid(val) if val is not None else "", formula=bool(x is not None and x.has_f)))
    flags = [dict(key=p["key"], label=p["label"], col=p["col"], code=p["code"], default=p["default"], active=bool(p["code"]) or p["key"] in ("hub", "edc"),
                  value=sid(bv(book, HOME, r, colnum(p["col"])))) for p in project_defs()]
    return dict(id=i, fields=fields, flags=flags)


def view_events():
    book = S.book; sx = book.sheet(EV); out = []
    for r, row in sx.rows.items():
        if r < 2: continue
        d = {k: cxv(book, row.cells.get(c)) for k, c in EV_COLS.items()}
        if all(v in (None, "") for v in d.values()): continue
        cr = S.wb.sheets[EV].cells.get((r, 3)); title = cr.v if cr else None
        out.append(dict(row=r, date=iso(d["date"]) or "", engineer=d["engineer"] or "", store=sid(d["store"]), name=d["name"] or "", content=d["content"] or "",
                        time=d["time"] or "", special=d["special"] or "", title=title or ""))
    return out


def status_blocks():
    book = S.book; sx = book.sheet(ST); titles = {}
    for r, row in sx.rows.items():
        if r <= 60 and 1 in row.cells:
            v = cxv(book, row.cells[1])
            if v in ("撤店", "跨季保養", "新開", "裝潢"): titles[r] = v
    order = sorted(titles); blocks = {}
    for k, t in enumerate(order):
        nxt = order[k + 1] if k + 1 < len(order) else 10 ** 6; slots = []; r = t + 2
        while r < nxt:
            x = sx.cell(r, 2)
            if x is None or not x.has_f: break
            slots.append(r); r += 1
        blocks[titles[t]] = slots
    return blocks


def view_status():
    book = S.book; sx = book.sheet(ST); out = {}
    for name, slots in status_blocks().items():
        rows = []
        for r in slots:
            i = bv(book, ST, r, 1); cr = S.wb.sheets[ST].cells
            attr = bv(book, ST, r, 4)
            if attr is None: attr = cached(ST, r, 4)
            rows.append(dict(row=r, id=sid(i), name=cached(ST, r, 2) if i not in (None, "") else "", date=iso(cached(ST, r, 3)) if is_num(cached(ST, r, 3)) else (cached(ST, r, 3) or ""),
                             attr=attr or "", cross=cached(ST, r, 5) or "", pm=cached(ST, r, 6) or "", deco_date=iso(cached(ST, r, 7)) if is_num(cached(ST, r, 7)) else "",
                             pause=bv(book, ST, r, 3) == "-"))
        out[name] = rows
    dm = []
    for r in range(2, 30):
        d = bv(book, ST, r, 7)
        if d not in (None, "") and bv(book, ST, r, 8) not in (None, ""):
            dm.append(dict(row=r, district=sid(d), engineer=sid(bv(book, ST, r, 8)), empid=sid(bv(book, ST, r, 9))))
    out["districts"] = dm; out["engineers"] = [n for n in ENG_SHEETS if n in S.book.order]
    return out



# ---------- IM / MFP 瀏覽 ----------
VIEW_DEF = {"IM": dict(default=[3, 6, 7, 2, 19, 9, 24, 16, 29, 30, 39], eng=19, store=6, dates=[24, 38, 39], datecols=[24, 9], extra={38: "提取日期", 39: "完工日期"}),
            "MFP": dict(default=[2, 5, 6, 7, 8, 9, 10, 12, 13, 15, 19, 26], eng=5, store=None, dates=[], datecols=[6], extra={})}
_VC = {}


def fmt_serial(v):
    d = from_serial(v); fr = v - int(v)
    if fr > 1e-9:
        sec = round(fr * 86400); return f"{d.isoformat()} {sec // 3600:02d}:{sec % 3600 // 60:02d}"
    return d.isoformat()


def sheet_cache(name):
    key = (id(S.wb), name)
    if key in _VC: return _VC[key]
    ws = S.wb.sheets[name]; df = VIEW_DEF[name]; maxc = ws.maxc; heads = []
    for c in range(1, maxc + 1):
        x = ws.cells.get((1, c)); h = str(x.v).replace("\n", "") if x is not None and x.v not in (None, "") else df["extra"].get(c, f"欄{colname(c)}")
        heads.append(h)
    isdate = [("時間" in h or "日期" in h) or (c + 1) in df["dates"] for c, h in enumerate(heads)]
    rows = []
    for r in range(2, ws.maxr + 1):
        vals = []; any_ = False
        for c in range(1, maxc + 1):
            x = ws.cells.get((r, c)); v = None if x is None else x.v
            if isinstance(v, XlErr): v = str(v)
            elif is_num(v) and isdate[c - 1] and 30000 < v < 80000: v = fmt_serial(v)
            elif isinstance(v, float) and v.is_integer(): v = int(v)
            if v not in (None, ""): any_ = True
            vals.append("" if v is None else v)
        if any_: rows.append((r, vals))
    blobs = ["\x01".join(str(x) for x in vals).casefold() for _, vals in rows]
    dk = lambda vals: next((str(vals[c - 1]) for c in df["datecols"] if vals[c - 1] != ""), "")
    out = dict(heads=heads, rows=rows, blobs=blobs, dates=[dk(v) for _, v in rows]); _VC.clear(); _VC[key] = out; return out


def filter_sheet(name, q):
    c = sheet_cache(name); df = VIEW_DEF[name]; idx = list(range(len(c["rows"])))
    terms = [t.casefold() for t in (q.get("q", "") or "").split() if t]
    eng = q.get("eng", ""); month = q.get("month", ""); store = q.get("store", "")
    def ok(i):
        v = c["rows"][i][1]
        if terms and not all(t in c["blobs"][i] for t in terms): return False
        if eng and str(v[df["eng"] - 1]) != eng: return False
        if store and df["store"] and str(v[df["store"] - 1]) != store: return False
        if month and not c["dates"][i].startswith(month): return False
        return True
    idx = [i for i in idx if ok(i)]
    idx.sort(key=lambda i: c["dates"][i], reverse=(q.get("order", "desc") != "asc"))
    blank = [i for i in idx if not c["dates"][i]]; idx = [i for i in idx if c["dates"][i]] + blank
    return c, idx


def view_sheet(name, q):
    if name not in VIEW_DEF: return {"error": "不支援的分頁"}
    c, idx = filter_sheet(name, q); df = VIEW_DEF[name]
    page = max(1, int(q.get("page", 1))); size = min(300, max(10, int(q.get("size", 100)))); allc = q.get("cols") == "all"
    cols = list(range(1, len(c["heads"]) + 1)) if allc else df["default"]
    engs = sorted({str(v[df["eng"] - 1]) for _, v in c["rows"] if v[df["eng"] - 1] != ""})
    months = sorted({d[:7] for d in c["dates"] if d}, reverse=True)
    part = idx[(page - 1) * size: page * size]
    return dict(name=name, total=len(idx), all_total=len(c["rows"]), page=page, size=size, cols=cols, heads=[c["heads"][k - 1] for k in cols],
                rows=[dict(r=c["rows"][i][0], v=[c["rows"][i][1][k - 1] for k in cols]) for i in part], engineers=engs, months=months)


def view_sheet_row(name, r):
    c = sheet_cache(name)
    for rr, vals in c["rows"]:
        if rr == r: return dict(fields=[dict(label=h, value=sid(v)) for h, v in zip(c["heads"], vals)])
    return {}


def sheet_csv(name, q):
    import csv
    c, idx = filter_sheet(name, q); buf = io.StringIO(); w = csv.writer(buf); w.writerow(c["heads"])
    for i in idx: w.writerow(c["rows"][i][1])
    return ("\ufeff" + buf.getvalue()).encode("utf-8")


# ---------- 個別工程師分頁 ----------
def view_engineer(name):
    if name not in ENG_SHEETS or name not in S.wb.sheets: return {"error": "找不到工程師分頁"}
    ws = S.wb.sheets[name]; book = S.book; g = lambda r, c: (lambda x: None if x is None else x.v)(ws.cells.get((r, c)))
    def tbl(r0, c1, c2, rows):
        return dict(headers=[sid(g(r0, c)) for c in range(c1, c2 + 1)], rows=[[sid(g(r, c)) for c in range(c1, c2 + 1)] for r in rows])
    t1 = tbl(1, 2, 7, [2, 3, 4]); t2 = tbl(5, 2, 7, [6]); t3 = tbl(1, 8, 12, [2, 3, 4, 5, 6])
    t2["rows"][0] = [sid(g(2, 1)), sid(g(4, 1)), sid(g(6, 1))] + t2["rows"][0]; t2["headers"] = [sid(g(1, 1)), sid(g(3, 1)), sid(g(5, 1))] + t2["headers"]
    base = [bv(book, name, 1, c) for c in (14, 15, 16)]; months = [sid(g(7, c)) for c in (3, 4, 5)]
    cols = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 17, 18, 19]; heads = [sid(g(7, c)) or colname(c) for c in cols]
    cid = {sid(g(r, 1)): r for r in range(8, ws.maxr + 1) if g(r, 1) not in (None, "")}; sx = book.sheet(name); stores = []
    for r in sorted(sx.rows):
        if r < 8: continue
        i = cxv(book, sx.rows[r].cells.get(1))
        if i in (None, ""): continue
        i = sid(i); cr = cid.get(i)
        stores.append(dict(id=i, pending=cr is None, v=[sid(g(cr, c)) if cr else ("" if c != 1 else i) for c in cols]))
    return dict(name=name, base=[sid(x) for x in base], months=months, t1=t1, t2=t2, t3=t3, heads=heads, stores=stores)


# ---------- 首頁設定 ----------
HOME_FIELDS = [("保養代號", "A2", "text"), ("專案 1", "B2", "text"), ("專案 2", "C2", "text"), ("專案 3", "D2", "text"), ("專案 4", "E2", "text"), ("專案 5", "F2", "text"),
               ("年度起（換季資料）", "I1", "date"), ("年度迄（換季資料）", "J1", "date"), ("季內月份 1", "I2", "num"), ("季內月份 2", "J2", "num"), ("季內月份 3", "K2", "num"),
               ("月份標籤 1", "I3", "text"), ("月份標籤 2", "J3", "text"), ("月份標籤 3", "K3", "text"), ("下一季首月", "L3", "num"),
               ("日期表：第 1 個月初", "I4", "date"), ("日期表：第 2 個月初", "J4", "date"), ("日期表：第 3 個月初", "K4", "date"), ("日期表：下一季季初", "L4", "date"), ("日期表：本季季初", "M4", "date"),
               ("報修累計區間：起", "O2", "date"), ("報修累計區間：迄", "P2", "date"),
               ("平均月工單：PM 月數", "I6", "num"), ("平均月工單：新裝撤移 月數", "J6", "num"), ("平均月工單：其他 月數", "K6", "num"), ("版本號（儲存時可自動更新）", "G1", "text")]


def ref_rc(ref):
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref); return int(m.group(2)), colnum(m.group(1))


def view_home():
    book = S.book; out = []
    for label, ref, kind in HOME_FIELDS:
        r, c = ref_rc(ref); v = bv(book, HP, r, c)
        if kind == "date": v = iso(v) if is_num(v) else (sid(v) if v is not None else "")
        out.append(dict(label=label, ref=ref, kind=kind, value="" if v is None else sid(v)))
    ws = S.wb.sheets[HP]; sx = book.sheet(HP); grid = []
    for r in range(40, 70):                              # 第 40 列起：MFP／POS 分區統計等；專案與 HUB 清單已獨立成分頁
        row = sx.rows.get(r)
        if not row: continue
        cells = []
        for c in range(1, 25):
            x = row.cells.get(c)
            if x is None: continue
            if x.has_f: v = ws.cells.get((r, c)); cells.append(dict(c=colname(c), f=True, v=sid(v.v if v else "")))
            else:
                v = cxv(book, x)
                if v not in (None, ""): cells.append(dict(c=colname(c), f=False, v=sid(v), num=is_num(v)))
        if cells: grid.append(dict(r=r, cells=cells))
    engs = [dict(name=n, base=[sid(bv(book, n, 1, c)) for c in (14, 15, 16)]) for n in ENG_SHEETS if n in book.order]
    projects = [sid(bv(book, HP, 2, c)) for c in range(2, 7)]
    return dict(fields=out, grid=grid, engineers=engs, projects=projects)


PROJECT_CELLS = {"B2": "AF", "C2": "AG", "D2": "AH", "E2": "AD"}


def op_home_set(d):
    """d = {cells: {ref: {v, kind}}}：寫入首頁的手動輸入格（只能改非公式格）。"""
    sx = S.book.sheet(HP); n = 0; cleared = 0; kept_total = 0
    for ref, spec in (d.get("cells") or {}).items():
        r, c = ref_rc(ref); x = sx.cell(r, c)
        if x is not None and x.has_f: raise ValueError(f"{ref} 是公式計算的格子，不能直接修改")
        v = str(spec.get("v", "")).strip(); kind = spec.get("kind", "text"); old = cxv(S.book, x)
        if kind == "date":
            if v:
                try: val = serial(datetime.date.fromisoformat(v[:10]))
                except ValueError: raise ValueError(f"{ref}：日期格式需為 YYYY-MM-DD")
            else: val = None
        elif kind == "num":
            if v:
                try: val = float(v); val = int(val) if val.is_integer() else val
                except ValueError: raise ValueError(f"{ref}：必須是數字")
            else: val = None
        else: val = v or None
        if val == old or (val is None and old in (None, "")): continue
        sx.set_const(r, c, val); n += 1
        if d.get("clear_old") and ref in PROJECT_CELLS and val and str(val).strip() not in ("--",):          # 更換專案代號 → 一併清除該欄的舊旗標
            k, kp = _clear_flags(PROJECT_CELLS[ref], str(val).strip()); cleared += k; kept_total += kp
            if k: note(f"專案代號 {ref} 改為 {val}：清除舊專案代號 {k} 家")
    if not n: return "沒有變更"
    note(f"首頁設定：修改 {n} 個欄位"); return f"已記錄 {n} 個首頁設定變更" + (f"，並清除 {cleared} 家門市的舊專案代號" if cleared else "") + (f"（另有 {kept_total} 家的內容是備註文字，已保留）" if kept_total else "")


def op_base_set(d):
    name = d["engineer"]
    if name not in ENG_SHEETS or name not in S.book.order: raise ValueError("找不到工程師分頁")
    sx = S.book.sheet(name); vals = d.get("values") or []
    for c, v in zip((14, 15, 16), vals):
        v = str(v).strip()
        try: n = float(v) if v else None
        except ValueError: raise ValueError("基礎件數必須是數字")
        sx.set_const(1, c, (int(n) if n is not None and n.is_integer() else n))
    note(f"{name} 月基礎件數 → {'、'.join(str(x) for x in vals)}"); return "已記錄"



# ---------- 專案旗標 / HUB 檢查清單 ----------
HUB_FIRST, HUB_LAST_FIXED, HUB_EXT_FIRST = 24, 39, 68          # 首頁 HUB 清單：第 24～39 列；40～67 列被其他表格佔用，不夠用時延伸到第 68 列以後


def project_defs():
    book = S.book; g = lambda c: sid(bv(book, HP, 2, c))
    code = lambda c: "" if g(c) in ("", "--") else g(c)
    return [dict(key="hub", label="HUB（常態專案，代號固定 +HUB）", code="+HUB", default="+HUB", col="AC", fixed=True),
            dict(key="p1", label="專案 1", code=code(2), default="", col="AF", slotcell="B2"),
            dict(key="p2", label="專案 2", code=code(3), default="", col="AG", slotcell="C2"),
            dict(key="p3", label="專案 3", code=code(4), default="", col="AH", slotcell="D2"),
            dict(key="edc", label="專案 4（EDC，常態專案）", code=code(5), default="+EDC", col="AD", slotcell="E2")]


def flag_value(p): return p["code"] or p["default"]


def view_projects():
    book = S.book; hs = book.sheet(HOME); defs = project_defs(); hm = home_map(); out = []
    for p in defs:
        c = colnum(p["col"]); members = []; other = 0; want = flag_value(p)
        pend = 0; old_tags = 0; texts = collections.Counter()
        for i, r in hm.items():
            v = bv(book, HOME, r, c); saved = cached(HOME, r, c)
            if (sid(v) if v is not None else "") != (sid(saved) if saved not in (None, "") else ""): pend += 1
            if v in (None, ""): continue
            if want and sid(v) == want: members.append(dict(id=i, name=sid(bv(book, HOME, r, 3)), value=sid(v), pending=(saved in (None, "") or sid(saved) != sid(v))))
            else:
                other += 1
                if sid(v).startswith("+"): old_tags += 1
                else: texts[sid(v)] += 1
        out.append(dict(p, active=bool(p["code"]), members=members, other=other, old_tags=old_tags, other_text=sum(texts.values()), text_samples=[f"{k}×{n}" for k, n in texts.most_common(6)], pending=pend))
    rowmap = {"hub": 14, "edc": 15, "p1": 16, "p2": 17, "p3": 18}        # 首頁專案表：第 14～18 列
    for p in out:
        r = rowmap[p["key"]]; g = lambda c: cached(HP, r, c)
        p["stat"] = dict(name=sid(g(1)), total=g(2) if is_num(g(2)) else None, pct=g(3) if is_num(g(3)) or isinstance(g(3), str) else None,
                         done=g(4) if is_num(g(4)) else None, left=g(5) if is_num(g(5)) else None)
    tbl = [[sid(cached(HP, r, c)) for c in range(1, 6)] for r in range(13, 19)]
    return dict(projects=out, table=tbl, hub_formula=hub_formula_status(S.book))


def view_hub():
    book = S.book; sx = book.sheet(HP); entries = []
    donemap = {}
    for r in range(HUB_FIRST, 70):
        i = cached(HP, r, 1)
        if i not in (None, "") and cached(HP, r, 5) is not None: donemap[sid(i)] = sid(cached(HP, r, 5))
    rows = list(range(HUB_FIRST, HUB_LAST_FIXED + 1)) + list(range(HUB_EXT_FIRST, HUB_EXT_FIRST + 200))
    for r in rows:
        row = sx.rows.get(r)
        if not row: continue
        i = cxv(book, row.cells.get(1))
        if i in (None, ""): continue
        entries.append(dict(row=r, id=sid(i), name=sid(cxv(book, row.cells.get(2))), spec=sid(cxv(book, row.cells.get(3))), reason=sid(cxv(book, row.cells.get(4))),
                            done=donemap.get(sid(i), "")))
    hm = home_map(); flagged = [i for i, r in hm.items() if sid(bv(book, HOME, r, colnum("AC"))) == "+HUB"]
    for e in entries: e["flagged"] = e["id"] in flagged
    listed = {e["id"] for e in entries}
    return dict(flagged_not_listed=[i for i in flagged if i not in listed], hub_formula=hub_formula_status(book),
                stats=dict(check=cached(HP, 22, 1), done=cached(HP, 22, 2), pct=cached(HP, 22, 3)), entries=entries,
                capacity=HUB_LAST_FIXED - HUB_FIRST + 1, start=sid(cached(HP, 4, 9)) if cached(HP, 4, 9) is not None else "")


def canon(i, hm):
    """店編正規化：先精確比對，找不到再用不分大小寫比對（與 Excel 一致）。"""
    i = str(i).strip()
    if i in hm: return i
    low = {k.casefold(): k for k in hm}
    return low.get(i.casefold(), i)


def _ids(text):
    return [x for x in re.split(r"[\s,，、;；]+", str(text or "")) if x]


def op_flag_set(d):
    """批次設定門市的專案旗標。d = {project, mode: add|remove|replace, ids}"""
    defs = {p["key"]: p for p in project_defs()}; p = defs.get(d.get("project"))
    if not p: raise ValueError("請選擇專案")
    want = flag_value(p)
    if not want: raise ValueError(f"{p['label']} 尚未設定專案代號，請先到「首頁設定」填入")
    mode = d.get("mode", "add"); hm = home_map(); ids = [canon(sid(num(x)), hm) for x in _ids(d.get("ids"))]
    if not ids and mode != "replace": raise ValueError("請貼上門市編號")
    sx = S.book.sheet(HOME); c = colnum(p["col"]); unknown = [i for i in ids if i not in hm]; ok = [i for i in ids if i in hm]
    if mode == "replace":
        for i, r in hm.items():
            if bv(S.book, HOME, r, c) not in (None, "") and i not in ok: sx.set_const(r, c, None)
    for i in ok: sx.set_const(hm[i], c, want if mode in ("add", "replace") else None)
    note(f"專案旗標 {p['label']}：{ {'add': '加入', 'remove': '移除', 'replace': '以名單取代'}[mode] } {len(ok)} 家"); msg = f"已{ {'add': '加入', 'remove': '移除', 'replace': '以名單取代為'}[mode] } {len(ok)} 家（旗標＝{want}）"
    if unknown: msg += f"；找不到這些店編，已略過：{'、'.join(unknown)}"
    return msg


def op_hub_upgrade(d):
    """套用新版 HUB 檢核公式（只檢核名單內＝0919 旗標 +HUB 的門市；完工只在門市主檔 W 欄判定一次）。"""
    if hub_formula_status(S.book) == "new": return "HUB 檢核公式已經是新版，不需要再套用"
    rep = upgrade_hub_formulas(S.book)
    note("HUB 檢核公式升級：" + "、".join(f"{k} {v}" for k, v in rep.items() if v)); return "已套用新版 HUB 檢核公式（共改寫 %d 格）" % sum(rep.values())


def _clear_flags(col, keep, only_tags=True):
    """清除 0919 某旗標欄的舊資料。only_tags=True：只清「以 + 開頭、且不等於 keep」的舊專案代號，備註類文字保留；
    only_tags=False：清除全部。回傳 (清除家數, 保留的其他內容家數)。"""
    hm = home_map(); sx = S.book.sheet(HOME); c = colnum(col); n = kept = 0
    for i, r in hm.items():
        v = bv(S.book, HOME, r, c)
        if v in (None, ""): continue
        t = sid(v)
        if keep and t == keep: continue
        if only_tags and not t.startswith("+"): kept += 1; continue
        sx.set_const(r, c, None); n += 1
    return n, kept


def op_flag_clear(d):
    """d = {project, scope: old|all}：清空該專案旗標欄中的舊資料（old＝與目前專案代號不同者；all＝全部）。"""
    p = {x["key"]: x for x in project_defs()}.get(d.get("project"))
    if not p: raise ValueError("請選擇專案")
    scope = d.get("scope", "old"); keep = ""
    if scope == "old":
        keep = flag_value(p)
        if not keep: raise ValueError(f"{p['label']} 尚未設定專案代號，無法判斷哪些是舊資料；請先設定代號，或改用「清空全部」")
    n, kept = _clear_flags(p["col"], keep, only_tags=(scope == "old"))
    note(f"專案旗標 {p['label']}：清除{'舊專案代號' if scope == 'old' else '全部內容'} {n} 家"); name = p["label"].split("（")[0]
    return f"已清除 {n} 家門市在『{name}』欄位的{'舊專案代號' if scope == 'old' else '全部內容'}" + (f"；另有 {kept} 家的內容不是專案代號（例如備註文字），已保留" if kept else "")


def _hub_formula_cell(sx, src_row, row):
    """確保首頁第 row 列的 E 欄（完工確認）有公式（由第 src_row 列平移複製）。"""
    src = sx.cell(src_row, 5)
    if src is None or not src.has_f: return
    x = sx.cell(row, 5)
    if x is not None and x.has_f: return
    sx.unshare(); src = sx.cell(src_row, 5); x = sx.cell(row, 5, True)
    x.v = x.is_ = None; x.va = ""; x.attrs = collections.OrderedDict((k, v) for k, v in src.attrs.items() if k == "s"); x.fattrs = {k: v for k, v in src.fattrs.items()}
    x.ftext = escape(shift_formula(unescape(src.ftext), row - src_row, 0)); x.dirty = True; sx.rows[row].dirty = True; sx.touched = True


def op_hub_save(d):
    """整份 HUB 檢查清單重寫。d = {entries: [{id, name, spec, reason}]}"""
    ents = d.get("entries") or []; book = S.book; sx = book.sheet(HP); hm = home_map(); seen = set(); clean = []
    for e in ents:
        i = canon(str(e.get("id", "")).strip(), hm)
        if not i: continue
        if i in seen: raise ValueError(f"店編 {i} 重複")
        seen.add(i); name = str(e.get("name", "")).strip()
        if not name and i in hm: name = sid(bv(book, HOME, hm[i], 3))
        clean.append((i, name, str(e.get("spec", "")).strip(), str(e.get("reason", "")).strip()))
    slots = list(range(HUB_FIRST, HUB_LAST_FIXED + 1)) + list(range(HUB_EXT_FIRST, HUB_EXT_FIRST + max(0, len(clean) - (HUB_LAST_FIXED - HUB_FIRST + 1))))
    ref_row = max((r for r in range(HUB_FIRST, HUB_LAST_FIXED + 1) if sx.cell(r, 1) is not None), default=HUB_FIRST)
    for r in range(HUB_FIRST, HUB_LAST_FIXED + 1):                 # 先清空固定區與已用過的延伸區
        for c in (1, 2, 3, 4): x = sx.cell(r, c); (x is not None) and sx.set_const(r, c, None)
    for r in [r for r in sx.rows if r >= HUB_EXT_FIRST]:
        for c in (1, 2, 3, 4): x = sx.cell(r, c); (x is not None) and sx.set_const(r, c, None)
    for k, (i, name, spec, reason) in enumerate(clean):
        r = slots[k]
        for c, v in ((1, num(i)), (2, name), (3, spec), (4, reason)):
            style = sx.cell(ref_row, c).attrs.get("s") if sx.cell(ref_row, c) is not None else None
            sx.set_const(r, c, v or None, style=style)
        _hub_formula_cell(sx, HUB_FIRST, r)
    note(f"HUB 檢查清單：共 {len(clean)} 家"); return f"已記錄 HUB 檢查清單（{len(clean)} 家）"


# ---------- 修改操作（只改記憶體中的 XML 模型；按「儲存並重算」才寫檔）----------
def note(msg): S.pending.append(msg)


def op_add_store(d):
    i = str(d.get("id", "")).strip(); name = str(d.get("name", "")).strip(); addr = str(d.get("addr", "")).strip()
    if not (i and name and addr): raise ValueError("店編、店名、住址都必須填寫")
    extra = {}
    for k, col in HOME_COLS.items():
        if k != "seq" and str(d.get(k, "")).strip(): extra[col] = str(d[k]).strip()
    if "X" not in extra:                                       # 部門：預設沿用最後一家店
        hs = S.book.sheet(HOME); last = _last_data_row(hs, 2, 2); dv = cxv(S.book, hs.cell(last, colnum("X")))
        if dv not in (None, ""): extra["X"] = dv
    eng, msg = add_store(S.book, datetime.date.today(), i, name, addr, extra=extra)
    note(f"新增門市 {i} {name}（{eng}）"); return msg


def op_remove_store(d):
    i = str(d["id"]).strip(); where = remove_store(S.book, i)
    # 也把狀態區塊與行事曆裡引用到的店編留著（不自動刪除），只提醒
    note(f"刪除門市 {i}"); return f"已刪除 {i}（{'、'.join(where)}）"


def _date_style(book):
    sx = book.sheet(EV)
    for r, row in sx.rows.items():
        x = row.cells.get(1)
        if r >= 2 and x is not None and not x.has_f and x.attrs.get("t") is None and x.v and "s" in x.attrs: return x.attrs["s"]
    return None


def op_event_save(d):
    book = S.book; sx = book.sheet(EV); row = d.get("row")
    date = str(d.get("date", "")).strip(); content = str(d.get("content", "")).strip()
    if not date or not content: raise ValueError("日期與內容必須填寫")
    try: ser = serial(datetime.date.fromisoformat(date[:10]))
    except ValueError: raise ValueError("日期格式需為 YYYY-MM-DD")
    last = max(r for r, rw in sx.rows.items() if 2 in rw.cells and rw.cells[2].has_f)
    if row in (None, ""):
        row = next((r for r in range(2, last + 1) if all(cxv(book, sx.cell(r, c)) in (None, "") for c in EV_COLS.values())), None)
        if row is None: sx.unshare(); sx.copy_row(last, last + 1); row = last + 1
    row = int(row)
    store = str(d.get("store", "")).strip(); name = str(d.get("name", "")).strip()
    if store and not name:
        hm = home_map()
        if store in hm: name = sid(bv(book, HOME, hm[store], 3))
    sx.set_const(row, EV_COLS["date"], ser, style=_date_style(book))
    for k in ("engineer", "store", "name", "content", "time", "special"):
        v = {"store": store, "name": name}.get(k, str(d.get(k, "")).strip()); sx.set_const(row, EV_COLS[k], num(v) if k == "store" and v else (v or None))
    # 閉店 / 新開 事件：提示使用者到「門市狀態」同步
    note(f"{'修改' if d.get('row') else '新增'}行事曆 {date[:10]} {content}"); return "已記錄行事曆事件"


def op_event_delete(d):
    row = int(d["row"]); sx = S.book.sheet(EV); last = max(r for r, rw in sx.rows.items() if 2 in rw.cells and rw.cells[2].has_f)
    rows = list(range(2, last + 1)); order = [r for r in rows if r != row] + [row]
    permute_rows(sx, 2, order, 1, 12)
    for c in EV_COLS.values(): sx.set_const(last, c, None)
    note(f"刪除行事曆第 {row} 列"); return "已刪除事件"


def _restore_formula(sx, slots, row, col):
    src = next((r for r in slots if r != row and sx.cell(r, col) is not None and sx.cell(r, col).has_f), None)
    if src is None: return
    sx.unshare(); s = sx.cell(src, col); x = sx.cell(row, col, True)
    x.v = x.is_ = None; x.va = ""; x.attrs.pop("t", None); x.fattrs = {k: v for k, v in s.fattrs.items()}
    x.ftext = escape(shift_formula(unescape(s.ftext), row - src, 0)); x.dirty = True; sx.rows[row].dirty = True; sx.touched = True


def op_slot_set(d):
    kind = d["block"]; i = canon(str(d.get("id", "")).strip(), home_map()); slots = status_blocks().get(kind)
    if not slots: raise ValueError("找不到區塊：" + kind)
    if not i: raise ValueError("請輸入店編")
    if i not in home_map(): raise ValueError(f"店編 {i} 不在門市主檔")
    sx = S.book.sheet(ST); book = S.book; row = d.get("row")
    if row in (None, ""):
        row = next((r for r in slots if bv(book, ST, r, 1) in (None, "")), None)
        if row is None: raise ValueError(f"『{kind}』區塊已經填滿，請先移除不需要的舊資料")
    row = int(row); sx.set_const(row, 1, num(i))
    if kind == "跨季保養": sx.set_const(row, 4, "跨季補單")
    if kind == "裝潢":
        if d.get("pause"): sx.set_const(row, 3, "-")
        elif bv(book, ST, row, 3) == "-": _restore_formula(sx, slots, row, 3)
    note(f"門市狀態『{kind}』設定 {i}"); return "已記錄"


def op_slot_clear(d):
    kind = d["block"]; row = int(d["row"]); slots = status_blocks().get(kind); sx = S.book.sheet(ST)
    if not slots or row not in slots: raise ValueError("找不到該列")
    sx.set_const(row, 1, None)
    if kind == "裝潢" and bv(S.book, ST, row, 3) == "-": _restore_formula(sx, slots, row, 3)
    note(f"門市狀態『{kind}』移除第 {row} 列"); return "已移除"


def _eng_sheet_of(book, store):
    for n in ENG_SHEETS:
        if n in book.order and _find_row(book.sheet(n), 1, 8, store): return n


def move_store(book, i, eng):
    """把店編 i 從目前的工程師分頁搬到 eng 分頁（原處移除、新處補一列）。回傳是否有搬動。"""
    src = _eng_sheet_of(book, i)
    if src == eng: return False
    if src:
        rs = book.sheet(src); r0 = _find_row(rs, 1, 8, i); last = _last_data_row(rs, 1, 8)
        permute_rows(rs, 8, [x for x in range(8, last + 1) if x != r0] + [r0], 1, 26); rs.remove_row(last)
    ds = book.sheet(eng); el = last_formula_row(ds, 6, 8); ds.unshare(); ds.copy_row(el, el + 1); ds.set_const(el + 1, 1, num(i))
    return True


def op_store_edit(d):
    """修改既有門市的手動輸入欄位。d = {id, fields: {欄位字母: 新值}}。地址改變時自動判斷是否換工程師。"""
    i = str(d["id"]).strip(); hm = home_map()
    if i not in hm: raise ValueError(f"找不到店編 {i}")
    sx = S.book.sheet(HOME); row = hm[i]; changed = []
    for col, val in (d.get("fields") or {}).items():
        if col in ("B", "A"): continue                      # 店編、序號不可改
        x = sx.cell(row, colnum(col))
        if x is not None and x.has_f: raise ValueError(f"{col} 欄是公式計算的，不能直接修改")
        old = cxv(S.book, x); new = str(val).strip()
        if sid(old) == new: continue
        sx.set_const(row, colnum(col), num(new) if (re.fullmatch(r"\d+", new) and col not in ("I", "K", "M")) else (new or None)); changed.append(col)
    if not changed: return "沒有變更"
    moved = ""
    if "J" in changed or "C" in changed:
        eng = Engine(Workbook(S.book.to_bytes()), today=datetime.date.today()).cell(HOME, row, 25)
        if not (isinstance(eng, str) and eng in S.book.order): raise ValueError(f"新地址無法判斷負責工程師（得到 {eng!r}），請確認縣市與區名")
        if move_store(S.book, i, eng): moved = f"，負責工程師變更為 {eng}"
    note(f"修改門市 {i}（{'、'.join(changed)}）{moved}"); return f"已修改 {i}{moved}"


def op_district_set(d):
    dist = str(d["district"]).strip(); eng = str(d["engineer"]).strip(); emp = str(d.get("empid", "")).strip()
    if eng not in S.book.order: raise ValueError("工程師必須是現有的分頁：" + "、".join(ENG_SHEETS))
    if not dist.endswith(("區", "鄉", "鎮", "市")): raise ValueError("區域名稱需以 區／鄉／鎮／市 結尾，例如 土城區")
    sx = S.book.sheet(ST); book = S.book
    row = next((r for r in range(2, 30) if sid(bv(book, ST, r, 7)) == dist), None)
    old = sid(bv(book, ST, row, 8)) if row else None
    if row is None:
        row = next((r for r in range(2, 29) if bv(book, ST, r, 7) in (None, "") and sx.cell(r, 7) is not None and not sx.cell(r, 7).has_f), None)
        if row is None: raise ValueError("區域對照表已滿")
    sx.set_const(row, 7, dist); sx.set_const(row, 8, eng)
    if emp: sx.set_const(row, 9, num(emp))
    moved = 0
    if old and old != eng:                                    # 既有區域改負責人 → 把該區門市搬到新工程師分頁
        cached_district = {sid(cached(MAIN, r, 1)): cached(MAIN, r, 4) for r in range(22, 500) if cached(MAIN, r, 1) not in (None, "")}
        for i, dd in cached_district.items():
            if dd == dist and move_store(book, i, eng): moved += 1
    note(f"區域對照 {dist} → {eng}" + (f"（搬移 {moved} 家店）" if moved else "")); return f"已設定 {dist} → {eng}" + (f"，並搬移 {moved} 家店到新工程師分頁" if moved else "")


# ---------- MFP_零件表（整張表可編輯；網站「MFP 零件」頁讀這張）----------
PARTS = "MFP_零件表"; PARTS_COLS = 6


def view_parts():
    sh = S.wb.sheets[PARTS]; heads = [sid(sh.cells[(1, c)].v) if (1, c) in sh.cells else "" for c in range(1, PARTS_COLS + 1)]; rows = []
    for r in range(2, sh.maxr + 1):
        v = [sid(sh.cells[(r, c)].v) if (r, c) in sh.cells else "" for c in range(1, PARTS_COLS + 1)]
        if any(x != "" for x in v): rows.append(v)
    return dict(heads=heads, rows=rows, numeric=[i + 1 for i in _parts_numeric_cols(S.book)])


def _parts_numeric_cols(book):
    sx = book.sheet(PARTS); out = []
    for c in range(1, PARTS_COLS + 1):
        vals = [cxv(book, rw.cells[c]) for r, rw in sx.rows.items() if r >= 2 and c in rw.cells]; vals = [v for v in vals if v not in (None, "")]
        if vals and all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in vals): out.append(c - 1)
    return out


def op_parts_save(d):
    """整張零件表以 rows 取代（保留標題列與格式；欄位若原本全是數字，輸入的數字會存成數字）。"""
    book = S.book; sx = book.sheet(PARTS); rows = d.get("rows")
    if not isinstance(rows, list): raise ValueError("資料格式錯誤")
    numeric = set(_parts_numeric_cols(book)); clean = []
    for row in rows:
        row = (list(row) + [""] * PARTS_COLS)[:PARTS_COLS]; row = [("" if v is None else str(v)).strip() for v in row]
        if any(row): clean.append(row)
    sx.unshare(); old_last = max([r for r, rw in sx.rows.items() if any(c in rw.cells for c in range(1, PARTS_COLS + 1))] or [1])
    for i, row in enumerate(clean):
        r = 2 + i
        if r > old_last and (r not in sx.rows or not sx.rows[r].cells): sx.copy_row(2, r)
        for c, v in enumerate(row, start=1):
            if v == "": sx.set_const(r, c, None); continue
            if (c - 1) in numeric:
                t = v.replace(",", "")
                try: f = float(t); v = int(f) if f.is_integer() else f
                except ValueError: pass
            sx.set_const(r, c, v)
    for r in [r for r in list(sx.rows) if r > 1 + len(clean)]: sx.remove_row(r)
    note(f"MFP_零件表：共 {len(clean)} 筆"); return f"已更新零件表（{len(clean)} 筆）"


OPS = {"parts_save": op_parts_save, "hub_upgrade": op_hub_upgrade, "flag_clear": op_flag_clear, "flag_set": op_flag_set, "hub_save": op_hub_save, "home_set": op_home_set, "base_set": op_base_set, "store_edit": op_store_edit, "store_add": op_add_store, "store_remove": op_remove_store, "event_save": op_event_save, "event_delete": op_event_delete,
       "slot_set": op_slot_set, "slot_clear": op_slot_clear, "district_set": op_district_set}


# ---------- 儲存並重算 ----------
def prune_backups():
    fs = sorted(glob.glob(os.path.join(BACKUP, "*_" + os.path.basename(DATA))))
    for f in fs[:-int(CFG.get("keep_backups", 20))]:
        try: os.remove(f)
        except OSError: pass


def save_core(J, step, opts):
    """儲存並重算（持有 S.lock 呼叫）：版本號 → 排序 → 重算 → 框架檢查 → 備份 → 原子寫入 data.xlsx。"""
    step("等待重算基準準備完成…")
    t0 = time.time()
    while not S.base_ready:
        if time.time() - t0 > 300: raise RuntimeError("重算基準準備逾時")
        time.sleep(0.5)
    with S.lock:
        book = S.book; today = datetime.date.today()
        if opts.get("bump_version", True):
            ver = datetime.datetime.now().strftime("%m%d%H%M"); book.sheet(HP).set_const(1, 7, ver); step(f"版本號 → {ver}")
        book.touch = True
        step("排序與重算中（約 30～60 秒）…"); data, issues, notes, vals = build(book, today, S.base)
        for n in notes: J["log"].append("  " + n)
        if issues: raise RuntimeError("框架檢查未通過，已取消寫入：\n" + "\n".join(issues))
        step("備份並寫入 data.xlsx…"); save_atomic(DATA, data, BACKUP); prune_backups()
        S.wb = Workbook(data); S.book = Book(data); S.base = vals; S.base_ready = True; S.pending = []; S.calc_date = today


def run_save(opts):
    J = S.job
    def step(msg): J["log"].append(msg); J["step"] = msg
    try:
        save_core(J, step, opts); step("完成"); J["ok"] = True
    except Exception as e:
        J["error"] = str(e); J["log"].append("❌ " + str(e)); S.err = traceback.format_exc()
    finally: J["done"] = True


# ---------- 資料更新：EIP 下載（呼叫你原本的 Excel_Edge.py，不修改它）----------
class EIP:
    proc = None; log = []; running = False; rc = None; started = 0; cancelled = False; timeout = False


def _kill_tree(p):
    try:
        if os.name == "nt": subprocess.run(["taskkill", "/PID", str(p.pid), "/T", "/F"], capture_output=True, creationflags=0x08000000)
        else: p.kill()
    except Exception: pass


def eip_status():
    return dict(running=EIP.running, rc=EIP.rc, log=EIP.log[-400:], cancelled=EIP.cancelled, timeout=EIP.timeout,
                elapsed=int(time.time() - EIP.started) if EIP.running else 0, script=EIP_SCRIPT, script_exists=os.path.exists(EIP_SCRIPT),
                python=CFG["eip_python"], env_user=bool(os.environ.get("EIP_USER")), env_pass=bool(os.environ.get("EIP_PASS")), timeout_min=CFG["eip_timeout_min"])


def eip_start(d):
    if EIP.running: raise ValueError("下載正在進行中")
    if not os.path.exists(EIP_SCRIPT): raise ValueError(f"找不到下載程式：{EIP_SCRIPT}（請把 Excel_Edge.py 放在資料夾，或修改 pm_config.json 的 eip_script）")
    env = dict(os.environ); env["PYTHONIOENCODING"] = "utf-8"; env["PYTHONUNBUFFERED"] = "1"
    if d.get("user"): env["EIP_USER"] = d["user"]                 # 只放在這次子程序的環境變數；不寫檔、不記錄
    if d.get("password"): env["EIP_PASS"] = d["password"]
    if not env.get("EIP_USER") or not env.get("EIP_PASS"): raise ValueError("尚未提供 EIP 帳號密碼：請在下方輸入（只用於這次下載、不會存檔），或先設定環境變數 EIP_USER / EIP_PASS")
    py = d.get("python") or CFG["eip_python"]
    try:
        p = subprocess.Popen([py, "-u", EIP_SCRIPT], cwd=os.path.dirname(EIP_SCRIPT), env=env, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                             creationflags=0x08000000 if os.name == "nt" else 0)
    except FileNotFoundError: raise ValueError(f"找不到 Python：{py}（pm_config.json 的 eip_python 要指向有安裝 selenium 的 python.exe）")
    EIP.proc, EIP.log, EIP.running, EIP.rc, EIP.started, EIP.cancelled, EIP.timeout = p, [f"[{time.strftime('%H:%M:%S')}] 開始執行 {os.path.basename(EIP_SCRIPT)}（上限 {CFG['eip_timeout_min']} 分鐘）"], True, None, time.time(), False, False
    def reader():
        for line in iter(p.stdout.readline, b""):
            t = line.decode("utf-8", "replace").rstrip()
            if env.get("EIP_PASS") and env["EIP_PASS"] in t: t = t.replace(env["EIP_PASS"], "******")
            EIP.log.append(t)
        p.wait(); EIP.rc = p.returncode; EIP.running = False
        EIP.log.append(f"[{time.strftime('%H:%M:%S')}] 結束（代碼 {p.returncode}）" + ("：已取消" if EIP.cancelled else "：逾時已中止" if EIP.timeout else "：成功" if p.returncode == 0 else "：失敗"))
    def watchdog():
        lim = float(CFG["eip_timeout_min"]) * 60
        while EIP.running and EIP.proc is p:
            if time.time() - EIP.started > lim: EIP.timeout = True; _kill_tree(p); return
            time.sleep(1)
    threading.Thread(target=reader, daemon=True).start(); threading.Thread(target=watchdog, daemon=True).start()
    return "已開始下載"


def eip_cancel():
    if EIP.running and EIP.proc: EIP.cancelled = True; _kill_tree(EIP.proc); return "已送出取消"
    return "目前沒有進行中的下載"


# ---------- 資料更新：預覽與合併 ----------
def update_targets():
    """要更新的三個檔案：data.xlsx（使用記憶體中的版本）＋ MFP 資料夾的兩個。"""
    return [("data.xlsx", DATA, True)] + [(os.path.basename(f), f, False) for f in MFP_FILES]


def update_preview(tags_text):
    tags = PMG.parse_tags(tags_text); data, missing = PMG.load_reports(REPORT_DIR, tags)
    out = dict(tags=tags, report_dir=REPORT_DIR, files=PMG.list_report_files(REPORT_DIR), missing=missing, targets=[], reports={k: [dict(tag=t, rows=len(r)) for t, r in v] for k, v in data.items()},
               pending=len(S.pending), mfp_backup=MFP_BACKUP)
    if not any(data.values()): return out
    for label, path, is_main in update_targets():
        t = dict(name=label, path=path, exists=os.path.exists(path), plan={}, error="")
        if t["exists"]:
            try:
                book = S.book if is_main else Book(open(path, "rb").read())
                plans, stats = PMG.plan_book(book, data)
                t["plan"] = {k: dict(new=len(plans[k]), per_tag=[dict(tag=tg, **st) for tg, st in stats[k]]) for k in plans}
            except Exception as e: t["error"] = str(e)
        out["targets"].append(t)
    return out


def run_update(opts):
    """資料更新（取代 run_data_all.py → add_ver → data_updw → save_excel2）。
    順序：先在記憶體算好 MFP\\MFP.xlsx、MFP\\output.xlsx（還沒寫檔）→ 再處理 data.xlsx 的儲存 → 最後才寫入兩個 MFP 檔，任何一步失敗就不會留下半套。"""
    J = S.job; J["title"] = "資料更新"
    def step(msg): J["log"].append(msg); J["step"] = msg
    try:
        tags = PMG.parse_tags(opts.get("tags")); step(f"月份：{'、'.join(tags)}")
        data, missing = PMG.load_reports(REPORT_DIR, tags)
        for m in missing: J["log"].append("⚠ 找不到報表：" + m)
        if not any(data.values()): raise RuntimeError("沒有任何可用的報表檔（請先下載，或確認月份與 IM 資料夾）")
        today = datetime.date.today(); aux = []
        for path in MFP_FILES:
            if not os.path.exists(path): J["log"].append(f"⚠ 找不到 {path}，略過"); continue
            step(f"整理 {os.path.basename(path)}…"); raw = open(path, "rb").read()
            r = PMG.build_merged(path, data, today, file_calc_date(raw), build, lambda m: J["log"].append(m))
            if r["issues"]: raise RuntimeError(f"{os.path.basename(path)} 框架檢查未通過，已取消（沒有寫入任何檔案）：\n" + "\n".join(r["issues"]))
            add = {k: (v["added"] if isinstance(v, dict) else v) for k, v in r["added"].items()}
            J["log"].append(f"  {os.path.basename(path)}：IM +{add.get('IM', 0)}、MFP +{add.get('MFP', 0)}")
            if r["data"] is not None: aux.append((path, r["data"]))
        t0 = time.time()
        while not S.base_ready:                       # 先等基準好（不能在持鎖時等，背景執行緒要拿鎖才能完成）
            if time.time() - t0 > 300: raise RuntimeError("重算基準準備逾時")
            time.sleep(0.5)
        with S.lock:
            if S.pending: raise RuntimeError("還有未儲存的修改，請先到「儲存」分頁處理後再更新資料")
            step("併入 data.xlsx…"); plans, stats = PMG.plan_book(S.book, data); snap = {n: copy.deepcopy(sx) for n, sx in S.book._sheets.items()}
            try:
                got = {k: PMG.apply_to_book(S.book, k, rows) for k, rows in plans.items()}
                for k, g in got.items():
                    J["log"].append(f"  data.xlsx {k}：新增 {g['added']} 筆" + (f"（第 {g['first_row']}～{g['last_row']} 列）" if g["added"] else "") + (f"，日期無法解析 {g['bad_dates']} 筆" if g["bad_dates"] else ""))
                    for tg, st in stats[k]: J["log"].append(f"    {tg}：讀取 {st['read']}、非數字案件號 {st['non_numeric']}、被過濾 {st['filtered']}、已存在 {st['duplicate']}、新增 {st['new']}")
                if sum(g["added"] for g in got.values()) == 0 and not aux:
                    S.book._sheets = snap; J["log"].append("沒有任何新資料，未修改檔案"); step("完成（沒有新資料）"); J["ok"] = True; return
                save_core(J, step, dict(bump_version=opts.get("bump_version", True)))
            except Exception:
                S.book._sheets = snap; raise
        for path, bts in aux:
            step(f"寫入 {os.path.basename(path)}…"); save_atomic(path, bts, MFP_BACKUP)
            fs = sorted(glob.glob(os.path.join(MFP_BACKUP, "*_" + os.path.basename(path))))
            for f in fs[:-int(CFG.get("keep_backups", 20))]:
                try: os.remove(f)
                except OSError: pass
        step("完成"); J["ok"] = True
    except Exception as e:
        J["error"] = str(e); J["log"].append("❌ " + str(e)); S.err = traceback.format_exc()
    finally: J["done"] = True


# ---------- Git 發佈（需要按鈕確認；預設關閉）----------
class G:
    token = None; facts = None; log = []; running = False; done = False; ok = False; error = ""; version = ""


def current_version():
    try: return str(S.wb.sheets[HP].cells[(1, 7)].v)
    except Exception: return ""


def git_state():
    return dict(enabled=bool(CFG.get("git_enabled")), repo_dir=GIT_DIR, version=current_version(), plan_text=GIT.describe(GIT.build_plan(current_version() or "<版本號>")),
                pending=len(S.pending), preflight=G.facts, running=G.running, done=G.done, ok=G.ok, error=G.error, log=G.log[-300:], forbidden=list(GIT.FORBIDDEN_PARTS))


def git_preflight(d):
    if not CFG.get("git_enabled"): raise ValueError("Git 功能目前是關閉的（pm_config.json 的 git_enabled 為 false）。確認整個專案沒問題後，再把它改成 true。")
    if G.running: raise ValueError("發佈進行中")
    if S.pending: raise ValueError("還有未儲存的修改。請先「儲存並重算」，再發佈。")
    ver = current_version()
    if not ver: raise ValueError("讀不到首頁 G1 版本號")
    f = GIT.preflight(GIT_DIR, ver, CFG.get("git_expect_remote", ""), CFG.get("git_expect_branch", ""), CFG.get("git_warn_remote", ""))
    f["version"] = ver; f["need_ack"] = bool(f["warnings"]) and any("正式環境" in w for w in f["warnings"])
    G.token = secrets.token_hex(8) if f["ok"] else None; f["token"] = G.token; G.facts = f; G.done = False; G.ok = False; G.error = ""; G.log = []
    return f


def run_git(d):
    try:
        GIT.execute(GIT_DIR, G.facts["version"], lambda m: G.log.append(m))
        G.ok = True; G.log.append("✅ 已推送。Render 會依 GitHub 更新自動重新部署。")
    except Exception as e: G.error = str(e); G.log.append("❌ " + str(e))
    finally: G.running = False; G.done = True; G.token = None


def git_publish(d):
    if not CFG.get("git_enabled"): raise ValueError("Git 功能目前是關閉的")
    if G.running: raise ValueError("發佈進行中")
    if not G.token or d.get("token") != G.token: raise ValueError("請先按「檢查發佈內容」，並在確認內容後再發佈（檢查結果已失效或不存在）")
    if d.get("confirm") is not True: raise ValueError("需要明確確認")
    if G.facts.get("need_ack") and d.get("ack") is not True: raise ValueError("請勾選「我已確認目標倉庫無誤」")
    if S.pending: raise ValueError("還有未儲存的修改")
    if current_version() != G.facts["version"]: raise ValueError("版本號在檢查後又變動了，請重新檢查發佈內容")
    G.running = True; G.done = False; G.ok = False; G.error = ""; G.log = [f"[{time.strftime('%H:%M:%S')}] 開始發佈（版本 {G.facts['version']}）"]
    threading.Thread(target=run_git, args=(d,), daemon=True).start(); return "已開始發佈"


# ---------- HTTP ----------
class H(BaseHTTPRequestHandler):
    def log_message(self, *a): pass

    def send(self, code, body, ctype="application/json; charset=utf-8", headers=None):
        b = body if isinstance(body, bytes) else json.dumps(body, ensure_ascii=False, default=str).encode()
        self.send_response(code); self.send_header("Content-Type", ctype); self.send_header("Content-Length", str(len(b))); self.send_header("Cache-Control", "no-store")
        for k, v in (headers or {}).items(): self.send_header(k, v)
        self.end_headers(); self.wfile.write(b)

    def need_ready(self):
        if not S.ready: self.send(503, {"error": S.status}); return False
        return True

    def do_GET(self):
        try:
            u = urllib.parse.urlparse(self.path); p = u.path; q = {k: v[0] for k, v in urllib.parse.parse_qs(u.query).items()}
            if p == "/": return self.send(200, open(os.path.join(BASE, "index.html"), "rb").read(), "text/html; charset=utf-8")
            if p == "/api/ping": return self.send(200, {"pm": True})
            if p == "/api/state": return self.send(200, dict(ready=S.ready, status=S.status, base_ready=S.base_ready, pending=list(S.pending), job=S.job, file=DATA))
            if not self.need_ready(): return
            with S.lock:
                if p == "/api/overview": return self.send(200, view_overview())
                if p == "/api/stores": return self.send(200, view_stores())
                if p == "/api/store": return self.send(200, view_store(urllib.parse.unquote(q.get("id", ""))))
                if p == "/api/events": return self.send(200, view_events())
                if p == "/api/status": return self.send(200, view_status())
                if p == "/api/sheet": return self.send(200, view_sheet(q.get("name", ""), q))
                if p == "/api/sheetrow": return self.send(200, view_sheet_row(q.get("name", ""), int(q.get("r", 0))))
                if p == "/api/sheet.csv": return self.send(200, sheet_csv(q.get("name", ""), q), "text/csv; charset=utf-8", {"Content-Disposition": "attachment; filename=" + urllib.parse.quote(q.get("name", "data") + ".csv")})
                if p == "/api/engineer": return self.send(200, view_engineer(q.get("name", "")))
                if p == "/api/home": return self.send(200, view_home())
                if p == "/api/projects": return self.send(200, view_projects())
                if p == "/api/hub": return self.send(200, view_hub())
                if p == "/api/parts": return self.send(200, view_parts())
                if p == "/api/update/preview": return self.send(200, update_preview(q.get("tags", "")))
            if p == "/api/eip/status": return self.send(200, eip_status())
            if p == "/api/git/state": return self.send(200, git_state())
            self.send(404, {"error": "not found"})
        except Exception as e: self.send(500, {"error": str(e)}); S.err = traceback.format_exc()

    def do_POST(self):
        try:
            d = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0))) or b"{}"); p = self.path
            if p == "/api/quit":
                if S.job and not S.job.get("done"): return self.send(409, {"error": "儲存／更新進行中，請等它完成再結束程式"})
                if G.running: return self.send(409, {"error": "Git 發佈進行中，請等它完成"})
                if EIP.running: eip_cancel()
                self.send(200, {"ok": True}); threading.Timer(0.5, lambda: os._exit(0)).start(); return
            if p == "/api/reload":
                if S.job and not S.job.get("done"): return self.send(409, {"error": "儲存進行中"})
                load_all(); return self.send(200, {"ok": True})
            if p == "/api/eip/start": return self.send(200, {"ok": True, "msg": eip_start(d)})
            if p == "/api/eip/cancel": return self.send(200, {"ok": True, "msg": eip_cancel()})
            if p == "/api/git/preflight": return self.send(200, git_preflight(d))
            if p == "/api/git/publish": return self.send(200, {"ok": True, "msg": git_publish(d)})
            if not self.need_ready(): return
            if S.job and not S.job.get("done"): return self.send(409, {"error": "儲存進行中，請稍候"})
            if p == "/api/update/run":
                S.job = dict(log=[], step="", done=False, ok=False, error="", title="資料更新"); threading.Thread(target=run_update, args=(d,), daemon=True).start(); return self.send(200, {"ok": True})
            if p == "/api/save":
                S.job = dict(log=[], step="", done=False, ok=False, error=""); threading.Thread(target=run_save, args=(d,), daemon=True).start(); return self.send(200, {"ok": True})
            if p == "/api/discard": load_all(); return self.send(200, {"ok": True})
            if p.startswith("/api/op/"):
                op = OPS.get(p.rsplit("/", 1)[1])
                if not op: return self.send(404, {"error": "unknown op"})
                with S.lock:                                    # 交易：失敗就整個還原，不留半套資料
                    snap = {n: copy.deepcopy(sx) for n, sx in S.book._sheets.items()}; npend = len(S.pending)
                    try: msg = op(d)
                    except Exception as e:
                        S.book._sheets = snap; del S.pending[npend:]
                        if isinstance(e, (ValueError, KeyError)): return self.send(400, {"error": str(e.args[0]) if e.args else str(e)})
                        raise
                return self.send(200, {"ok": True, "msg": msg, "pending": len(S.pending)})
            self.send(404, {"error": "not found"})
        except Exception as e: self.send(400 if isinstance(e, ValueError) else 500, {"error": str(e)}); S.err = traceback.format_exc()


def open_url(url):
    """自動開啟預設瀏覽器。Windows 優先用 os.startfile（最可靠），失敗再退回 webbrowser；都失敗就在視窗提示。
    可在 pm_config.json 加 "browser": "C:\\\\Program Files (x86)\\\\Microsoft\\\\Edge\\\\Application\\\\msedge.exe" 指定瀏覽器。"""
    try:
        b = CFG.get("browser")
        if b and os.path.exists(b): subprocess.Popen([b, url]); return
        if os.name == "nt": os.startfile(url); return
        if webbrowser.open(url): return
    except Exception: pass
    try:
        if webbrowser.open(url): return
    except Exception: pass
    print("（無法自動開啟瀏覽器，請手動複製網址貼到瀏覽器：", url, "）")


def main():
    if not os.path.exists(DATA): sys.exit(f"找不到資料檔：{DATA}\n請把 data.xlsx 放在這個資料夾，或修改 pm_config.json 的 data 路徑。")
    port = int(CFG.get("port", 8765)); srv = None
    try:                                           # 已經有一個在跑了：直接把瀏覽器帶過去，不要再開第二個
        import urllib.request
        if json.loads(urllib.request.urlopen(f"http://127.0.0.1:{port}/api/ping", timeout=1.5).read()).get("pm"):
            print("PM 保養管理已經在執行，直接開啟網頁。"); 
            if "--no-browser" not in sys.argv: open_url(f"http://127.0.0.1:{port}/")
            return
    except Exception: pass
    for p in range(port, port + 20):
        try: srv = ThreadingHTTPServer(("127.0.0.1", p), H); break
        except OSError: continue
    url = f"http://127.0.0.1:{srv.server_port}/"
    print("PM 保養管理已啟動：", url, "\n資料檔：", DATA, "\n備份資料夾：", BACKUP, "\n（關閉此視窗即結束程式）")
    # 先讓伺服器開始服務、立刻開瀏覽器，最後才開始讀取資料。
    # （讀取 data.xlsx 很吃 CPU，若先讀，計時器開瀏覽器的動作會被拖到讀完才執行，造成「算完又多開一個分頁」）
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    if "--no-browser" not in sys.argv: open_url(url)
    threading.Thread(target=load_all, daemon=True).start()
    try:
        while True: time.sleep(3600)
    except KeyboardInterrupt: pass


if __name__ == "__main__": main()
