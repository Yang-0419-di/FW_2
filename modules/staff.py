# modules/staff.py
# ===== 部門人員設定（唯一來源）=====
# 增減人員或改顏色只要改這裡：排程表顏色、圖例、人員2 下拉建議、個人頁白名單都會跟著變。
#   name    ：顯示名稱，也是 data.xlsx 工程師分頁名稱
#   color   ：排程表事件顏色（CSS 顏色）
#   aliases ：「行事曆」D 欄（屬性）出現這些代號時，視為此人
#   sheet   ：data.xlsx 是否有此人的個人分頁（/personal 頁面用）

STAFF = [
    {"name": "狄澤洋", "color": "red",     "aliases": ["V"], "sheet": True},
    {"name": "湯家瑋", "color": "green",   "aliases": [],    "sheet": True},
    {"name": "吳宗鴻", "color": "orange",  "aliases": [],    "sheet": True},
    {"name": "劉柏均", "color": "skyblue", "aliases": [],    "sheet": True},
]

DEFAULT_COLOR = "blue"   # 屬性欄空白或不在名單內
PAST_COLOR = "gray"      # 已過期事件

STAFF_NAMES = [s["name"] for s in STAFF]
SHEET_NAMES = [s["name"] for s in STAFF if s.get("sheet")]

# 屬性欄的值（含代號）→ 人員設定
_LOOKUP = {}
for _s in STAFF:
    _LOOKUP[_s["name"]] = _s
    for _a in _s.get("aliases", []):
        _LOOKUP[_a] = _s


def find_staff(value):
    """依『屬性』欄的值找人員；找不到回傳 None。"""
    if value is None:
        return None
    return _LOOKUP.get(str(value).strip())


def color_of(value):
    s = find_staff(value)
    return s["color"] if s else DEFAULT_COLOR


def display_name(value):
    """代號轉成正式姓名（例如 V → 狄澤洋）；不在名單內則原樣顯示。"""
    s = find_staff(value)
    if s:
        return s["name"]
    return "" if value is None else str(value).strip()
