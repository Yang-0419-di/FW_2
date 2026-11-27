# line_remind.py
# ===== 個人行事曆 LINE 提醒（每天由 GitHub Actions 執行一次）=====
# 規則：事件有設定「提前通知」N 天（0＝當日）時，從「事件日期 − N 天」起到事件當天之間，
#       若尚未通知，就在當天的排程送出一次（因此排程某天沒跑成，隔天仍會補送）。
#       每位人員每天最多一則訊息（當天要提醒的事件合併在同一則），以節省每月免費則數。
# 送出後在「個人行事曆」分頁的「已通知」欄寫入「日期|天數」，避免重複提醒。
#
# 需要的環境變數：
#   GOOGLE_SERVICE_ACCOUNT_FILE   Google 服務帳號 JSON 檔路徑（GitHub Actions 會由 Secret 產生）
#   LINE_CHANNEL_ACCESS_TOKEN     LINE Messaging API 的 Channel access token (long-lived)
#   DRY_RUN=1                     只顯示會送出的內容，不發送、不寫入（測試用）
#   TODAY=YYYY-MM-DD              指定「今天」（測試用，可省略）
import os
import sys
from collections import defaultdict
from datetime import datetime, timedelta, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from modules import personal_calendar as PC   # noqa: E402
from modules import line_bind, line_api        # noqa: E402

TW = timezone(timedelta(hours=8))
WEEK = "一二三四五六日"


def md(d):
    return f"{d.month}/{d.day}（{WEEK[d.weekday()]}）"


def when(delta):
    return {0: "今天", 1: "明天", 2: "後天"}.get(delta, f"{delta} 天後")


def title_of(r):
    # 與網站顯示相同：店編_店名_內容_時間
    return "_".join(x for x in (r[3], r[4], r[5], r[6]) if x) or "（未命名）"


def build_message(person, today, items):
    lines = [f"📅 個人行事曆提醒（{person}）", f"今天是 {md(today)}"]
    by_day = defaultdict(list)
    for it in items:
        by_day[it["date"]].append(it)
    for d in sorted(by_day):
        lines.append("")
        lines.append(f"【{when((d - today).days)}】{md(d)}")
        for it in by_day[d]:
            lines.append(("⭐ " if it["special"] else "・") + it["title"])
    return "\n".join(lines)


def main():
    dry = os.environ.get("DRY_RUN", "") == "1"
    today = (datetime.strptime(os.environ["TODAY"], "%Y-%m-%d").date() if os.environ.get("TODAY")
             else datetime.now(TW).date())
    print(f"=== 個人行事曆 LINE 提醒：{today}{'（DRY_RUN：只預覽，不發送）' if dry else ''} ===")

    ws = PC._worksheet()
    values = ws.get_all_values()
    binds = line_bind.bindings()
    print(f"已綁定 LINE 的人員：{'、'.join(binds) or '（無）'}")

    due = defaultdict(list)
    for r in values[1:]:
        r = [x.strip() for x in r] + [""] * (len(PC.HEADER) - len(r))
        event_id, person, notify, notified = r[0], r[1], r[12], r[13]
        if not event_id or notify == "":
            continue
        try:
            days = int(notify)
            d = datetime.strptime(r[2][:10], "%Y-%m-%d").date()
        except ValueError:
            print(f"  略過格式錯誤的事件 {event_id}：日期={r[2]!r} 提前通知={notify!r}")
            continue
        key = f"{d.isoformat()}|{days}"
        if notified == key:
            continue                                   # 已通知過
        if not (d - timedelta(days=days) <= today <= d):
            continue                                   # 還沒到通知日，或事件已過
        due[person].append({"id": event_id, "date": d, "key": key, "title": title_of(r),
                            "special": r[7].upper() == "Y"})

    if not due:
        print("今天沒有需要提醒的事件。")
        return 0

    failed = False
    for person, items in sorted(due.items()):
        items.sort(key=lambda x: (x["date"], x["title"]))
        uid = binds.get(person)
        print(f"\n--- {person}：{len(items)} 筆 ---")
        if not uid:
            print(f"  ⚠ {person} 尚未綁定 LINE，略過（綁定後隔天會補送仍在期限內的提醒）")
            continue
        text = build_message(person, today, items)
        print(text)
        if dry:
            continue
        try:
            line_api.push(uid, text)
        except Exception as e:
            failed = True
            print(f"  ❌ 發送失敗：{e}")
            continue
        # 送出後標記「已通知」（重新讀 id → 列號，避免期間有人刪除事件造成列號位移）
        ids = ws.col_values(1)
        pos = {v.strip(): i + 1 for i, v in enumerate(ids) if i > 0 and v.strip()}
        updates = [{"range": f"N{pos[it['id']]}", "values": [[it["key"]]]} for it in items if it["id"] in pos]
        if updates:
            ws.batch_update(updates, value_input_option="RAW")
        print(f"  ✅ 已送出並標記 {len(updates)} 筆")

    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
