import sqlite3

conn = sqlite3.connect("billing.db")
c = conn.cursor()

# 如果欄位不存在，才新增
try:
    c.execute("ALTER TABLE billing_summary ADD COLUMN last_date TEXT")
    print("已新增欄位 last_date")
except sqlite3.OperationalError as e:
    if "duplicate column name" in str(e):
        print("欄位 last_date 已存在")
    else:
        raise

conn.commit()
conn.close()
