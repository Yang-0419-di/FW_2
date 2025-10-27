import sqlite3

conn = sqlite3.connect(r"D:\flask\billing.db")
c = conn.cursor()

# 查看有沒有那筆 T352500089
c.execute("SELECT * FROM contracts WHERE device_id = 'T352500089';")
rows = c.fetchall()

print(f"筆數：{len(rows)}")
for r in rows:
    print(r)

conn.close()
