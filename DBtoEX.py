import sqlite3
import pandas as pd

# 連線 SQLite
conn = sqlite3.connect("billing.db")

# 要匯出的表格清單
tables = ["contracts", "customers"]

# 使用 ExcelWriter 將多個表格寫入同一個 Excel 檔，不同工作表
with pd.ExcelWriter("output.xlsx") as writer:
    for table in tables:
        df = pd.read_sql(f"SELECT * FROM {table}", conn)
        df.to_excel(writer, sheet_name=table, index=False)

conn.close()
