import sqlite3
import pandas as pd

# 連線 SQLite
conn = sqlite3.connect("billing.db")

# 讀取 Excel 的所有工作表
xls = pd.ExcelFile("output.xlsx")
for sheet_name in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name, dtype={'tax_id': str})
    
    # 將 DataFrame 寫入 SQLite 表格
    # 如果表格已存在，可以選擇 'replace' 覆蓋，或 'append' 追加
    df.to_sql(sheet_name, conn, if_exists='replace', index=False)

conn.close()
