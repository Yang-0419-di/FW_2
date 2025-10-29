import sqlite3

conn = sqlite3.connect("billing.db")
c = conn.cursor()
c.execute("PRAGMA table_info(invoice_log);")
print(c.fetchall())
conn.close()
