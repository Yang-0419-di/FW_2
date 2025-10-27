import tkinter as tk
from tkinter import messagebox
import subprocess

def run_selected():
    selected = [var.get() for var in step_vars]
    tags = date_entry.get().strip()

    if selected[1] and not tags:
        messagebox.showerror("錯誤", "請輸入步驟 2 的年月")
        return
    if selected[2] and not tags:
        messagebox.showerror("錯誤", "請輸入步驟 3 的年月")
        return

    try:
        if selected[0]:
            subprocess.run(["python", "Excel_Edge.py"], check=True)

        if selected[1]:
            subprocess.run(["python", "run_update2.py"], input=tags.encode(), check=True)

        if selected[2]:
            subprocess.run(["python", "run_MFP_update.py"], input=tags.encode(), check=True)

        if selected[3]:
            version = ver_entry.get().strip()
            if not version:
                messagebox.showerror("錯誤", "請輸入版本號")
                return
            subprocess.run(["python", "add_ver.py", version], check=True)

        if selected[4]:
            subprocess.run(["python", "data_updw.py"], check=True)

        if selected[5]:
            subprocess.run(["save_excel.exe"], check=True)

        messagebox.showinfo("完成", "所有選擇的步驟已完成！")

    except subprocess.CalledProcessError as e:
        messagebox.showerror("錯誤", f"執行失敗：{e}")

app = tk.Tk()
app.title("自動化任務控制台")

step_names = [
    "Step 1：下載 Excel 報表 (Excel_Edge.py)",
    "Step 2：執行 run_update2.py",
    "Step 3：執行 run_MFP_update.py",
    "Step 4：寫入版本號 (add_ver.py)",
    "Step 5：執行 data_updw.py",
    "Step 6：啟動 save_excel.exe"
]

step_vars = [tk.BooleanVar() for _ in step_names]

for i, name in enumerate(step_names):
    tk.Checkbutton(app, text=name, variable=step_vars[i]).pack(anchor='w')

tk.Label(app, text="輸入年月（步驟 2 & 3）：").pack(anchor='w')
date_entry = tk.Entry(app)
date_entry.pack(fill='x')

tk.Label(app, text="輸入版本號（步驟 4）：").pack(anchor='w')
ver_entry = tk.Entry(app)
ver_entry.pack(fill='x')

tk.Button(app, text="開始執行", command=run_selected).pack(pady=10)

app.mainloop()
