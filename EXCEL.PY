import os
import re
import subprocess
import platform
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime

# 定義標準 36 個標題欄位
TARGET_COLUMNS = [
    "項次", "工作類別", "台芝工作案號", "客戶工作案號", "案件類別", 
    "門店編號", "門店名稱", "縣市鄉鎮", "報修時間", "星期", 
    "時段", "報修類別代碼", "報修類別", "報修項目代碼", "報修項目", 
    "報修說明", "設備號碼", "負責部門", "服務人員", "到場次數", 
    "完工結果", "派修時間", "到場時間", "離場時間", "完修時限", 
    "處理工時", "逾時時間", "結案類別", "工作內容", "保養", 
    "裝/換徹機", "耗材自送", "耗材配送", "發票收款", "合約事宜", 
    "其它", "客戶簽章"
]

def clean_str(s):
    """
    強效字串淨化：
    去除 \r, \n, \t, 全半形空格、引號等所有不可見符號與換行，僅保留純文字做比對
    """
    if pd.isna(s):
        return ""
    text = str(s)
    text = re.sub(r'[\r\n\t\s]+', '', text)
    return text

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 案件資料自動彙整器")
        self.root.geometry("760x600")
        self.root.resizable(True, True)

        self.file_paths = []
        self.last_saved_path = None  # 記錄最新匯出的檔案路徑

        self.create_widgets()

    def create_widgets(self):
        # 標題區
        header_frame = tk.Frame(self.root, bg="#2c3e50", pady=12)
        header_frame.pack(fill=tk.X)
        
        title_label = tk.Label(
            header_frame, 
            text="📊 Excel 多檔資料自動彙整器", 
            font=("Microsoft JhengHei", 16, "bold"), 
            fg="white", 
            bg="#2c3e50"
        )
        title_label.pack()

        # 第一列按鈕操作區（主要操作）
        btn_frame1 = tk.Frame(self.root, pady=8)
        btn_frame1.pack(fill=tk.X, padx=15)

        self.btn_select = tk.Button(
            btn_frame1, 
            text="📁 選擇/瀏覽 Excel 檔案", 
            font=("Microsoft JhengHei", 11, "bold"),
            bg="#3498db", 
            fg="white", 
            relief=tk.RAISED,
            padx=10, pady=5,
            command=self.select_files
        )
        self.btn_select.pack(side=tk.LEFT, padx=5)

        self.btn_clear = tk.Button(
            btn_frame1, 
            text="🗑️ 清空清單", 
            font=("Microsoft JhengHei", 10),
            bg="#e74c3c", 
            fg="white",
            relief=tk.RAISED,
            padx=10, pady=5,
            command=self.clear_files
        )
        self.btn_clear.pack(side=tk.LEFT, padx=5)

        self.btn_process = tk.Button(
            btn_frame1, 
            text="⚡ 開始彙整並匯出", 
            font=("Microsoft JhengHei", 11, "bold"),
            bg="#2ecc71", 
            fg="white",
            relief=tk.RAISED,
            padx=15, pady=5,
            command=self.process_files
        )
        self.btn_process.pack(side=tk.RIGHT, padx=5)

        # 第二列按鈕操作區（匯出後的快速開啟功能）
        btn_frame2 = tk.Frame(self.root, pady=2)
        btn_frame2.pack(fill=tk.X, padx=15)

        self.btn_open_folder = tk.Button(
            btn_frame2, 
            text="📂 開啟存檔位置", 
            font=("Microsoft JhengHei", 10, "bold"),
            bg="#f39c12", 
            fg="white",
            state=tk.DISABLED,  # 匯出前預設停用
            relief=tk.RAISED,
            padx=10, pady=4,
            command=self.open_saved_folder
        )
        self.btn_open_folder.pack(side=tk.RIGHT, padx=5)

        self.btn_open_file = tk.Button(
            btn_frame2, 
            text="📊 開啟存檔 Excel", 
            font=("Microsoft JhengHei", 10, "bold"),
            bg="#9b59b6", 
            fg="white",
            state=tk.DISABLED,  # 匯出前預設停用
            relief=tk.RAISED,
            padx=10, pady=4,
            command=self.open_saved_excel
        )
        self.btn_open_file.pack(side=tk.RIGHT, padx=5)

        # 檔案清單顯示區
        list_frame = tk.LabelFrame(self.root, text=" 待彙整 Excel 檔案清單 ", font=("Microsoft JhengHei", 10, "bold"), padx=10, pady=10)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.file_listbox = tk.Listbox(
            list_frame, 
            selectmode=tk.EXTENDED, 
            font=("Consolas", 10),
            yscrollcommand=scrollbar.set
        )
        self.file_listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)

        # 狀態訊息日誌區
        log_frame = tk.LabelFrame(self.root, text=" 執行狀態與訊息 Log ", font=("Microsoft JhengHei", 10, "bold"), padx=10, pady=5)
        log_frame.pack(fill=tk.X, padx=15, pady=5)

        self.log_text = tk.Text(log_frame, height=7, font=("Microsoft JhengHei", 9), state=tk.DISABLED)
        self.log_text.pack(fill=tk.X)

    def log(self, message):
        """更新狀態日誌訊息"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def select_files(self):
        """選擇多個 Excel 檔案"""
        files = filedialog.askopenfilenames(
            title="選擇要彙整的 Excel 檔案",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if files:
            for f in files:
                if f not in self.file_paths:
                    self.file_paths.append(f)
                    self.file_listbox.insert(tk.END, f)
            self.log(f"已新增 {len(files)} 個檔案，當前共有 {len(self.file_paths)} 個待處理檔案。")

    def clear_files(self):
        """清空選擇的檔案清單"""
        self.file_paths.clear()
        self.file_listbox.delete(0, tk.END)
        self.log("清單已清空。")

    def open_saved_folder(self):
        """開啟最新存檔檔案所在的位置（檔案總管）"""
        if not self.last_saved_path or not os.path.exists(self.last_saved_path):
            messagebox.showwarning("警告", "找不到存檔檔案或尚未進行匯出！")
            return

        folder_path = os.path.dirname(self.last_saved_path)
        system_platform = platform.system()

        try:
            if system_platform == "Windows":
                # 在 Windows 下開啟並選取該檔案
                subprocess.run(f'explorer /select,"{os.path.normpath(self.last_saved_path)}"')
            elif system_platform == "Darwin":  # macOS
                subprocess.run(["open", "-R", self.last_saved_path])
            else:  # Linux
                subprocess.run(["xdg-open", folder_path])
            self.log("已開啟存檔資料夾。")
        except Exception as e:
            self.log(f"無法開啟資料夾: {str(e)}")

    def open_saved_excel(self):
        """直接使用預設應用程式開啟儲存的 Excel 檔案"""
        if not self.last_saved_path or not os.path.exists(self.last_saved_path):
            messagebox.showwarning("警告", "找不到存檔檔案或尚未進行匯出！")
            return

        try:
            os.startfile(self.last_saved_path)  # Windows 專用
            self.log("已開啟存檔 Excel 檔案。")
        except AttributeError:
            # 跨平台相容處理 (macOS / Linux)
            system_platform = platform.system()
            if system_platform == "Darwin":
                subprocess.run(["open", self.last_saved_path])
            else:
                subprocess.run(["xdg-open", self.last_saved_path])
            self.log("已開啟存檔 Excel 檔案。")
        except Exception as e:
            self.log(f"無法開啟檔案: {str(e)}")
            messagebox.showerror("開啟失敗", f"無法開啟檔案：\n{str(e)}")

    def load_and_standardize_excel(self, file_path):
        """讀取 Excel，自動剔除標題換行並精確對齊 36 個欄位"""
        excel_file = pd.ExcelFile(file_path)
        combined_sheets_df = []

        for sheet_name in excel_file.sheet_names:
            raw_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            if raw_df.empty:
                continue

            header_row_idx = None
            for idx, row in raw_df.iterrows():
                row_str = "".join([clean_str(val) for val in row.values])
                if any(k in row_str for k in ["派修時間", "工作類別", "門店名稱", "台芝工作案號", "報修說明"]):
                    header_row_idx = idx
                    break

            if header_row_idx is None:
                header_row_idx = 0

            df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row_idx)

            col_cleaned_map = {}
            for original_col in df.columns:
                cleaned_col = clean_str(original_col)
                if cleaned_col:
                    col_cleaned_map[cleaned_col] = original_col

            new_df = pd.DataFrame()
            matched_count = 0

            for target_col in TARGET_COLUMNS:
                target_clean = clean_str(target_col)
                if target_clean in col_cleaned_map:
                    orig_col_name = col_cleaned_map[target_clean]
                    new_df[target_col] = df[orig_col_name]
                    matched_count += 1
                else:
                    new_df[target_col] = None

            new_df = new_df.dropna(how='all')

            if not new_df.empty and matched_count > 0:
                combined_sheets_df.append(new_df)

        if combined_sheets_df:
            return pd.concat(combined_sheets_df, ignore_index=True)
        else:
            return pd.DataFrame()

    def process_files(self):
        if not self.file_paths:
            messagebox.showwarning("警告", "請先點擊『選擇/瀏覽 Excel 檔案』加入至少一個 Excel 檔！")
            return

        save_path = filedialog.asksaveasfilename(
            title="另存新檔彙整結果",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not save_path:
            return

        try:
            self.log("開始讀取與分析 Excel 檔案...")
            df_list = []

            for idx, file_path in enumerate(self.file_paths, start=1):
                filename = os.path.basename(file_path)
                self.log(f"({idx}/{len(self.file_paths)}) 解析中: {filename}")
                
                df = self.load_and_standardize_excel(file_path)

                if df.empty:
                    self.log(f"⚠️ 警告：[{filename}] 未能讀取到有效資料，請檢查欄位。")
                else:
                    self.log(f"  └─ 成功讀取 {len(df)} 筆資料")
                    df_list.append(df)

            if not df_list:
                messagebox.showerror("錯誤", "未能從所選檔案中讀取到任何有效資料！請確認檔案欄位。")
                self.log("❌ 彙整終止：無任何有效資料。")
                return

            # 1. 合併所有資料表
            combined_df = pd.concat(df_list, ignore_index=True)
            self.log(f"資料合併完成，共取得 {len(combined_df)} 筆原始紀錄。")

            # 2. 轉換日期時間格式以進行正確排序
            self.log("正在解析【派修時間】與【離場時間】進行排序...")
            combined_df["_sort_dispatch"] = pd.to_datetime(combined_df["派修時間"], errors="coerce")
            combined_df["_sort_leave"] = pd.to_datetime(combined_df["離場時間"], errors="coerce")

            # 3. 雙重排序：先依【派修時間】(舊到新)，再依【離場時間】(舊到新)
            combined_df = combined_df.sort_values(
                by=["_sort_dispatch", "_sort_leave"], 
                ascending=[True, True],
                na_position='last'
            )

            # 移除輔助排序欄位
            combined_df = combined_df.drop(columns=["_sort_dispatch", "_sort_leave"])

            # 4. 重新編號『項次』
            combined_df["項次"] = range(1, len(combined_df) + 1)

            # 5. 寫入 Excel 並美化格式
            self.log("正在寫入 Excel 並設置美化排版...")
            with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                combined_df.to_excel(writer, index=False, sheet_name="彙整結果")
                
                workbook = writer.book
                worksheet = writer.sheets["彙整結果"]

                # 📌 1. 設定凍結頂端列 (第一列標題)
                worksheet.freeze_panes = "A2"

                from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

                header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                header_font = Font(name="微軟正黑體", size=10, bold=True)
                cell_font = Font(name="微軟正黑體", size=9)
                align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
                thin_border = Border(
                    left=Side(style="thin", color="D3D3D3"),
                    right=Side(style="thin", color="D3D3D3"),
                    top=Side(style="thin", color="D3D3D3"),
                    bottom=Side(style="thin", color="D3D3D3")
                )

                total_rows = len(combined_df) + 1

                # 📌 2. 設定所有列的高 (列高 = 21.66)
                for row_idx in range(1, total_rows + 1):
                    worksheet.row_dimensions[row_idx].height = 21.66

                # 設定標題樣式
                for col_idx, col_name in enumerate(TARGET_COLUMNS, 1):
                    cell = worksheet.cell(row=1, column=col_idx)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = align_center
                    cell.border = thin_border

                # 設定資料內容列樣式
                for row in worksheet.iter_rows(min_row=2, max_row=total_rows, min_col=1, max_col=len(TARGET_COLUMNS)):
                    for cell in row:
                        cell.font = cell_font
                        cell.border = thin_border

                # 自動調整欄寬
                for col in worksheet.columns:
                    max_len = 0
                    col_letter = col[0].column_letter
                    for cell in col:
                        val_str = str(cell.value or '')
                        cell_len = sum(2 if ord(c) > 127 else 1 for c in val_str)
                        if cell_len > max_len:
                            max_len = cell_len
                    worksheet.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

            # 記錄成功儲存的路徑並啟用快捷按鈕
            self.last_saved_path = save_path
            self.btn_open_folder.config(state=tk.NORMAL)
            self.btn_open_file.config(state=tk.NORMAL)

            self.log(f"🎉 成功匯出 {len(combined_df)} 筆資料至檔案！")
            messagebox.showinfo("成功", f"Excel 彙整完成！\n共成功匯出：{len(combined_df)} 筆資料\n儲存路徑：\n{save_path}")

        except Exception as e:
            self.log(f"❌ 錯誤: {str(e)}")
            messagebox.showerror("彙整失敗", f"處理過程中發生錯誤：\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()