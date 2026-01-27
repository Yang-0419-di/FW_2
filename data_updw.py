import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

file_path = r"D:\flask2\data.xlsx"
workbook = excel.Workbooks.Open(file_path)

xlAscending = 1
xlDescending = 2

def sort_sheet(sheet, start_row, end_col, key_col, order, name, has_header=False):
    """
    sheet: 工作表對象
    start_row: 排序範圍起始列號 (整數)
    end_col: 排序範圍結束欄 (字母，如 "AC")
    key_col: 排序鍵欄 (字母，如 "J")
    order: xlAscending 或 xlDescending
    name: 工作表名稱（僅用於打印）
    has_header: 是否有標題列
    """
    try:
        # 計算最後一列
        last_row = sheet.UsedRange.Rows.Count
        first_used_row = sheet.UsedRange.Row  # 實際第一列（避免空白列影響）
        actual_last_row = last_row + first_used_row - 1

        # 完整排序範圍
        full_range = f"A{start_row}:{end_col}{actual_last_row}"
        key_range = f"{key_col}{start_row}:{key_col}{actual_last_row}"

        # 清空原有排序
        sheet.Sort.SortFields.Clear()
        sheet.Sort.SortFields.Add(
            Key=sheet.Range(key_range),
            SortOn=0,   # xlSortOnValues
            Order=order,
            DataOption=0
        )
        sheet.Sort.SetRange(sheet.Range(full_range))
        sheet.Sort.Header = 1 if has_header else 0
        sheet.Sort.Apply()

        print(f"[成功] {name} 排序完成")
    except Exception as e:
        print(f"[失敗] {name} 排序錯誤: {e}")

# 範例排序
sort_sheet(workbook.Sheets("門市主檔"), 23, "AC", "J", xlAscending, "門市主檔（J 從最舊到最新）")
sort_sheet(workbook.Sheets("吳宗鴻"), 7, "Z", "F", xlDescending, "吳宗鴻（F 從 Z 到 A）")
sort_sheet(workbook.Sheets("湯家瑋"), 7, "Z", "F", xlDescending, "湯家瑋（F 從 Z 到 A）")
sort_sheet(workbook.Sheets("劉柏均"), 7, "Z", "F", xlDescending, "劉柏均（F 從 Z 到 A）")
sort_sheet(workbook.Sheets("狄澤洋"), 7, "Z", "F", xlDescending, "狄澤洋（F 從 Z 到 A）")

# 儲存關閉
workbook.Save()
workbook.Close(False)
excel.Quit()
print("[完成] 所有排序完成，檔案已儲存並關閉。")
