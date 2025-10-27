import win32com.client

try:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    file_path = r"D:\flask\data.xlsx"
    workbook = excel.Workbooks.Open(file_path)

    def sort_sheet(sheet, range_str, key_str, order, name):
        last_row = sheet.UsedRange.Rows.Count
        full_range = f"{range_str}{last_row}"
        key_range = f"{key_str}{last_row}"
        try:
            sheet.Sort.SortFields.Clear()
            sheet.Sort.SortFields.Add(
                Key=sheet.Range(key_range),
                SortOn=0,
                Order=order,
                DataOption=0
            )
            sheet.Sort.SetRange(sheet.Range(full_range))
            sheet.Sort.Header = 0
            sheet.Sort.Apply()
            print(f"[成功] {name} 排序完成")
        except Exception as e:
            print(f"[失敗] {name} 排序錯誤: {e}")

    sort_sheet(workbook.Sheets("門市主檔"), "A23:AC", "J15:J", 1, "門市主檔（J15 從最舊到最新）")
    sort_sheet(workbook.Sheets("吳宗鴻"), "A7:Z", "F7:F", 2, "吳宗鴻（F7 從 Z 到 A）")
    sort_sheet(workbook.Sheets("湯家瑋"), "A7:Z", "F7:F", 2, "湯家瑋（F7 從 Z 到 A）")
    sort_sheet(workbook.Sheets("狄澤洋"), "A7:Z", "F7:F", 2, "狄澤洋（F7 從 Z 到 A）")

    workbook.Save()
    workbook.Close(False)
    excel.Quit()
    print("[完成] 所有排序完成，檔案已儲存並關閉。")

except Exception as e:
    print(f"[錯誤] Excel 操作過程中發生錯誤: {e}")
