import win32com.client

# ===== Excel 常數 =====
xlAscending = 1
xlDescending = 2
xlYes = 1
xlNo = 2
xlCalculationManual = -4135
xlCalculationAutomatic = -4105
xlUp = -4162

# ===== 安全取得工作表 =====
def get_sheet_safe(workbook, name):
    try:
        return workbook.Sheets(name)
    except Exception:
        print(f"[跳過] 找不到工作表: {name}")
        return None

# ===== 抓最後一列（穩定版）=====
def get_last_row(sheet, col):
    return sheet.Cells(sheet.Rows.Count, col).End(xlUp).Row

# ===== 排序核心 =====
def sort_sheet(sheet, start_row, end_col, key_cols, orders, name, has_header=False):
    try:
        if sheet is None:
            return

        first_col = key_cols[0] if isinstance(key_cols, list) else key_cols
        last_row = get_last_row(sheet, first_col)

        if last_row < start_row:
            print(f"[跳過] {name} 無資料可排序")
            return

        full_range = f"A{start_row}:{end_col}{last_row}"

        sheet.Sort.SortFields.Clear()

        # 多欄排序
        if isinstance(key_cols, list):
            for col, order in zip(key_cols, orders):
                key_range = f"{col}{start_row}:{col}{last_row}"
                sheet.Sort.SortFields.Add(
                    Key=sheet.Range(key_range),
                    SortOn=0,
                    Order=order,
                    DataOption=0
                )
        else:
            key_range = f"{key_cols}{start_row}:{key_cols}{last_row}"
            sheet.Sort.SortFields.Add(
                Key=sheet.Range(key_range),
                SortOn=0,
                Order=orders,
                DataOption=0
            )

        sheet.Sort.SetRange(sheet.Range(full_range))
        sheet.Sort.Header = xlYes if has_header else xlNo
        sheet.Sort.Apply()

        print(f"[成功] {name} 排序完成（列 {start_row} ~ {last_row}）")

    except Exception as e:
        print(f"[失敗] {name} 排序錯誤: {e}")

# ===== 主程式 =====
def main():
    excel = None
    workbook = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")

        # ===== 效能優化 =====
        excel.Visible = False
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False
        try:
            excel.Calculation = xlCalculationManual
        except:
            print("[警告] 無法設定 Calculation（可能 Excel 狀態異常）")

        file_path = r"D:\flask2\data.xlsx"
        workbook = excel.Workbooks.Open(file_path)

        # ===== 排序設定 =====
        configs = [
            ("門市主檔", 22, "AC", "J", xlAscending, False, "門市主檔（J 由舊到新）"),
            ("吳宗鴻", 8, "Z", ["F","L","S"], [xlDescending,xlDescending,xlDescending], False, "吳宗鴻（F→L→S）"),
            ("湯家瑋", 8, "Z", ["F","L","S"], [xlDescending,xlDescending,xlDescending], False, "湯家瑋（F→L→S）"),
            ("劉柏均", 8, "Z", ["F","L","S"], [xlDescending,xlDescending,xlDescending], False, "劉柏均（F→L→S）"),
            ("狄澤洋", 8, "Z", ["F","L","S"], [xlDescending,xlDescending,xlDescending], False, "狄澤洋（F→L→S）"),
        ]

        # ===== 執行排序（← 已修正縮排）=====
        for sheet_name, start_row, end_col, key_cols, orders, has_header, desc in configs:
            sheet = get_sheet_safe(workbook, sheet_name)
            sort_sheet(sheet, start_row, end_col, key_cols, orders, desc, has_header)

        # ===== 儲存 =====
        workbook.Save()
        print("[完成] 所有排序完成，檔案已儲存")

    except Exception as e:
        print(f"[重大錯誤] {e}")

    finally:
        if excel:
            try:
                excel.ScreenUpdating = True
                excel.DisplayAlerts = True
                excel.Calculation = xlCalculationAutomatic
            except:
                print("[警告] Excel 已異常，略過還原設定")

        if workbook:
            try:
                workbook.Close(False)
            except:
                print("[警告] workbook.Close 失敗")

        if excel:
            try:
                excel.Quit()
            except:
                print("[警告] excel.Quit 失敗")

        print("[完成] Excel 已關閉")

# ===== 執行 =====
if __name__ == "__main__":
    main()