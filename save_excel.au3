; 開啟 Excel
Local $oExcel = ObjCreate("Excel.Application")
$oExcel.Visible = True

; 開啟檔案
Local $filePath = "D:\flask\data.xlsx"
Local $oWorkbook = $oExcel.Workbooks.Open($filePath)

; 等待 2 秒（可視需要調整，讓使用者看到打開）
Sleep(2000)

; 儲存與關閉
$oWorkbook.Save()
$oWorkbook.Close(False)
$oExcel.Quit()

; 清除物件
$oWorkbook = 0
$oExcel = 0

Exit 0  ; 成功時
