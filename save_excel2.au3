; 開啟 Excel（只開一次）
Local $oExcel = ObjCreate("Excel.Application")
$oExcel.Visible = True

; 檔案清單
Local $files[3] = [ _
    "D:\flask2\data.xlsx", _
    "D:\flask2\MFP\MFP.xlsx", _
    "D:\flask2\MFP\output.xlsx" _
]

For $i = 0 To UBound($files) - 1
    Local $filePath = $files[$i]

    ; 開啟
    Local $oWorkbook = $oExcel.Workbooks.Open($filePath)

    ; ✅ 等 Excel 完全載入（取代 Sleep）
    While Not $oExcel.Ready
        Sleep(200)
    WEnd

    ; 再補一點緩衝（避免極少數狀況）
    Sleep(300)

    ; 儲存 + 關閉
    $oWorkbook.Save()
    $oWorkbook.Close(False)

    ; 清掉 Workbook
    $oWorkbook = 0
Next

; 關閉 Excel（最後才關）
$oExcel.Quit()
$oExcel = 0

Exit 0