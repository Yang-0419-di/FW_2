@echo off
setlocal enabledelayedexpansion

cd /d D:\flask2

:: -----------------------------
:: 詢問是否直接開始 Git 操作
:: -----------------------------
set /p skipAll=是否直接開始 Git 操作 [Y/N]：
if /i "!skipAll!"=="Y" (
    goto git_only
)

:: -----------------------------
:: 詢問是否跳過下載檔案
:: -----------------------------
set /p skipDownload=是否跳過下載檔案 [Y/N]：
if /i "!skipDownload!"=="Y" (
    echo [1/7] 已選擇跳過下載檔案。
) else (
    echo [1/7] 執行 Excel_Edge.py...
    python "Excel_Edge.py"
    if errorlevel 1 (
        echo ? 檔案下載異常
        pause
        exit /b
    )
)

:: -----------------------------
:: [2/7] 備份檔案
:: -----------------------------
echo [2/7] 備份 data.xlsx...
copy /Y "data.xlsx" "臨時備份\自動備份\data.xlsx"
if errorlevel 1 goto error

echo [2/7] 備份 MFP.xlsx...
copy /Y "MFP\MFP.xlsx" "MFP\自動備份\MFP.xlsx"
if errorlevel 1 goto error

echo [2/7] 備份 output.xlsx...
copy /Y "MFP\output.xlsx" "MFP\自動備份\output.xlsx"
if errorlevel 1 goto error

:: -----------------------------
:: [3/7] run_update3（合併版）
:: -----------------------------
echo [3/7] 執行 run_update3.py...
python "run_update3.py"
if errorlevel 1 goto error

:: -----------------------------
:: [4/7] run_MFP_update3（合併版）
:: -----------------------------
echo [4/7] 執行 run_MFP_update3.py...
python "run_MFP_update3.py"
if errorlevel 1 goto error

:: -----------------------------
:: [5/7] add_ver
:: -----------------------------
echo [5/7] 執行 add_ver.py（10 秒後自動填入）...
call python "add_ver.py"

for /f %%a in ('python -c "import openpyxl;wb=openpyxl.load_workbook('data.xlsx');print(wb['首頁']['G1'].value)"') do set vernum=%%a
echo 使用版本號: %vernum%

:: -----------------------------
:: [6/7] data_updw
:: -----------------------------
echo [6/7] 執行 data_updw.py...
python "data_updw.py"
if errorlevel 1 goto error

:: -----------------------------
:: [7/7] Excel 強制儲存
:: -----------------------------
echo 等待 Excel 寫入穩定...
timeout /t 10 /nobreak >nul

echo [7/7] 啟動 save_excel2.exe...
start /wait "" "save_excel2.exe"
if errorlevel 1 goto error

:: -----------------------------
:: Git 前等待（很重要）
:: -----------------------------
echo 等待資料完全寫入磁碟...
timeout /t 10 /nobreak >nul

goto git_operation


:: -----------------------------
:: Git-only 模式
:: -----------------------------
:git_only

for /f %%a in ('python "add_ver.py" --auto') do set vernum=%%a
echo 使用版本號: %vernum%

echo [1/3] 儲存 Excel...
start /wait "" "save_excel2.exe"
if errorlevel 1 goto error

echo 等待 Excel 完全釋放...
timeout /t 5 /nobreak >nul

echo [2/3] Git 操作中...
cd /d D:\flask2
git pull
if errorlevel 1 goto error

git add -A
git commit -m "Auto commit - %vernum%"
if errorlevel 1 goto error

git push
if errorlevel 1 goto error

echo [3/3] 完成！
pause
exit /b


:: -----------------------------
:: Git 主流程
:: -----------------------------
:git_operation

echo [Git] 進行 Git 操作...
cd /d D:\flask2

git pull
if errorlevel 1 goto error

git add -A
git commit -m "Auto commit - %vernum%"
if errorlevel 1 goto error

git push
if errorlevel 1 goto error

echo ? 所有程序已完成！
pause
exit /b


:: -----------------------------
:: 錯誤處理
:: -----------------------------
:error
echo ? 錯誤：某個步驟執行失敗，流程中止。
pause
exit /b