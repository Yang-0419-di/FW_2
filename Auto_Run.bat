@echo off
setlocal enabledelayedexpansion

cd /d D:\flask2

:: -----------------------------
:: 是否直接 Git
:: -----------------------------
set /p skipAll=是否直接開始 Git 操作 [Y/N]：
if /i "!skipAll!"=="Y" goto git_only

:: -----------------------------
:: 是否跳過下載
:: -----------------------------
set /p skipDownload=是否跳過下載檔案 [Y/N]：
if /i "!skipDownload!"=="Y" (
    echo [1/6] 已跳過下載
) else (
    echo [1/6] 執行 Excel_Edge.py...
    python "Excel_Edge.py"
    if errorlevel 1 (
        echo ? 檔案下載異常
        pause
        exit /b
    )
)

:: -----------------------------
:: [2/6] 備份
:: -----------------------------
echo [2/6] 備份檔案...

copy /Y "data.xlsx" "臨時備份\自動備份\data.xlsx" || goto error
copy /Y "MFP\MFP.xlsx" "MFP\自動備份\MFP.xlsx" || goto error
copy /Y "MFP\output.xlsx" "MFP\自動備份\output.xlsx" || goto error

:: -----------------------------
:: [3/6] 合併更新
:: -----------------------------
echo [3/6] 執行 run_data_all.py...
python "run_data_all.py"
if errorlevel 1 goto error

:: -----------------------------
:: [4/6] 版本號
:: -----------------------------
echo [4/6] 執行 add_ver.py（10 秒自動填入）...
call python "add_ver.py"

for /f %%a in ('python -c "import openpyxl;wb=openpyxl.load_workbook('data.xlsx');print(wb['首頁']['G1'].value)"') do set vernum=%%a
echo 使用版本號: %vernum%

:: -----------------------------
:: [5/6] data_updw
:: -----------------------------
echo [5/6] 執行 data_updw.py...
python "data_updw.py"
if errorlevel 1 goto error

:: -----------------------------
:: [6/6] Excel 強制儲存
:: -----------------------------
echo 等待 Excel 穩定寫入...
timeout /t 10 /nobreak >nul

echo [6/6] 執行 save_excel2.exe...
start /wait "" "save_excel2.exe"
if errorlevel 1 goto error

echo 等待磁碟寫入完成...
timeout /t 10 /nobreak >nul

goto git_operation


:: =============================
:: Git Only 模式
:: =============================
:git_only

for /f %%a in ('python "add_ver.py" --auto') do set vernum=%%a
echo 使用版本號: %vernum%

echo [1/3] 儲存 Excel...
start /wait "" "save_excel2.exe" || goto error

timeout /t 5 /nobreak >nul

echo [2/3] Git 操作...
git pull || goto error
git add -A
git commit -m "Auto commit - %vernum%" || goto error
git push || goto error

echo [3/3] 完成！
pause
exit /b


:: =============================
:: Git 主流程
:: =============================
:git_operation

echo [Git] 執行 Git...

git pull || goto error
git add -A
git commit -m "Auto commit - %vernum%" || goto error
git push || goto error

echo ? 全部完成！
pause
exit /b


:: =============================
:: 錯誤
:: =============================
:error
echo ? 發生錯誤，流程中止
pause
exit /b