@echo off
chcp 65001 >nul
:loop
cd /d D:\flask2

echo ===== 第 %count% 次啟動 Flask =====
set /a count+=1

REM 使用 start /wait 啟動 Python，Ctrl+C 停止 Flask 不會跳詢問
start "" /wait python app.py

echo.
echo ===== Flask 已停止 =====
echo 按 R 重新啟動
echo 按其他任意鍵離開...

set /p key=請輸入選擇: 

if /i "%key%"=="r" goto loop
echo 離開程序
pause
exit