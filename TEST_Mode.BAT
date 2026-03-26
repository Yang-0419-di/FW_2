@echo off
:loop
cd /d D:\flask2

echo 啟動 Flask...
python app.py

echo.
echo ===== Flask 已停止 =====
echo 按 R 重新啟動（刷新Excel/HTML）
echo 按其他鍵離開
choice /c RQ /n /m "請選擇: "

if errorlevel 2 exit
if errorlevel 1 goto loop
