@echo off
setlocal enabledelayedexpansion

cd /d D:\flask

:: -----------------------------
:: �߰ݬO�_�����}�l Git �ާ@
:: -----------------------------
set /p skipAll=�O�_�����}�l Git �ާ@ [Y/N]�G
if /i "!skipAll!"=="Y" (
    goto git_only
)

:: -----------------------------
:: �߰ݬO�_���L�U���ɮ�
:: -----------------------------
set /p skipDownload=�O�_���L�U���ɮ� [Y/N]�G
if /i "!skipDownload!"=="Y" (
    echo [1/7] �w��ܸ��L�U���ɮסC
) else (
    echo [1/7] ���� Excel_Edge.py...
    python "Excel_Edge.py"
    if errorlevel 1 (
        echo �ɮפU�����`
        pause
        exit /b
    )
)

:: -----------------------------
:: [2/7] �ƻs data.xlsx ���{�ɳƥ�
:: -----------------------------
echo [2/7] �ƻs data.xlsx ���{�ɳƥ���Ƨ�...
copy /Y "data.xlsx" "�{�ɳƥ�\�۰ʳƥ�\data.xlsx"
if errorlevel 1 goto error

:: -----------------------------
:: [3/7] ���� run_update2.py
:: -----------------------------
echo [3/7] ���� run_update2.py...
python "run_update2.py"
if errorlevel 1 goto error

:: -----------------------------
:: [4/7] ���� run_MFP_update.py
:: -----------------------------
echo [4/7] ���� run_MFP_update.py...
python "run_MFP_update.py"
if errorlevel 1 goto error

:: -----------------------------
:: [5/7] ���� add_ver.py�]����y�{�^
:: -----------------------------
echo [5/7] ���� add_ver.py�]�������i��ʿ�J�A10 ���۰ʶ�J�^...
call python "add_ver.py"

for /f %%a in ('python -c "import openpyxl;wb=openpyxl.load_workbook('data.xlsx');print(wb['����']['G1'].value)"') do set vernum=%%a
echo �ϥΪ�����: %vernum%


:: -----------------------------
:: [6/7] ���� data_updw.py
:: -----------------------------
echo [6/7] ���� data_updw.py...
python "data_updw.py"
if errorlevel 1 goto error

:: -----------------------------
:: [7/7] ���� save_excel.exe
:: -----------------------------
echo [7/7] �Ұ� save_excel.exe...
start /wait "" "save_excel.exe"
if errorlevel 1 goto error

:: -----------------------------
:: ���� Git �ާ@
:: -----------------------------
goto git_operation

:: -----------------------------
:: Git-only �Ҧ�
:: -----------------------------
:git_only
:: �۰ʥͦ�������
for /f %%a in ('python "add_ver.py" --auto') do set vernum=%%a
echo �ϥΪ�����: %vernum%

echo [1/3] �g�J�������� Excel�]�w�� add_ver.py �g�J�^
echo [2/3] �x�s Excel�]save_excel.exe�^...
start /wait "" "save_excel.exe"
if errorlevel 1 goto error

echo [3/3] �i�� Git �ާ@...
cd /d D:\flask
git pull
if errorlevel 1 goto error
git add -A
git commit -m "Auto commit - %vernum%"
if errorlevel 1 goto error
git push
if errorlevel 1 goto error

echo �Ҧ��{�Ǥw�����I
pause
exit /b

:: -----------------------------
:: Git �ާ@�l�y�{
:: -----------------------------
:git_operation
echo [Git] �i�� Git �ާ@...
cd /d D:\flask
git pull
if errorlevel 1 goto error
git add -A
git commit -m "Auto commit - %vernum%"
if errorlevel 1 goto error
git push
if errorlevel 1 goto error

echo �Ҧ��{�Ǥw�����I
pause
exit /b

:: -----------------------------
:: ���~�B�z
:: -----------------------------
:error
echo ���~�G�Y�ӨB�J���楢�ѡA�y�{����C
pause
exit /b
