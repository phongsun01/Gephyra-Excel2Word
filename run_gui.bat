@echo off
title Gephyra Launcher

:: Di chuyen vao thu muc du an
cd /d "%~dp0"

echo ==========================================
echo       GEPHYRA - BATCH GENERATOR
echo ==========================================
echo Phien ban: 1.0.0
echo.

:: 1. Kich hoat moi truong ao
if not exist ".venv\Scripts\activate.bat" goto no_venv
echo [INFO] Dang kich hoat moi truong ao...
call ".venv\Scripts\activate.bat"
goto check_libs

:no_venv
echo [!] Khong tim thay moi truong ao .venv.
goto check_libs

:check_libs
echo [INFO] Dang kiem tra thu vien...
python -c "import yaml, pandas, flet, docx" >nul 2>&1
if errorlevel 1 goto install_libs
goto menu

:install_libs
echo.
echo [!] Phat hien thieu thu vien quan trong (PyYAML, Pandas, Flet...).
set /p install_choice=Ban co muon tu dong cai dat cac thu vien nay khong? (y/n): 
if /i not "%install_choice%"=="y" goto menu
echo.
echo Dang cai dat thu vien...
python -m pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Khong the cai dat thu vien. Vui long kiem tra ket noi mang.
    pause
)
goto menu

:menu
echo.
echo [1] Chay giao dien Flet (Hien dai)
echo [2] Chay giao dien Tkinter (On dinh)
echo [3] Thoat
echo.

set choice=
set /p choice=Nhap so (1, 2, 3) roi go Enter: 

if "%choice%"=="1" goto run_flet
if "%choice%"=="2" goto run_tk
if "%choice%"=="3" exit /b
goto menu

:run_flet
echo Dang khoi dong Flet...
python main.py --gui flet
if errorlevel 1 pause
goto end

:run_tk
echo Dang khoi dong Tkinter...
python main.py --gui tk
if errorlevel 1 pause
goto end

:end
echo.
echo Chuong trinh ket thuc.
pause
