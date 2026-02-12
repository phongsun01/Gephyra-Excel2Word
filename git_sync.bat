@echo off
setlocal enabledelayedexpansion

:: 1. Nhập message commit (nếu không có tham số)
set "COMMIT_MSG=%~1"
if "%COMMIT_MSG%"=="" (
    set /p "COMMIT_MSG=Nhap noi dung thay doi (Commit message): "
)

if "%COMMIT_MSG%"=="" (
    set "COMMIT_MSG=Update: %date% %time%"
)

echo.
echo [1/3] Dang staging cac thay doi (git add .)...
git add .

echo [2/3] Dang commit voi noi dung: "%COMMIT_MSG%"
git commit -m "%COMMIT_MSG%"

echo [3/3] Dang day code len server (git push)...
git push

echo.
set /p "CHOICE=Ban co muon tao Tag (version) cho ban nay khong? (y/n): "

if /i "%CHOICE%"=="y" (
    set /p "TAG_NAME=Nhap ten Tag (vi du: v1.0.0): "
    set /p "TAG_DESC=Nhap mo ta cho Tag: "
    
    if "!TAG_NAME!"=="" (
        echo [ERROR] Ten Tag khong duoc de trong. Da huy tao Tag.
    ) else (
        echo Dang tao Tag !TAG_NAME!...
        git tag -a !TAG_NAME! -m "!TAG_DESC!"
        echo Dang day Tag !TAG_NAME! len server...
        git push origin !TAG_NAME!
    )
)

echo.
echo === HOAN THANH ===
pause
