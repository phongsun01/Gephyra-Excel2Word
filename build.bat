@echo off
echo [Gephyra] Installing Dependencies...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [Error] Failed to install dependencies.
    pause
    exit /b %errorlevel%
)

echo.
echo [Gephyra] Building Executable...
python build_exe.py
if %errorlevel% neq 0 (
    echo [Error] Build failed.
    pause
    exit /b %errorlevel%
)

echo.
echo [Success] Build complete! Executable is in 'dist' folder.
pause
