@echo off
echo Starting Gephyra Flet GUI...
call .venv\Scripts\activate.bat
python main.py --gui flet
pause
