@echo off
cd /d "%~dp0"
echo 若首次使用请先运行: pip install -r requirements.txt
echo.
python app.py
if errorlevel 1 pause
