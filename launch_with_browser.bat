@echo off
cd /d "%~dp0"
REM 约 2 秒后再打开默认浏览器（Flask 已监听）；机器较慢可改为 3
start "" cmd /c "timeout /t 2 /nobreak >nul & start http://127.0.0.1:5000/"
python app.py
if errorlevel 1 pause
