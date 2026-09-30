@echo off
cd /d "%~dp0"
start python trade_review_app.py
:wait
timeout /t 1 /nobreak >nul
curl -s http://localhost:5500/ >nul 2>&1
if errorlevel 1 goto wait
start "" "chrome.exe" "http://localhost:5500/"
