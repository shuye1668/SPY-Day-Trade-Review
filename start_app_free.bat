@echo off
REM ─────────────────────────────────────────────────────────
REM  Trade Review — FREE version launcher
REM  Double-click to start the local web app on http://localhost:5501
REM ─────────────────────────────────────────────────────────
cd /d "%~dp0"

REM Optional: catch up missing days first (silent if nothing to do)
python daily_fetch.py --catchup

REM Start the app
start "" http://localhost:5501
python trade_review_app_free.py
pause
