@echo off
rem =============================================================================
rem  SPY daytrade daily review -- ONE-CLICK runner (works with NO AI at all)
rem  Place:  same folder as spy_daytrade_engine.py  (paths are all relative)
rem  Usage:  double-click, or drag a statement .csv onto this file
rem
rem  Route A (operation manual section 1-A), fully automatic:
rem     1 environment check  2 stop running app  3 pick today's CSV
rem     4 engine dry-run + green gate            5 writer --commit
rem     6 start review web app + open browser
rem  If no CSV is found, a menu offers: app only / route B (cs_from_trades).
rem
rem  ---------------------------------------------------------------------------
rem  ENCODING POLICY -- READ BEFORE EDITING
rem  This file MUST stay PURE ASCII (7-bit). Do not paste Chinese in here.
rem  Reason: cmd.exe parses a .bat with the *console* code page and seeks by
rem  byte offset on goto/call. Multi-byte text inside a .bat (UTF-8 or Big5)
rem  is the classic cause of "batch stops half way" / mojibake / a Big5 trail
rem  byte 0x7C being read as a pipe. So all Chinese UI text is passed to
rem  :say as \uXXXX escapes and decoded+printed by Python (UTF-8), which is
rem  encoding-safe. PYTHONIOENCODING=utf-8 also forces every child Python to
rem  emit UTF-8, so the captured logs are always UTF-8 and never crash on emoji.
rem
rem  To add or change a message, generate the escape with:
rem     python -c "s=input();print(''.join('\\u%04x'%ord(c) if ord(c)>126 else c for c in s))"
rem  and paste the result into a call :say "..." line.
rem  ---------------------------------------------------------------------------
rem
rem  SAFETY: this runner never computes or edits any number itself. It only
rem  runs the same scripts the operation manual tells a human to run, and it
rem  refuses to continue unless the engine output is fully green. Any red flag
rem  stops it with nothing written (writer does its own backup/restore).
rem =============================================================================

setlocal EnableExtensions DisableDelayedExpansion

rem -- remember the console code page so we can put it back on exit ------------
set "OLDCP="
for /f "tokens=2 delims=:" %%a in ('chcp 2^>nul') do set "OLDCP=%%a"
if defined OLDCP set "OLDCP=%OLDCP: =%"
if defined OLDCP set "OLDCP=%OLDCP:.=%"
chcp 65001 >nul 2>&1

title SPY Daily Review - One Click
set "PYTHONIOENCODING=utf-8"
set "RC=0"

rem -- work in the folder this .bat lives in (pushd also handles UNC paths) ----
pushd "%~dp0" 2>nul || (
  echo [FATAL] cannot enter script folder: %~dp0
  pause
  exit /b 1
)

set "LOGDIR=%CD%\_logs"
if not exist "%LOGDIR%" md "%LOGDIR%" >nul 2>&1
if not exist "%LOGDIR%" set "LOGDIR=%TEMP%"
set "ENGLOG=%LOGDIR%\engine_last.txt"
set "WRLOG=%LOGDIR%\writer_last.txt"
set "CSLOG=%LOGDIR%\cs_from_trades_last.txt"

echo.
echo ==============================================================
echo    SPY Daily Review - One Click
echo ==============================================================
echo.
echo   [1/6] checking environment ...

rem ===========================================================================
rem  [1/6] environment  (plain ASCII until a working Python is found, because
rem        :say needs Python to render the Chinese text)
rem ===========================================================================
set "PY="
set "PYANY="
call :probe py -3
call :probe python
call :probe python3
for /f "delims=" %%d in ('dir /b /ad /o-n "%LOCALAPPDATA%\Programs\Python\Python3*" 2^>nul') do call :probe "%LOCALAPPDATA%\Programs\Python\%%d\python.exe"
for /f "delims=" %%d in ('dir /b /ad /o-n "%ProgramFiles%\Python3*" 2^>nul') do call :probe "%ProgramFiles%\%%d\python.exe"
for /f "delims=" %%d in ('dir /b /ad /o-n "C:\Python3*" 2^>nul') do call :probe "C:\%%d\python.exe"

if defined PY goto :py_ok
if defined PYANY goto :py_nodeps

rem -- no Python at all: cannot use :say here, print plain ASCII --------------
echo.
echo   [X] Python 3 not found on this PC.
echo       Install Python 3 and tick "Add python.exe to PATH" during setup,
echo       see the operation manual appendix B, then run this file again.
set "RC=1"
goto :finish

:py_nodeps
set "PY=%PYANY%"
call :say "\U0001f534 \u7f3a\u5c11\u5fc5\u8981\u7684 Python \u5957\u4ef6\u3002\u8acb\u5728\u9019\u500b\u8996\u7a97\u57f7\u884c\u4e0b\u9762\u9019\u4e00\u884c\uff0c\u5b8c\u6210\u5f8c\u518d\u8dd1\u4e00\u6b21\u672c\u7a0b\u5f0f\uff1a"
echo.
echo        pip install flask pandas numpy openpyxl yfinance
echo.
set "RC=1"
goto :finish

:py_ok
call :say "   SPY \u7576\u6c96\u6bcf\u65e5\u8907\u76e4 \u2014 \u4e00\u9375\u57f7\u884c\uff08\u4e0d\u9700\u8981 AI\uff09"
call :say "   \u7167\u300a\u64cd\u4f5c\u624b\u518a\u300b\u8def\u7dda A\uff1a\u5f15\u64ce\u6aa2\u67e5 \u2192 \u5beb\u5165\u5169\u672c\u5e33 \u2192 \u958b\u8907\u76e4\u7db2\u9801"
echo.
echo   Python : %PY%
echo   Folder : %CD%

rem -- required files ---------------------------------------------------------
call :files >nul 2>&1
if not errorlevel 1 goto :files_ok
call :say "\U0001f534 \u7f3a\u5c11\u5fc5\u8981\u6a94\u6848\uff0c\u7121\u6cd5\u57f7\u884c\u3002\u8acb\u78ba\u8a8d\u6574\u500b TradeReview \u8cc7\u6599\u593e\u90fd\u8907\u88fd\u904e\u4f86\u4e86\u3002\u7f3a\u5c11\uff1a"
call :files
set "RC=1"
goto :finish

:files_ok
rem -- Excel lock files (~$xxx.xlsx) ------------------------------------------
if not exist "~$*.xls*" goto :lock_ok
call :say "\U0001f534 \u5075\u6e2c\u5230 Excel \u9396\u5b9a\u6a94\uff08~$ \u958b\u982d\uff09\uff0c\u8868\u793a\u5e33\u672c\u6b63\u958b\u5728 Excel \u88e1\u3002\n   \u8acb\u5148\u628a\u5169\u672c Excel \u90fd\u95dc\u6389\uff0c\u518d\u91cd\u65b0\u57f7\u884c\u672c\u7a0b\u5f0f\u3002"
dir /b "~$*.xls*" 2>nul
set "RC=1"
goto :finish

:lock_ok
call :say "  \u74b0\u5883\u6aa2\u67e5\u901a\u904e"

rem ===========================================================================
rem  [2/6] stop a running review app (frees port 5500 and the xlsx files)
rem ===========================================================================
echo.
call :hdr "[2/6] \u95dc\u9589\u6b63\u5728\u57f7\u884c\u7684\u8907\u76e4\u7db2\u9801"
call :killapp

rem ===========================================================================
rem  [3/6] choose today's statement CSV
rem ===========================================================================
echo.
call :hdr "[3/6] \u9078\u64c7\u4eca\u5929\u7684\u5c0d\u5e33\u55ae CSV"

set "CSVFILE="
if "%~1"=="" goto :scaninbox
if exist "%~1" set "CSVFILE=%~1"
if defined CSVFILE goto :have_csv

:scaninbox
for /f "delims=" %%f in ('dir /b /a-d /o-d "_inbox\*.csv" 2^>nul') do call :pickcsv "%%f"
if defined CSVFILE goto :have_csv

rem -- nothing to process -> menu ---------------------------------------------
call :say "\u5728 _inbox \u8cc7\u6599\u593e\u88e1\u627e\u4e0d\u5230\u4eca\u5929\u7684\u5c0d\u5e33\u55ae CSV\u3002\n\uff08\u8acb\u5148\u628a\u5c0d\u5e33\u55ae\u5b58\u6210 _inbox\\YYYY-MM-DD.csv\uff0c\u898b\u300a\u64cd\u4f5c\u624b\u518a\u300b\u7b2c 2 \u7bc0\u6b65\u9a5f 1\uff09"
echo.
call :say "\u8acb\u9078\u64c7\u8981\u505a\u4ec0\u9ebc\uff1a\n   1 = \u53ea\u555f\u52d5\u8907\u76e4\u7db2\u9801\uff08\u4e0d\u5beb\u5e33\uff09\n   2 = \u624b\u52d5\u8dd1\u7b2c\u4e8c\u6bb5\uff1atrades_all \u2192 CS \u5e33\uff08\u8981\u81ea\u5df1\u8f38\u5165\u671f\u521d/\u6536\u76e4\u9918\u984d\uff09\n   3 = \u96e2\u958b"
choice /c 123 /n /m "[1/2/3] "
if errorlevel 3 goto :aborted
if errorlevel 2 goto :routeb
call :say "\u53ea\u555f\u52d5\u8907\u76e4\u7db2\u9801\uff0c\u4e0d\u5beb\u4efb\u4f55\u5e33\u3002"
goto :startapp

:have_csv
call :say "  \u5df2\u9078\u64c7\u5c0d\u5e33\u55ae\uff1a"
echo        "%CSVFILE%"
echo.
call :say "  \u8981\u7528\u9019\u500b\u6a94\u6848\u8dd1\u55ce\uff1f20 \u79d2\u5f8c\u81ea\u52d5\u7e7c\u7e8c\uff1b\u6309 N \u53d6\u6d88"
where choice >nul 2>&1 || goto :run_engine
choice /c YN /n /t 20 /d Y /m "[Y/N] "
if errorlevel 2 goto :aborted

rem ===========================================================================
rem  [4/6] engine dry-run  (writes nothing; must come out fully green)
rem ===========================================================================
:run_engine
echo.
call :hdr "[4/6] \u5f15\u64ce\u4e7e\u8dd1\u6aa2\u67e5\uff08\u7d14\u6aa2\u67e5\uff0c\u4e0d\u6703\u52d5\u5230\u4efb\u4f55\u6a94\u6848\uff09"
echo.
%PY% spy_daytrade_engine.py "%CSVFILE%" > "%ENGLOG%" 2>&1
set "ERR=%errorlevel%"
type "%ENGLOG%"
if not "%ERR%"=="0" goto :eng_crash
call :gate "%ENGLOG%"
if errorlevel 1 goto :eng_red
echo.
call :say "  \u5168\u7da0\uff1a\u03a3col1 = \u03a3pnl\u3001\u7121\u7559\u5009\u3001\u7121\u8b66\u544a"
goto :run_writer

:eng_crash
echo.
call :say "\U0001f534 \u5f15\u64ce\u57f7\u884c\u4e2d\u65b7\uff08\u7a0b\u5f0f\u932f\u8aa4\u6216 CSV \u683c\u5f0f\u4e0d\u5c0d\uff09\u3002\n   \u5e38\u898b\u539f\u56e0\uff1aCSV \u6c92\u6284\u5230\u5c0d\u5e33\u55ae\u90a3\u6bb5\u3001\u6216\u6a94\u6848\u4e0d\u662f UTF-8\u3002\n   \u8acb\u628a\u4e0a\u9762\u8a0a\u606f\u622a\u5716\u4ea4\u7d66\u8ca0\u8cac\u4eba\u3002"
call :say "  \u5b8c\u6574\u8f38\u51fa\u5df2\u5b58\u5230 _logs \u8cc7\u6599\u593e\uff08engine_last.txt / writer_last.txt\uff09"
set "RC=2"
goto :finish

:eng_red
echo.
call :say "\U0001f534 \u6aa2\u67e5\u6c92\u904e\uff0c\u6d41\u7a0b\u5df2\u505c\u6b62\uff0c\u5169\u672c\u5e33\u5b8c\u5168\u6c92\u6709\u88ab\u4fee\u6539\u3002\n\n   \u9019\u662f\u7cfb\u7d71\u6b63\u5e38\u7684\u4fdd\u8b77\uff0c\u4e0d\u662f\u58de\u6389\u3002\u4e0a\u9762\u8a0a\u606f\u5c31\u662f\u539f\u56e0\uff0c\u5e38\u898b\u7684\u662f\uff1a\n     - BALANCE \u5c0d\u4e0d\u4e0a \u2192 \u67d0\u7b46\u6578\u5b57\u6284\u932f\u6216\u6f0f\u55ae\n     - \u7559\u5009 / \u8de8\u65e5\u5e73\u5009 \u2192 \u7279\u4f8b\u65e5\uff0c\u8981\u4eba\u5de5\u5224\u65b7\n     - \u9700\u4eba\u5de5\u78ba\u8a8d \u2192 \u5f15\u64ce\u4e0d\u6562\u81ea\u52d5\u8655\u7406\n\n   \u8acb\u628a\u4e0a\u9762\u6574\u6bb5\u8a0a\u606f\u622a\u5716\uff0c\u4ea4\u7d66\u8ca0\u8cac\u4eba\uff08\u898b\u300a\u64cd\u4f5c\u624b\u518a\u300b\u7b2c 4 \u7bc0\uff09\u3002\n   \u5343\u842c\u4e0d\u8981\u81ea\u5df1\u624b\u52d5\u6539\u6578\u5b57\u786c\u6e4a \u2014\u2014 \u9019\u662f\u4ee5\u524d\u51fa\u5927\u932f\u7684\u539f\u56e0\u3002"
call :say "  \u5b8c\u6574\u8f38\u51fa\u5df2\u5b58\u5230 _logs \u8cc7\u6599\u593e\uff08engine_last.txt / writer_last.txt\uff09"
set "RC=2"
goto :finish

rem ===========================================================================
rem  [5/6] writer --commit  (backup -> write -> full post-verify -> restore)
rem ===========================================================================
:run_writer
echo.
call :hdr "[5/6] \u5beb\u5165\u5169\u672c\u5e33\uff08\u81ea\u52d5\u5099\u4efd \u2192 \u5beb\u5165 \u2192 \u5beb\u5f8c\u5168\u6838\u9a57\uff1b\u5931\u6557\u81ea\u52d5\u9084\u539f\uff09"
echo.
%PY% spy_daytrade_writer.py "%CSVFILE%" --commit > "%WRLOG%" 2>&1
set "ERR=%errorlevel%"
type "%WRLOG%"
if not "%ERR%"=="0" goto :wr_fail
call :wgate "%WRLOG%"
if errorlevel 1 goto :wr_fail
echo.
call :say "  \u5beb\u5165\u5b8c\u6210\u4e26\u901a\u904e\u6838\u9a57\uff08\u82e5\u986f\u793a \u51aa\u7b49\u8df3\u904e \u4ee3\u8868\u4eca\u5929\u5148\u524d\u5df2\u5beb\u904e\uff0c\u6b63\u5e38\uff09"
goto :startapp

:wr_fail
echo.
call :say "\U0001f534 \u5beb\u5165\u6c92\u6709\u5b8c\u6210\u3002\u4e0a\u9762\u8a0a\u606f\u5c31\u662f\u539f\u56e0\uff1b\u5e33\u672c\u5df2\u81ea\u52d5\u9084\u539f\u6210\u5beb\u5165\u524d\u7684\u72c0\u614b\u3002\n   \u8acb\u622a\u5716\u4ea4\u7d66\u8ca0\u8cac\u4eba\uff0c\u4e0d\u8981\u81ea\u5df1\u6539\u3002"
call :say "  \u5b8c\u6574\u8f38\u51fa\u5df2\u5b58\u5230 _logs \u8cc7\u6599\u593e\uff08engine_last.txt / writer_last.txt\uff09"
set "RC=3"
goto :finish

rem ===========================================================================
rem  route B: trades_all.xlsx -> CS book only (pure program, never needs AI)
rem ===========================================================================
:routeb
echo.
call :say "\u8acb\u8f38\u5165\u4ea4\u6613\u65e5\uff0c\u683c\u5f0f 2026-07-24"
set "BD="
set /p "BD=> "
call :say "\u8acb\u8f38\u5165\u671f\u521d\u9918\u984d\uff08\u5c0d\u5e33\u55ae Cash balance at the start of business day \u90a3\u4e00\u884c\uff09"
set "BO="
set /p "BO=> "
call :say "\u8acb\u8f38\u5165\u6536\u76e4\u9918\u984d\uff08\u7576\u5929\u6700\u5f8c\u4e00\u7b46\u7684 BALANCE\uff09"
set "BC="
set /p "BC=> "
call :valid "%BD%" "%BO%" "%BC%"
if not errorlevel 1 goto :routeb_dry
call :say "\u8f38\u5165\u4e0d\u5b8c\u6574\uff0c\u5df2\u53d6\u6d88\u3002"
set "RC=1"
goto :finish

:routeb_dry
echo.
call :say "\u5148\u8a66\u7b97\uff08\u4e0d\u6703\u52d5\u5230\u6a94\u6848\uff09"
echo.
%PY% cs_from_trades.py --date "%BD%" --open-bal "%BO%" --close-bal "%BC%" > "%CSLOG%" 2>&1
set "ERR=%errorlevel%"
type "%CSLOG%"
if "%ERR%"=="0" goto :routeb_conf
echo.
call :say "\U0001f534 \u8a66\u7b97\u6c92\u904e\uff0c\u6c92\u6709\u5beb\u5165\u4efb\u4f55\u6771\u897f\u3002\u8acb\u770b\u4e0a\u9762\u539f\u56e0\uff0c\u4ea4\u8ca0\u8cac\u4eba\u8655\u7406\u3002"
set "RC=2"
goto :finish

:routeb_conf
echo.
call :say "  \u8a66\u7b97\u901a\u904e\u3002\u8981\u771f\u7684\u5beb\u5165 CS \u5e33\u55ce\uff1f20 \u79d2\u5f8c\u81ea\u52d5\u5beb\u5165\uff1b\u6309 N \u53d6\u6d88"
where choice >nul 2>&1 || goto :routeb_commit
choice /c YN /n /t 20 /d Y /m "[Y/N] "
if errorlevel 2 goto :aborted

:routeb_commit
echo.
%PY% cs_from_trades.py --date "%BD%" --open-bal "%BO%" --close-bal "%BC%" --commit > "%CSLOG%" 2>&1
set "ERR=%errorlevel%"
type "%CSLOG%"
if not "%ERR%"=="0" goto :routeb_fail
call :wgate "%CSLOG%"
if errorlevel 1 goto :routeb_fail
echo.
call :say "  CS \u5e33\u5beb\u5165\u5b8c\u6210\u4e26\u901a\u904e\u6838\u9a57"
goto :startapp

:routeb_fail
echo.
call :say "\U0001f534 CS \u5e33\u5beb\u5165\u5931\u6557\uff0c\u5df2\u81ea\u52d5\u9084\u539f\u5099\u4efd\u3002\u8acb\u622a\u5716\u4ea4\u8ca0\u8cac\u4eba\u3002"
set "RC=3"
goto :finish

rem ===========================================================================
rem  [6/6] start the review web app and open the browser
rem ===========================================================================
:startapp
echo.
call :hdr "[6/6] \u555f\u52d5\u8907\u76e4\u7db2\u9801"
start "TradeReview App" %PY% trade_review_app.py
call :say "  \u6b63\u5728\u555f\u52d5\uff0c\u7b2c\u4e00\u6b21\u8b80 K \u7dda\u8cc7\u6599\u53ef\u80fd\u8981 1\uff5e2 \u5206\u9418\uff0c\u8acb\u8010\u5fc3\u7b49"
call :waitport 5500 60
if errorlevel 1 goto :app_timeout
echo.
call :say "  \u7db2\u9801\u5df2\u5c31\u7dd2\uff1ahttp://localhost:5500"
start "" "http://localhost:5500/"
echo.
call :say "\u5168\u90e8\u5b8c\u6210\u3002\n   - \u700f\u89bd\u5668\u5df2\u81ea\u52d5\u6253\u958b http://localhost:5500\n   - \u770b\u5b8c\u8907\u76e4\u5f8c\uff0c\u628a TradeReview App \u90a3\u500b\u8996\u7a97\u95dc\u6389\u5373\u53ef\n   - \u5e33\u672c\u8207\u81ea\u52d5\u5099\u4efd\u90fd\u5728\u672c\u8cc7\u6599\u593e\u5167"
goto :finish

:app_timeout
echo.
call :say "\u26a0 \u7b49\u592a\u4e45\uff0c\u7db2\u9801\u9084\u6c92\u8d77\u4f86\u3002\n   \u8acb\u770b\u525b\u525b\u53e6\u5916\u8df3\u51fa\u7684 TradeReview App \u8996\u7a97\u88e1\u7684\u8a0a\u606f\uff08\u53ef\u80fd\u662f\u5957\u4ef6\u7f3a\u5c11\u6216\u8cc7\u6599\u6a94\u554f\u984c\uff09\u3002"
echo.
echo        %PY% trade_review_app.py
echo.
set "RC=4"
goto :finish

:aborted
echo.
call :say "\u5df2\u53d6\u6d88\uff0c\u6c92\u6709\u52d5\u5230\u4efb\u4f55\u6a94\u6848\u3002"
set "RC=1"
goto :finish

rem ===========================================================================
rem  helpers
rem ===========================================================================

rem -- print a \uXXXX-escaped string as real UTF-8 text ------------------------
:say
%PY% -c "import sys;print(sys.argv[1].encode('ascii','backslashreplace').decode('unicode_escape'))" "%~1"
goto :eof

rem -- section header ---------------------------------------------------------
:hdr
call :say "%~1"
goto :eof

rem -- probe a python launcher: needs python 3 AND the required packages -------
:probe
if defined PY goto :eof
%* -c "import sys;sys.exit(0 if sys.version_info[0]==3 else 1)" >nul 2>&1
if errorlevel 1 goto :eof
if not defined PYANY set "PYANY=%*"
%* -c "import flask,pandas,numpy,openpyxl" >nul 2>&1
if errorlevel 1 goto :eof
set "PY=%*"
goto :eof

rem -- list missing required files (exit 1 if any) ----------------------------
:files
%PY% -c "import os,sys;m=[p for p in ('spy_daytrade_engine.py','spy_daytrade_writer.py','cs_from_trades.py','trade_review_app.py','offset_state.json','trades_all.xlsx','CS\u4ea4\u6613\u7d00\u9304.xlsx') if not os.path.exists(p)];print(chr(10).join('        - '+x for x in m));sys.exit(1 if m else 0)"
goto :eof

rem -- pick newest inbox csv, skipping the bundled example/sample -------------
:pickcsv
if defined CSVFILE goto :eof
echo "%~1"| findstr /i /c:"example" >nul && goto :eof
echo "%~1"| findstr /i /c:"sample" >nul && goto :eof
set "CSVFILE=_inbox\%~1"
goto :eof

rem -- ENGINE GREEN GATE ------------------------------------------------------
rem  exit 0 only when: no warning marker anywhere in the log, at least one
rem  session summary line exists, and every session has Sigma-col1 == Sigma-pnl
rem  with no overnight position. Anything else -> exit 1 -> nothing is written.
:gate
%PY% -c "import sys,re;t=open(sys.argv[1],encoding='utf-8',errors='replace').read();bad=[k for k in ('\u26a0','\U0001f534','\u274c','BALANCE \u5c0d\u4e0d\u4e0a','\u7559\u5009=True','\u8b66\u544a/\u9700\u4eba\u5de5\u78ba\u8a8d','Traceback') if k in t];m=re.findall(r'\u03a3col1=(-?[0-9.]+)\s+\u03a3pnl=(-?[0-9.]+)\s+\u7559\u5009=(\S+)',t);ok=(not bad) and bool(m) and all(round(float(a)-float(b),2)==0.0 and c=='False' for a,b,c in m);sys.exit(0 if ok else 1)" "%~1"
goto :eof

rem -- WRITER / CS GATE -------------------------------------------------------
rem  exit 0 only when a success marker is present and no red/failure marker is.
:wgate
%PY% -c "import sys;t=open(sys.argv[1],encoding='utf-8',errors='replace').read();bad=any(k in t for k in ('\U0001f534','\u274c','\u26a0','Traceback'));good=any(k in t for k in ('\u2705','\u51aa\u7b49\u8df3\u904e','\u7121\u65b0\u8cc7\u6599\u53ef\u5beb'));sys.exit(0 if (good and not bad) else 1)" "%~1"
goto :eof

rem -- validate route B inputs ------------------------------------------------
:valid
%PY% -c "import sys,re;d,o,c=sys.argv[1],sys.argv[2],sys.argv[3];ok=bool(re.match(r'^[0-9]{4}-[0-9]{2}-[0-9]{2}$',d)) and (float(o)==float(o)) and (float(c)==float(c));sys.exit(0 if ok else 1)" "%~1" "%~2" "%~3" 2>nul
goto :eof

rem -- stop the review app if it is listening on 5500 (python processes only) -
:killapp
set "KILLED="
set "BUSY="
for /f "tokens=5" %%p in ('netstat -ano -p TCP 2^>nul ^| findstr /c:":5500 " ^| findstr /c:"LISTENING"') do call :killone %%p
if defined KILLED goto :killdone
if defined BUSY goto :eof
call :say "  \u6c92\u6709\u6b63\u5728\u57f7\u884c\u7684 App\uff0c\u7565\u904e"
goto :eof

:killdone
%PY% -c "import time;time.sleep(1.5)"
call :say "  \u5df2\u95dc\u9589\u820a\u7684 App"
goto :eof

rem -- kill one listener pid, python processes only. Port 5500 can have more
rem -- than one listener (an app left over from an old folder, for example),
rem -- so every pid found must be handled, not just the first one.
:killone
tasklist /fi "PID eq %~1" /nh 2>nul | findstr /i /c:"python" >nul
if errorlevel 1 goto :killbusy
taskkill /f /pid %~1 >nul 2>&1
set "KILLED=1"
goto :eof
:killbusy
if defined BUSY goto :eof
set "BUSY=1"
call :say "\u26a0 \u9023\u63a5\u57e0 5500 \u88ab\u975e Python \u7684\u7a0b\u5f0f\u5360\u7528\uff0c\u672c\u7a0b\u5f0f\u4e0d\u6703\u5f37\u5236\u95dc\u9589\u5b83\u3002\n   \u82e5\u7a0d\u5f8c\u7db2\u9801\u6253\u4e0d\u958b\uff0c\u8acb\u624b\u52d5\u8655\u7406\u8a72\u7a0b\u5f0f\u3002"
goto :eof

rem -- wait for a tcp port: %1 = port, %2 = tries (2s apart) -------------------
:waitport
%PY% -c "import sys;exec('import socket,time\nfor i in range(int(sys.argv[2])):\n    s=socket.socket()\n    s.settimeout(1)\n    r=s.connect_ex((chr(49)+chr(50)+chr(55)+chr(46)+chr(48)+chr(46)+chr(48)+chr(46)+chr(49),int(sys.argv[1])))\n    s.close()\n    if r==0:\n        sys.exit(0)\n    sys.stdout.write(chr(46))\n    sys.stdout.flush()\n    time.sleep(2)\nsys.exit(1)')" "%~1" "%~2"
goto :eof

rem ===========================================================================
:finish
echo.
echo ==============================================================
if defined PY call :say "\u6309\u4efb\u610f\u9375\u95dc\u9589\u9019\u500b\u8996\u7a97\u3002"
popd
if defined OLDCP chcp %OLDCP% >nul 2>&1
pause >nul
exit /b %RC%
