@echo off
REM ============================================================
REM   Reject Management  -  RUPS + EIMS web app
REM   Share the UI with your team. Backend runs on THIS machine.
REM   Keep this window OPEN while people are using the app.
REM ============================================================

set PYEXE=C:\Users\ngocluup\AppData\Local\miniforge3\envs\ngocluup\python.exe
cd /d "%~dp0"

echo ============================================================
echo   REJECT MANAGEMENT  -  RUPS + EIMS
echo.
echo   Share this URL with your team (same Intel network / VPN):
echo.
echo       http://NGOCLUUP-ILIS09:8501
echo   or  http://ngocluup-iLIS09.ger.corp.intel.com:8501
echo.
echo   Keep this window OPEN. Press Ctrl+C to stop.
echo ============================================================
echo.

"%PYEXE%" -m streamlit run streamlit_app.py

pause
