@echo off
REM ============================================================
REM  DEVELOPMENT sandbox - port 8601
REM  Scheduler is disabled; safe to restart while users are on prod.
REM ============================================================
set PY=C:\Users\ngocluup\AppData\Local\miniforge3\envs\ngocluup\python.exe
cd /d "%~dp0"

set RM_ENV=dev
set RM_PORT=8601
REM Share the data folder with production so dev does not re-download EIMS/MARS.
set RM_DATA_DIR=%~dp0data
set RM_CONFIG=%~dp0config.local.json

echo ============================================================
echo   REJECT MANAGEMENT  -  DEV SANDBOX
echo       http://localhost:8601
echo   Production for users stays on port 8600.
echo   Press Ctrl+C to stop.
echo ============================================================
"%PY%" web\app.py
pause
