@echo off
REM ============================================================
REM  PRODUCTION - port 8600 - served from the "prod" git worktree
REM  Never edit files under prod\ by hand; use deploy.bat instead.
REM ============================================================
set PY=C:\Users\ngocluup\AppData\Local\miniforge3\envs\ngocluup\python.exe
cd /d "%~dp0"

if not exist "prod\web\app.py" (
  echo [ERROR] The prod worktree is missing. Run setup_prod.bat first.
  pause
  exit /b 1
)

set RM_ENV=prod
set RM_PORT=8600
REM Shared data + secrets live in the main checkout, not in the worktree.
set RM_DATA_DIR=%~dp0data
set RM_CONFIG=%~dp0config.local.json

echo ============================================================
echo   REJECT MANAGEMENT  -  PRODUCTION
echo       http://%COMPUTERNAME%:8600
echo       http://ngocluup-iLIS09.ger.corp.intel.com:8600
echo       http://10.88.183.105:8600
echo   Keep this window OPEN. Press Ctrl+C to stop.
echo ============================================================
"%PY%" prod\web\app.py
pause
