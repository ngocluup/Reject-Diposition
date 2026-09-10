@echo off
REM Reject Management web app (Flask). Serves on port 8600 for the whole team.
set PY=C:\Users\ngocluup\AppData\Local\miniforge3\envs\ngocluup\python.exe
cd /d "%~dp0"
echo ============================================================
echo   REJECT MANAGEMENT  -  share ONE of these URLs (Intel net/VPN):
echo       http://%COMPUTERNAME%:8600
echo       http://ngocluup-iLIS09.ger.corp.intel.com:8600
echo       http://10.88.183.105:8600
echo   Keep this window OPEN. Press Ctrl+C to stop.
echo ============================================================
"%PY%" web\app.py 8600
pause
