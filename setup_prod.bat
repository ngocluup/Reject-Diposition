@echo off
REM ============================================================
REM  One-time setup: create the "prod" git worktree pinned to main.
REM  Production is served from prod\ so editing your working copy
REM  can never affect the users who are online.
REM ============================================================
cd /d "%~dp0"

if exist "prod\web\app.py" (
  echo The prod worktree already exists. Nothing to do.
  pause
  exit /b 0
)

echo Creating the prod worktree on branch "main"...
git worktree add prod main
if errorlevel 1 (
  echo [ERROR] Could not create the worktree.
  pause
  exit /b 1
)

echo.
echo Done. Now:
echo    run_prod.bat   ->  production for users   (port 8600)
echo    run_dev.bat    ->  your sandbox           (port 8601)
echo    deploy.bat     ->  publish dev to prod
pause
