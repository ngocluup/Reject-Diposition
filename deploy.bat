@echo off
REM ============================================================
REM  Publish the current dev branch to production.
REM    1. commit anything outstanding on dev
REM    2. fast-forward main inside the prod worktree
REM    3. push both branches to GitHub
REM    4. tell you to restart run_prod.bat
REM ============================================================
setlocal EnableDelayedExpansion
cd /d "%~dp0"

if not exist "prod\web\app.py" (
  echo [ERROR] The prod worktree is missing. Run setup_prod.bat first.
  pause
  exit /b 1
)

for /f "delims=" %%b in ('git rev-parse --abbrev-ref HEAD') do set BRANCH=%%b
if not "!BRANCH!"=="dev" (
  echo [ERROR] You are on branch "!BRANCH!". Switch to dev first:  git checkout dev
  pause
  exit /b 1
)

REM --- 1. commit outstanding work -------------------------------------------
git diff --quiet
set DIRTY=!errorlevel!
git diff --cached --quiet
set /a DIRTY=!DIRTY!+!errorlevel!

if not "!DIRTY!"=="0" (
  echo.
  git status --short
  echo.
  set "MSG="
  set /p "MSG=Commit message for the changes above: "
  if "!MSG!"=="" (
    echo [ERROR] A commit message is required.
    pause
    exit /b 1
  )
  git add -A
  git commit -m "!MSG!"
  if errorlevel 1 (
    echo [ERROR] Commit failed.
    pause
    exit /b 1
  )
) else (
  echo Working tree is clean - nothing new to commit.
)

REM --- 2. fast-forward main inside the prod worktree ------------------------
echo.
echo Publishing dev to main and updating the prod worktree...
git -C prod merge --ff-only dev
if errorlevel 1 (
  echo [ERROR] main could not be fast-forwarded onto dev.
  echo         Resolve it manually:  git -C prod merge dev
  pause
  exit /b 1
)

REM --- 3. push both branches to GitHub --------------------------------------
echo.
git remote get-url origin >nul 2>&1
if errorlevel 1 (
  echo No "origin" remote configured - skipping the push.
) else (
  echo Pushing dev and main to GitHub...
  git push origin dev main
  if errorlevel 1 (
    echo [WARNING] The push failed ^(offline or VPN down?^).
    echo           Production was still updated. Retry later with:  git push origin dev main
  ) else (
    echo Push complete.
  )
)

echo.
echo ============================================================
echo   Deployed. Restart production to pick up the new code:
echo       close the run_prod.bat window, then start it again.
echo ============================================================
git -C prod log --oneline -1
pause
