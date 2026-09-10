@echo off
REM ============================================================
REM  Publish the current dev branch to production.
REM    1. commit anything outstanding on dev
REM    2. fast-forward main to dev
REM    3. refresh the prod worktree
REM    4. tell you to restart run_prod.bat
REM ============================================================
setlocal
cd /d "%~dp0"

if not exist "prod\web\app.py" (
  echo [ERROR] The prod worktree is missing. Run setup_prod.bat first.
  pause
  exit /b 1
)

for /f "delims=" %%b in ('git rev-parse --abbrev-ref HEAD') do set BRANCH=%%b
if not "%BRANCH%"=="dev" (
  echo [ERROR] You are on branch "%BRANCH%". Switch to dev first:  git checkout dev
  pause
  exit /b 1
)

REM --- 1. commit outstanding work -------------------------------------------
git diff --quiet && git diff --cached --quiet
if errorlevel 1 (
  set /p MSG=Commit message for the pending changes: 
  git add -A
  git commit -m "%MSG%"
  if errorlevel 1 (
    echo [ERROR] Commit failed.
    pause
    exit /b 1
  )
)

REM --- 2. move main forward ---------------------------------------------------
echo Merging dev into main...
git fetch . dev:main
if errorlevel 1 (
  echo [ERROR] main could not be fast-forwarded. Resolve it manually:
  echo         git checkout main ^&^& git merge dev
  pause
  exit /b 1
)

REM --- 3. update the worktree -------------------------------------------------
echo Updating the prod worktree...
git -C prod checkout main -q
git -C prod reset --hard main -q
if errorlevel 1 (
  echo [ERROR] Could not update the prod worktree.
  pause
  exit /b 1
)

echo.
echo ============================================================
echo   Deployed. Restart production to pick up the new code:
echo       close the run_prod.bat window, then start it again.
echo ============================================================
git log --oneline -1 main
pause
