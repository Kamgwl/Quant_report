@echo off
:: ============================================================
::  deploy.bat  —  Quant Dashboard Update + Git Push
::  Place this in: d:\file_dash\
::  Usage:  deploy.bat  (double-click or run in CMD)
:: ============================================================
setlocal

set "GIT=C:\Program Files\Git\bin\git.exe"
set "PYTHON=python"
set "DIR=%~dp0"

echo.
echo ════════════════════════════════════════════════════
echo   Quant Dashboard  ^|  Update + Deploy Pipeline
echo ════════════════════════════════════════════════════
echo.

:: ── STEP 1: Read Excel and update dashboard files ────────────
echo [STEP 1/3]  Reading Excel and updating dashboard files...
echo.
cd /d "%DIR%"
%PYTHON% update_dashboard.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo [ERROR] Python script failed. Aborting.
    pause
    exit /b 1
)

:: ── STEP 2: Stage changed files ─────────────────────────────
echo.
echo [STEP 2/3]  Staging changes in git...
echo.
"%GIT%" add index.html indx.html quant-dashboard/src/data.js quant-dashboard/src/Dashboard.jsx "Quant Strategy(2026-27).xlsx"

"%GIT%" status --short
echo.

:: Check if there's anything to commit
"%GIT%" diff --cached --quiet
if %ERRORLEVEL% equ 0 (
    echo [INFO] No changes detected — dashboard is already up to date.
    echo        Nothing to commit. GitHub Pages unchanged.
    echo.
    pause
    exit /b 0
)

:: ── STEP 3: Commit and Push ──────────────────────────────────
echo [STEP 3/3]  Committing and pushing to GitHub...
echo.

:: Auto-generate commit message with timestamp
for /f "tokens=1-6 delims=/:. " %%a in ("%DATE% %TIME%") do (
    set "DT=%%c-%%b-%%a %%d:%%e"
)
:: Safer: just use a fixed message with date
set "MSG=Update Q1 FY2026-27 dashboard from Excel [%DATE%]"

"%GIT%" commit -m "%MSG%"
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Commit failed.
    pause
    exit /b 1
)

echo.
"%GIT%" push origin main
if %ERRORLEVEL% neq 0 (
    echo.
    echo [ERROR] Push failed. You may need to authenticate.
    echo         If prompted, use your GitHub Personal Access Token as the password.
    pause
    exit /b 1
)

echo.
echo ════════════════════════════════════════════════════
echo   SUCCESS!  Changes pushed to GitHub.
echo   Live in ~1-2 min: https://kamgwl.github.io/Quant_report/
echo ════════════════════════════════════════════════════
echo.
pause
