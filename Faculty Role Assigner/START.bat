@echo off
title Faculty Role Assigner
cd /d "%~dp0"

:: ── Check Python is installed ─────────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo  ERROR: Python is not installed.
    echo.
    echo  Please download and install Python from:
    echo  https://www.python.org/downloads/
    echo.
    echo  During install, tick "Add Python to PATH".
    echo.
    pause
    exit /b 1
)

:: ── Install required packages silently on first run ──────────────────────────
if not exist ".deps_ok" (
    echo  Installing required packages ^(first run only^)...
    python -m pip install --quiet --upgrade pip
    python -m pip install --quiet flask pandas openpyxl
    if errorlevel 1 (
        echo.
        echo  ERROR: Failed to install packages.
        echo  Please check your internet connection and try again.
        echo.
        pause
        exit /b 1
    )
    echo. > .deps_ok
    echo  Done. Starting app...
)

:: ── Launch the app ────────────────────────────────────────────────────────────
echo.
echo  Faculty Role Assigner is starting...
echo  Your browser will open automatically.
echo  Close this window to stop the app.
echo.
python launcher.py

pause
