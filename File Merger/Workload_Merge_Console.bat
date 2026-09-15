@echo off
setlocal EnableExtensions
title Workload Merge Console

REM Always run from the folder containing this BAT file.
cd /d "%~dp0"

set "PORT=8787"
set "URL=http://127.0.0.1:%PORT%"

echo.
echo ============================================================
echo                 WORKLOAD MERGE CONSOLE
echo ============================================================
echo.

REM Check Node.js
where node >nul 2>&1
if errorlevel 1 (
    echo ERROR: Node.js is not installed or is not in PATH.
    echo.
    echo Install Node.js and run this file again.
    echo.
    pause
    exit /b 1
)

REM Check server.js
if not exist "%~dp0server.js" (
    echo ERROR: server.js was not found.
    echo.
    echo Expected:
    echo "%~dp0server.js"
    echo.
    pause
    exit /b 1
)

REM Check HTML
if not exist "%~dp0workload_merge_console_corrected.html" (
    echo ERROR: workload_merge_console_corrected.html was not found.
    echo.
    echo Expected:
    echo "%~dp0workload_merge_console_corrected.html"
    echo.
    pause
    exit /b 1
)

echo Starting server...
echo.

REM /D and quoted paths safely handle spaces in folder names.
start "Workload Merge Server" /D "%~dp0" cmd /k node server.js

echo Waiting for server to become ready...

REM Wait up to 30 seconds before opening the browser.
for /l %%I in (1,1,30) do (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "$r=$null; try {$r=Invoke-WebRequest -UseBasicParsing -Uri '%URL%/' -TimeoutSec 1} catch {}; if($r -and $r.StatusCode -eq 200){exit 0}else{exit 1}" >nul 2>&1
    if not errorlevel 1 goto SERVER_READY
    timeout /t 1 /nobreak >nul
)

echo.
echo ERROR: Server did not start within 30 seconds.
echo Check the Workload Merge Server window for the actual error.
echo.
pause
exit /b 1

:SERVER_READY
echo.
echo Server started successfully.
echo Opening %URL%
echo.

start "" "%URL%"

endlocal
exit /b 0
