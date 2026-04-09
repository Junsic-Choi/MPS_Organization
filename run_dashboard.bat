@echo off
:: Powerful path tracking
pushd "%~dp0"
echo ------------------------------------------
echo MPS Project Directory: %CD%
echo ------------------------------------------

:: Check if server.js exists in the current folder
if not exist server.js (
    echo [ERROR] server.js NOT FOUND in this folder!
    echo Current files in this folder:
    dir /b
) else (
    echo [OK] server.js found. Starting Node...
    node server.js
)

if %ERRORLEVEL% neq 0 (
    echo.
    echo [CRITICAL] Node failed to start. 
    echo Please check if 'node' command is available in your terminal.
)

popd
pause
