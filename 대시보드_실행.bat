@chcp 65001 >nul
@echo off
cd /d "%~dp0"
title [MPS Dashboard Launcher]
cls
echo ==========================================================
echo   MPS Dashboard Server를 백그라운드에서 실행합니다.
echo   브라우저 창이 실행되면 이 콘솔 창은 자동으로 닫힙니다.
echo ==========================================================
echo.

taskkill /f /im mps_dashboard_app.exe >nul 2>&1
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :8890 ^| findstr LISTENING') do taskkill /f /pid %%a >nul 2>&1
ping 127.0.0.1 -n 2 >nul

:: 1. 내장 node.exe 존재 체크
if exist node.exe (
    goto RUN_LOCAL_NODE
)

:: 2. 전역 시스템 node 설치 체크
where node >nul 2>&1
if %errorlevel% equ 0 (
    goto RUN_SYSTEM_NODE
)

:: 3. node.exe 자동 다운로드 시도
echo [WARNING] 실행에 필요한 Node.js 환경이 발견되지 않았습니다.
echo [INFO] 인터넷에서 공식 Node.js 실행 파일(node.exe, 약 30MB)을 자동으로 다운로드합니다.
echo [INFO] 최초 1회만 다운로드하며, 네트워크 환경에 따라 5~20초 소요됩니다.
echo.
echo 다운로드 중... 잠시만 기다려 주세요.
echo.

powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $ProgressPreference = 'Continue'; Invoke-WebRequest -Uri 'https://nodejs.org/dist/v18.16.0/win-x64/node.exe' -OutFile 'node.exe'"

if exist node.exe (
    echo.
    echo [SUCCESS] Node.js 다운로드가 완료되었습니다!
    ping 127.0.0.1 -n 2 >nul
    goto RUN_LOCAL_NODE
)

:: 4. 다운로드 실패 시 mps_dashboard_app.exe 차선책 또는 에러 출력
if exist mps_dashboard_app.exe (
    echo [WARNING] 다운로드에 실패하여 기존 mps_dashboard_app.exe를 실행합니다.
    powershell -Command "Start-Process -FilePath 'mps_dashboard_app.exe' -WorkingDirectory '%~dp0' -WindowStyle Hidden"
    goto OPEN_BROWSER
) else (
    echo [ERROR] 인터넷 연결이 원활하지 않거나 다운로드에 실패했습니다.
    echo [ERROR] 프로그램을 시작할 수 없습니다. 인터넷 연결을 확인하거나 관리자에게 문의하세요.
    pause
    exit
)

:RUN_LOCAL_NODE
echo [INFO] 내장 Node.js [node.exe] 를 감지하여 백그라운드로 안전하게 실행합니다.
powershell -Command "Start-Process -FilePath '.\node.exe' -ArgumentList 'server.js' -WorkingDirectory '%~dp0' -WindowStyle Hidden"
goto OPEN_BROWSER

:RUN_SYSTEM_NODE
echo [INFO] 시스템 Node.js가 감지되어 백그라운드로 안전하게 실행합니다.
powershell -Command "Start-Process -FilePath 'node' -ArgumentList 'server.js' -WorkingDirectory '%~dp0' -WindowStyle Hidden"
goto OPEN_BROWSER

:OPEN_BROWSER
echo [INFO] 브라우저를 통해 대시보드를 엽니다...
ping 127.0.0.1 -n 3 >nul
start "" "http://localhost:8890"

echo [SUCCESS] 서버가 백그라운드에서 정상적으로 실행되었습니다.
ping 127.0.0.1 -n 2 >nul
exit
