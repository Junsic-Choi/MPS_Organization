@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo [1/2] MPS 대시보드를 준비 중입니다...
:: 기존 노드 프로세스 정리 (8890 포트 점유 방지)
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :8890 ^| findstr LISTENING') do taskkill /f /pid %%a >nul 2>&1

echo [2/2] 서버가 8890 포트에서 백그라운드 가공을 시작합니다.
echo.
start /b node server.js

echo 대시보드 주소: http://localhost:8890/dashboard.html
echo.
echo 이 창은 3초 후 자동으로 종료됩니다.
ping 127.0.0.1 -n 4 >nul
exit
exit


