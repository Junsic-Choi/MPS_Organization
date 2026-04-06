@echo off
setlocal
cd /d "%~dp0"

echo [1/2] MPS 대시보드 서버를 백그라운드에서 시작합니다...

:: 기존 서버가 있다면 종료 시도 (포트 8888 점유 중인 프로세스 종료)
powershell -Command "$p = Get-NetTCPConnection -LocalPort 8888 -ErrorAction SilentlyContinue; if ($p) { Stop-Process -Id $p.OwningProcess -Force }" 2>NUL

:: 실시간 로그 및 에러 초기화
if exist "server_start.log" del "server_start.log"

:: VBScript를 통해 서버를 투명하게 실행 (터미널 닫아도 유지됨)
if exist "run_dashboard_silent.vbs" (
    cscript //nologo "run_dashboard_silent.vbs"
    echo [2/2] 서버 기동 명령을 전달했습니다.
    echo.
    echo 대시보드 접속: http://localhost:8888
    echo.
    echo 이 창은 5초 후 자동으로 닫힙니다. (또는 아무 키나 누르세요)
    timeout /t 5 >nul
) else (
    echo [Error] run_dashboard_silent.vbs 파일을 찾을 수 없습니다.
    node server.js
)
exit


