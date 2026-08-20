@echo off
chcp 65001 >nul
title 남산 MC 공장 조립 지번(Bay) 수기관리 현황판 실행기
cls

echo ================================================================
echo    [남산 MC 공장 조립 지번(Bay) 수기관리 현황판 (Port 8895)]
echo ================================================================
echo.
echo  * 전용 서버(Port 8895) 연결 확인 중...
echo.

powershell -Command "try { $r = Invoke-WebRequest -Uri 'http://localhost:8895/api/shopfloor' -TimeoutSec 1; exit 0 } catch { exit 1 }" >nul 2>&1

if %ERRORLEVEL% equ 0 (
    echo  * 전용 관리 서버가 이미 실행 중입니다.
    echo  * 브라우저에서 지번 수기관리 화면을 엽니다...
    start http://localhost:8895
) else (
    echo  * 전용 관리 서버(Port 8895)를 시작합니다...
    start /b "" node shopfloor_server.js >nul 2>&1
    timeout /t 2 /nobreak >nul
    echo  * 브라우저에서 지번 수기관리 화면을 엽니다...
    start http://localhost:8895
)

echo.
echo ================================================================
echo  * 지번 수기관리 현황판이 브라우저에서 정상 실행되었습니다.
echo  * 접속 주소: http://localhost:8895
echo ================================================================
timeout /t 3 /nobreak >nul
exit
