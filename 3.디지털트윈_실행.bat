@echo off
chcp 65001 >nul
title 남산 MC 공장 샵플로어 디지털 트윈 실행기
cls

echo ================================================================
echo    [남산 MC 공장 샵플로어 게임형 디지털 트윈 시뮬레이터]
echo ================================================================
echo.
echo  * 서버 연결 확인 중...
echo.

powershell -Command "try { $r = Invoke-WebRequest -Uri 'http://localhost:8890/twin' -TimeoutSec 1; exit 0 } catch { exit 1 }" >nul 2>&1

if %ERRORLEVEL% equ 0 (
    echo  * 통합 서버가 이미 실행 중입니다.
    echo  * 브라우저에서 디지털 트윈 화면을 실행합니다...
    start http://localhost:8890/twin
) else (
    echo  * 통합 서버를 배경에서 시작합니다...
    start /b "" node server.js >nul 2>&1
    timeout /t 2 /nobreak >nul
    echo  * 브라우저에서 디지털 트윈 화면을 실행합니다...
    start http://localhost:8890/twin
)

echo.
echo ================================================================
echo  * 디지털 트윈이 브라우저에서 정상 실행되었습니다.
echo  * 창을 닫으셔도 백그라운드에서 계속 유지됩니다.
echo ================================================================
timeout /t 3 /nobreak >nul
exit
