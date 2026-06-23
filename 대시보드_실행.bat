@echo off
title [MPS 생산계획 대시보드 서버 - 실행 중]
cls
echo ==========================================================
echo   MPS 생산계획 대시보드 서버가 실행 중입니다.
echo   이 창을 닫으면 대시보드가 종료됩니다.
echo   사용 완료 후 브라우저의 [서버 종료] 버튼을 누르면
echo   이 창도 자동으로 함께 닫힙니다.
echo ==========================================================
echo.
taskkill /f /im mps_dashboard.exe >nul 2>&1
taskkill /f /im node.exe >nul 2>&1
timeout /t 1 /nobreak >nul
mps_dashboard.exe
exit
