@echo off
chcp 65001
echo.
echo ====================================================
echo        MPS 파일에서 대시보드용 CSV 자동 추출기 
echo ====================================================
echo.
echo 백그라운드 엑셀을 실행하여 폴더 안의 MPS 엑셀 파일을 찾고 있습니다...
echo (약 10초~30초 정도 소요될 수 있습니다)
echo.
powershell -ExecutionPolicy Bypass -NoProfile -File "Auto_Make_CSV.ps1"
exit
