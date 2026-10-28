@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>&1
if errorlevel 1 (
 echo Khong tim thay Python Launcher py. Gui thong bao nay de duoc ho tro.
 pause
 exit /b 1
)
py -3 bootstrap.py --diagnose
echo.
echo Xem tep emga-khoi-dong.log cung thu muc voi tep nay.
pause
