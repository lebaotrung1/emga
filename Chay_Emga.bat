@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>&1
if errorlevel 1 (
 echo Can cai Python 3.12 tro len tu python.org, sau do chay lai.
 pause
 exit /b 1
)
py -3 bootstrap.py
if errorlevel 1 goto :failed
exit /b 0
:failed
echo.
echo Chua khoi dong duoc. Xem thong bao cu the phia tren.
echo Neu pip/thu vien bi hong, dong cua so Emga cu va chay Sua_moi_truong.bat.
pause
exit /b 1
