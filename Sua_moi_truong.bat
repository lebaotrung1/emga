@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>&1
if errorlevel 1 (
 echo Can cai Python 3.12 tro len tu python.org, sau do chay lai.
 pause
 exit /b 1
)
echo Dong cac cua so Emga dang chay truoc khi sua moi truong.
echo Moi truong thu vien cu se duoc giu lai bang ten khac.
echo Du lieu va tai khoan khong bi xoa. Windows dung thu muc EmgaRuntime ngan hon.
py -3 bootstrap.py --repair --check
if errorlevel 1 goto :failed
pause
exit /b 0
:failed
echo Chua sua xong. Xem loi goc phia tren va tep emga-khoi-dong.log.
pause
exit /b 1
