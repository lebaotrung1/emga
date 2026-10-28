@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>&1
if errorlevel 1 (
 echo Can cai Python 3.12 tro len tu python.org, sau do chay lai.
 pause
 exit /b 1
)
if not exist ".venv\Scripts\python.exe" (
 py -3 -m venv .venv
 if errorlevel 1 goto :failed
)
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto :failed
set EMGA_HOST=0.0.0.0
if not defined EMGA_PORT set EMGA_PORT=8765
echo Mo trinh duyet tren may nay: http://127.0.0.1:%EMGA_PORT%
echo May LAN dung http://DIA_CHI_IP_MAY_CHU:%EMGA_PORT%
echo Chi mo Windows Firewall cho mang Private noi bo. Ctrl+C de dung.
.venv\Scripts\python.exe app.py
pause
exit /b 0
:failed
echo Khoi dong that bai. Kiem tra Python va ket noi Internet khi cai thu vien.
pause
exit /b 1
