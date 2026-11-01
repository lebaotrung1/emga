@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "NghiemThuStudio.exe" (
  echo Thieu NghiemThuStudio.exe. Hay giai nen day du goi chuong trinh.
  pause
  exit /b 1
)
start "" "NghiemThuStudio.exe"
