@echo off
setlocal
cd /d "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\10_update_excel.ps1"
if errorlevel 1 (
  echo.
  echo [ERROR] Update failed. See logs\run.log
  exit /b 1
)

exit /b 0
