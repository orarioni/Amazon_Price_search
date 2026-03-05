@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PS_EXE=%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS_EXE%" set "PS_EXE=powershell"

"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File ".\scripts\01_prepare_ps7_installer.ps1"
if errorlevel 1 (
  echo.
  echo [ERROR] PS7 インストーラーの取得に失敗しました。
  exit /b 1
)

echo.
echo [OK] PS7 インストーラーの準備が完了しました。
exit /b 0
