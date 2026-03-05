@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PS_EXE=%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS_EXE%" set "PS_EXE=powershell.exe"

"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\01_prepare_ps7_installer.ps1"
if errorlevel 1 (
  echo.
  echo [ERROR] PS7 インストーラーの取得に失敗しました。
  exit /b 1
)

echo.
echo [OK] PS7 インストーラーのダウンロードが完了しました。
echo [NEXT] インストールを開始します。
call "%~dp0run_install_ps7.bat"
if errorlevel 1 (
  exit /b 1
)

exit /b 0
