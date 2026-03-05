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
echo [NEXT] MSI を手動実行するか run_install_ps7.bat を実行してください。
exit /b 0
