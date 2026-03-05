@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PS7_EXE="
if exist "%ProgramFiles%\PowerShell\7\pwsh.exe" set "PS7_EXE=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not defined PS7_EXE if exist "%ProgramFiles(x86)%\PowerShell\7\pwsh.exe" set "PS7_EXE=%ProgramFiles(x86)%\PowerShell\7\pwsh.exe"
if not defined PS7_EXE set "PS7_EXE=pwsh"

"%PS7_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File ".\scripts\10_update_excel.ps1"
if errorlevel 1 (
  echo.
  echo [ERROR] 更新処理に失敗しました。メッセージを確認してください。
  echo [HINT] pwsh が見つからない場合は run_prepare_ps7_installer.bat を実行して PS7 インストーラーを取得してください。
  exit /b 1
)
exit /b 0
