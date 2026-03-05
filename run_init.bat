@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PS7_EXE="
if exist "%ProgramFiles%\PowerShell\7\pwsh.exe" set "PS7_EXE=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not defined PS7_EXE if exist "%ProgramFiles(x86)%\PowerShell\7\pwsh.exe" set "PS7_EXE=%ProgramFiles(x86)%\PowerShell\7\pwsh.exe"
if not defined PS7_EXE set "PS7_EXE=pwsh"

set "PS_FILE=%~dp0scripts\00_init_secrets.ps1"
"%PS7_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%PS_FILE%"
set "EXIT_CODE=%ERRORLEVEL%"

if %EXIT_CODE% neq 0 (
  echo.
  echo [ERROR] 初期設定に失敗しました。メッセージを確認してください。 (ExitCode=%EXIT_CODE%)
  echo [HINT] pwsh が見つからない場合は run_prepare_ps7_installer.bat を実行して PS7 インストーラーを取得してください。
  exit /b %EXIT_CODE%
)

exit /b 0
