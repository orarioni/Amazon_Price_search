@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PS7_EXE="
if exist "%ProgramFiles%\PowerShell\7\pwsh.exe" set "PS7_EXE=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not defined PS7_EXE if exist "%ProgramFiles(x86)%\PowerShell\7\pwsh.exe" set "PS7_EXE=%ProgramFiles(x86)%\PowerShell\7\pwsh.exe"
if not defined PS7_EXE for %%I in (pwsh.exe) do set "PS7_EXE=%%~$PATH:I"

if not defined PS7_EXE (
  echo.
  echo [ERROR] PowerShell 7 (pwsh) が見つかりません。
  echo [HINT] 先に run_prepare_ps7_installer.bat を実行して PS7 をインストールしてください。
  exit /b 1
)

"%PS7_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\00_init_secrets.ps1"
set "EXIT_CODE=%ERRORLEVEL%"

if %EXIT_CODE% neq 0 (
  echo.
  echo [ERROR] 初期設定に失敗しました。メッセージを確認してください。 (ExitCode=%EXIT_CODE%)
  exit /b %EXIT_CODE%
)

exit /b 0
