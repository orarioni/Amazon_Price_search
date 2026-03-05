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

"%PS7_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\10_update_excel.ps1"
if errorlevel 1 (
  echo.
  echo [ERROR] 更新処理に失敗しました。メッセージを確認してください。
  exit /b 1
)
exit /b 0
