@echo off
chcp 65001 >nul
cd /d "%~dp0"

rem Use explicit path to PowerShell if available to avoid PATH/parse issues
set "PS_EXE=%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS_EXE%" (
  rem fallback to powershell on PATH
  set "PS_EXE=powershell"
)

rem Build safe file path and arguments (keep quoting for the -File arg)
set "PS_FILE=%~dp0scripts\00_init_secrets.ps1"
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS_FILE%"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if %EXIT_CODE% neq 0 (
  echo [ERROR] 初期設定に失敗しました。メッセージを確認してください。 (ExitCode=%EXIT_CODE%)
) else (
  echo [OK] 初期設定が完了しました。
)

echo.
echo 終了するには何かキーを押してください...
pause >nul
exit /b %EXIT_CODE%
