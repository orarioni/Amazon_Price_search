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

if %EXIT_CODE% neq 0 (
  echo.
  echo [ERROR] Init failed. See logs\run.log (ExitCode=%EXIT_CODE%)
  exit /b %EXIT_CODE%
)

exit /b 0