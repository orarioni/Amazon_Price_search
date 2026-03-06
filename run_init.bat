@echo off
cd /d "%~dp0"

set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS_EXE%" set "PS_EXE=powershell"

set "PS_FILE=%~dp0scripts\00_init_secrets.ps1"
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS_FILE%"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if "%EXIT_CODE%"=="0" (
  echo [OK] Initialization completed.
) else (
  echo [ERROR] Initialization failed. ExitCode=%EXIT_CODE%
)

echo.
echo Press any key to close this window...
pause >nul
exit /b %EXIT_CODE%
