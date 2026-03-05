@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PS7_EXE="
set "PF64=%ProgramFiles%"
set "PF86=%ProgramFiles(x86)%"

if defined PF64 if exist "%PF64%\PowerShell\7\pwsh.exe" set "PS7_EXE=%PF64%\PowerShell\7\pwsh.exe"
if not defined PS7_EXE if defined PF86 if exist "%PF86%\PowerShell\7\pwsh.exe" set "PS7_EXE=%PF86%\PowerShell\7\pwsh.exe"
if not defined PS7_EXE for %%I in (pwsh.exe) do set "PS7_EXE=%%~$PATH:I"

if not defined PS7_EXE (
  echo.
  echo [ERROR] PowerShell 7 / pwsh was not found.
  echo [HINT] Run run_prepare_ps7_installer.bat (download + install).
  exit /b 1
)

"%PS7_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\00_init_secrets.ps1"
set "EXIT_CODE=%ERRORLEVEL%"

if %EXIT_CODE% neq 0 (
  echo.
  echo [ERROR] Initial setup failed. Check the message above. ExitCode=%EXIT_CODE%
  exit /b %EXIT_CODE%
)

exit /b 0
