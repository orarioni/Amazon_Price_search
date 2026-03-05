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
  echo [HINT] Run run_prepare_ps7_installer.bat first, then install PowerShell 7.
  exit /b 1
)

"%PS7_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\10_update_excel.ps1"
if errorlevel 1 (
  echo.
  echo [ERROR] Update failed. Check the message above.
  exit /b 1
)
exit /b 0
