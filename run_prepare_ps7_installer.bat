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

set "MSI_FILE="
for /f "delims=" %%F in ('dir /b /a:-d /o-d "installers\PowerShell-*-win-*.msi" 2^>nul') do (
  if not defined MSI_FILE set "MSI_FILE=installers\%%F"
)

if not defined MSI_FILE (
  echo.
  echo [ERROR] PS7 installer MSI was not found under installers\.
  exit /b 1
)

echo.
echo Installing: %MSI_FILE%
msiexec /i "%MSI_FILE%" /passive /norestart ADD_PATH=1
if errorlevel 1 (
  echo.
  echo [ERROR] PowerShell 7 installation failed. Please run MSI manually.
  exit /b 1
)

set "PS7_EXE="
if exist "%ProgramFiles%\PowerShell\7\pwsh.exe" set "PS7_EXE=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not defined PS7_EXE if exist "%ProgramW6432%\PowerShell\7\pwsh.exe" set "PS7_EXE=%ProgramW6432%\PowerShell\7\pwsh.exe"
if not defined PS7_EXE if exist "%ProgramFiles(x86)%\PowerShell\7\pwsh.exe" set "PS7_EXE=%ProgramFiles(x86)%\PowerShell\7\pwsh.exe"
if not defined PS7_EXE if exist "%LocalAppData%\Programs\PowerShell\7\pwsh.exe" set "PS7_EXE=%LocalAppData%\Programs\PowerShell\7\pwsh.exe"

if not defined PS7_EXE (
  echo.
  echo [WARN] MSI finished but pwsh.exe was not found in common install paths.
  echo [HINT] Re-run installer as Administrator, then reopen terminal.
  exit /b 1
)

echo.
echo [OK] PowerShell 7 download and installation completed.
echo [OK] Detected pwsh: %PS7_EXE%
echo [NEXT] Close and reopen terminal, then run: pwsh -v
exit /b 0
