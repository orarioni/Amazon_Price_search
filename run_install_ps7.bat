@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "MSI_FILE="
for /f "delims=" %%F in ('dir /b /a:-d /o-d "installers\PowerShell-*-win-*.msi" 2^>nul') do (
  if not defined MSI_FILE set "MSI_FILE=installers\%%F"
)

if not defined MSI_FILE (
  echo.
  echo [ERROR] PS7 installer MSI was not found under installers\.
  echo [HINT] Run run_prepare_ps7_installer.bat first.
  exit /b 1
)

echo Installing: %MSI_FILE%
msiexec /i "%MSI_FILE%" /passive /norestart ADD_PATH=1
if errorlevel 1 (
  echo.
  echo [ERROR] PowerShell 7 installation failed. Please run MSI manually.
  exit /b 1
)

echo.
echo [OK] PowerShell 7 installation command completed.
echo [NEXT] Close and reopen terminal, then run: pwsh -v
exit /b 0
