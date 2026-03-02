@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\00_init_secrets.ps1"
if errorlevel 1 (
  echo.
  echo [ERROR] 初期設定に失敗しました。メッセージを確認してください。
  exit /b 1
)
exit /b 0
