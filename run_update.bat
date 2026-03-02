@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\10_update_excel.ps1"
if errorlevel 1 (
  echo.
  echo [ERROR] 更新処理に失敗しました。メッセージを確認してください。
  exit /b 1
)
exit /b 0
