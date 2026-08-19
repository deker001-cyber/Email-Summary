@echo off
chcp 65001 >nul
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Email-Summary.ps1"
if errorlevel 1 (
  echo.
  echo Email Summary failed to start. See error above.
  pause
)
