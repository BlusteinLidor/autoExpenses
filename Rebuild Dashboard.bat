@echo off
title AutoExpenses - Rebuild dashboard
cd /d "%~dp0"
echo.
echo Rebuilding the dashboard (frontend-next)...
echo.
cd frontend-next
call npm run build
if errorlevel 1 (
  echo.
  echo Build failed.
  pause
  exit /b 1
)
echo.
echo Done. You can close this window, then use "Start AutoExpenses.bat".
pause
