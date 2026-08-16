@echo off
title AutoExpenses
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Start-AutoExpenses.ps1"
if errorlevel 1 pause
