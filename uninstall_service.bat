@echo off
chcp 65001 >nul
title Copot Dashboard HPS - Service
cd /d "%~dp0"

net session >nul 2>&1
if %errorlevel% neq 0 (
  echo  Meminta hak administrator untuk firewall...
  powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs -Wait" 2>nul
)

echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall_service.ps1"
echo.
pause
