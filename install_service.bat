@echo off
chcp 65001 >nul
title Pasang Dashboard HPS - Service Jaringan Kantor
cd /d "%~dp0"

net session >nul 2>&1
if %errorlevel% neq 0 (
  echo.
  echo  Meminta hak administrator untuk firewall + service...
  powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs -Wait" 2>nul
  if errorlevel 1 goto tanpa_admin
  exit /b
)

:tanpa_admin
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_service.ps1"
echo.
pause
