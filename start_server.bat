@echo off
chcp 65001 >nul
title Dashboard HPS - Server Lokal (cadangan)
cd /d "%~dp0"

echo.
echo  Menjalankan server lokal. Alamat untuk komputer lain:
python server.py 8080
echo.
pause
