@echo off
chcp 65001 >nul
title Dashboard HPS - Perbarui Data Lokal
cd /d "%~dp0"

echo.
echo  [1/2] Membaca Excel terbaru di folder ini...
python update_data.py
if errorlevel 1 goto gagal

echo.
echo  [2/2] Selesai. data.js sudah diperbarui.
echo         Untuk upload  : publish_data.bat (unggah ke GitHub Pages)
echo         Untuk lihat lokal : start_server.bat
echo.
pause
exit /b 0

:gagal
echo.
echo  GAGAL. Pastikan Python 3 terpasang, lalu jalankan: pip install openpyxl
echo.
pause
exit /b 1
