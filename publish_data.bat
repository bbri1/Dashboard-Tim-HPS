@echo off
chcp 65001 >nul
title Dashboard HPS - Unggah Data ke GitHub Pages
cd /d "%~dp0"

echo.
echo  ============================================================
echo   Upload data HPS ke GitHub Pages
echo   1. Simpan file Excel TERBARU ke folder ini
echo   2. Klik 2x file ini
echo   3. Tunggu 1-3 menit, situs jadi baru
echo  ============================================================
echo.

python publish_data.py
echo.
pause
