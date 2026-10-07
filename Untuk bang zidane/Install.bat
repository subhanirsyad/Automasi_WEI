@echo off
rem Klik dua kali SEKALI SAJA sebelum memakai aplikasi: memasang semua library (requirements.txt).
cd /d "%~dp0"
set "PY=python"
where python >nul 2>&1 || set "PY=py"

echo Memasang library yang dibutuhkan (butuh internet, bisa beberapa menit)...
echo.
%PY% -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo GAGAL memasang library. Pastikan Python 3.10 atau lebih baru sudah terpasang,
    echo centang "Add python.exe to PATH" saat memasangnya, dan internet aktif.
    pause
    exit /b 1
)

echo.
echo SELESAI. Sekarang klik dua kali Jalankan.bat untuk membuka aplikasi.
pause
