@echo off
rem Klik dua kali file ini untuk membuka aplikasi. Library dipasang otomatis kalau belum ada.
cd /d "%~dp0"
set "PY=python"
where python >nul 2>&1 || set "PY=py"

%PY% -c "import customtkinter, openpyxl, playwright, googleapiclient, google.oauth2" >nul 2>&1
if errorlevel 1 (
    echo Memasang library yang dibutuhkan, tunggu sebentar...
    %PY% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo Gagal memasang library. Pastikan Python 3.10+ terpasang dan internet aktif.
        pause
        exit /b 1
    )
)

%PY% app.py
if errorlevel 1 pause
