@echo off
title MatchaIn GC - INSTALL AND RUN

echo ===================================================
echo   MatchaIn GC (Matcha Input Gak Culun) - SETUP
echo ===================================================
echo.

:: 1. Cek Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python tidak ditemukan! Harap install Python terlebih dahulu.
    echo         Link Download: https://www.python.org/downloads/
    pause
    exit /b
)

:: 2. Cek/Buat Virtual Environment
if not exist ".venv" (
    echo [INFO] Membuat virtual environment...
    python -m venv .venv
)

:: 3. Aktifkan Virtual Environment
echo [INFO] Mengaktifkan virtual environment...
call .venv\Scripts\activate

:: 4. Install Dependencies
if exist "requirements.txt" (
    echo [INFO] Menginstall library...
    pip install -r requirements.txt
) else (
    echo [WARNING] File requirements.txt tidak ditemukan.
)

:: 5. Jalankan Aplikasi
echo.
echo [INFO] Instalasi Selesai...
echo ===================================================

echo.
echo ===================================================
echo   Silakan jalankan run.bat.
echo ===================================================
pause
