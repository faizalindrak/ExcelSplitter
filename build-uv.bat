@echo off
setlocal EnableExtensions EnableDelayedExpansion

set PROJECT_NAME=ExcelSplitter

REM 1) Pastikan uv ada
where uv >nul 2>&1
if errorlevel 1 (
  echo [ERR] 'uv' tidak ditemukan. Install dulu dengan PowerShell:
  echo   iwr -useb https://astral.sh/uv/install.ps1 ^| iex
  exit /b 1
)

REM 2) Buat venv (kalau belum)
if not exist ".venv\Scripts\python.exe" (
  echo [INFO] Membuat venv (.venv) via uv ...
  uv venv .venv
)

REM 3) Install dependencies ke venv (pakai uv, bukan pip)
if exist requirements.txt (
  echo [INFO] Sync deps dari requirements.txt ...
  uv pip sync requirements.txt
) else (
  echo [INFO] Install deps langsung ...
  uv pip install pyinstaller customtkinter pandas openpyxl reportlab
)

REM 4) Build pakai pyinstaller (via uv run agar pakai .venv)
echo [INFO] Build exe...
uv run pyinstaller main.spec --clean

if errorlevel 1 (
  echo [ERR] Build gagal.
  exit /b 1
)

echo [OK ] Selesai. Cek: dist\%PROJECT_NAME%\%PROJECT_NAME%.exe
