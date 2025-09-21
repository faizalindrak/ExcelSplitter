@echo off
setlocal EnableExtensions EnableDelayedExpansion

set PROJECT_NAME=ExcelSplitter
set VENV=.venv
set PY_EXE=%VENV%\Scripts\python.exe
set PIP_EXE=%VENV%\Scripts\pip.exe

echo === %PROJECT_NAME% Build Start ===

REM 1) Cek Python
where python >nul 2>&1 || goto NO_PY

REM 2) Buat venv jika belum ada
if exist "%PY_EXE%" goto HAVE_VENV
echo [INFO] Membuat virtualenv .venv
python -m venv "%VENV%" || goto FAIL
:HAVE_VENV

REM 3) Coba pakai pip dulu
"%PY_EXE%" -m pip --version >nul 2>&1 && goto USE_PIP

echo [INFO] pip tidak tersedia, akan mencoba uv
goto USE_UV

:USE_PIP
echo [INFO] Menggunakan pip
"%PY_EXE%" -m pip install --upgrade pip setuptools wheel || echo [WARN] Gagal upgrade pip
echo [INFO] Install dependencies (pip)
"%PIP_EXE%" install --upgrade pyinstaller customtkinter pandas openpyxl reportlab || goto FAIL

echo [INFO] Build exe (pip)
if exist "main.spec" (
  "%PY_EXE%" -m pyinstaller main.spec --clean || goto FAIL
) else (
  "%PY_EXE%" -m pyinstaller --noconsole --onefile --clean --name %PROJECT_NAME% main.py || goto FAIL
)
goto SUCCESS

:USE_UV
where uv >nul 2>&1 || goto NO_UV

if exist "%PY_EXE%" goto UV_HAVE_VENV
echo [INFO] Membuat venv dengan uv
uv venv "%VENV%" || goto FAIL
:UV_HAVE_VENV

if exist requirements.txt (
  echo [INFO] Sync dependencies dari requirements.txt (uv)
  uv pip sync requirements.txt || goto FAIL
) else (
  echo [INFO] Install dependencies (uv)
  uv pip install pyinstaller customtkinter pandas openpyxl reportlab || goto FAIL
)

echo [INFO] Build exe via uv run
if exist "main.spec" (
  uv run pyinstaller main.spec --clean || goto FAIL
) else (
  uv run pyinstaller --noconsole --onefile --clean --name %PROJECT_NAME% main.py || goto FAIL
)
goto SUCCESS

:SUCCESS
set OUT_EXE=dist\%PROJECT_NAME%\%PROJECT_NAME%.exe
if exist "%OUT_EXE%" (
  echo [OK ] Build sukses: %OUT_EXE%
) else (
  echo [OK ] Build selesai. Cek folder dist\
)
exit /b 0

:NO_PY
echo [ERR] Python tidak ditemukan di PATH. Install Python 3.10+ dan centang Add to PATH.
exit /b 1

:NO_UV
echo [ERR] uv tidak ditemukan. Install uv (PowerShell):
echo iwr -useb https://astral.sh/uv/install.ps1 ^| iex
exit /b 1

:FAIL
echo === Build FAILED ===
exit /b 1
