# main.spec
# Build: pyinstaller main.spec
# Hasil: dist/ExcelSplitter/ExcelSplitter.exe  (one-file)
# Catatan:
# - console=False => tidak munculkan jendela console
# - Mengumpulkan data/hiddenimports untuk reportlab, openpyxl, pandas, customtkinter
# - Jika ingin icon, set ICON_PATH di bawah.

from PyInstaller.utils.hooks import collect_submodules, collect_data_files
from PyInstaller.building.build_main import Analysis, PYZ, EXE
import sys
import os

# ---- opsi yang bisa kamu ubah ----
APP_NAME    = "ExcelSplitter"
ENTRY_SCRIPT= "main.py"
ICON_PATH   = None   # contoh: "app.ico"  (atau None jika tidak pakai)
CONSOLE     = False  # GUI app
UPX         = True   # butuh UPX terpasang agar efektif, kalau tidak ada tetap aman
# ----------------------------------

# Kumpulkan data & modul tambahan dari paket pihak-3
datas  = []
binaries = []
hiddenimports = []

# reportlab (PDF pure-Python)
datas += collect_data_files("reportlab", include_py_files=False)
hiddenimports += collect_submodules("reportlab")

# openpyxl (baca & tulis xlsx)
datas += collect_data_files("openpyxl", include_py_files=False)
hiddenimports += collect_submodules("openpyxl")

# pandas (IO excel & grouping)
hiddenimports += collect_submodules("pandas")

# customtkinter (tema/gambar jika ada)
datas += collect_data_files("customtkinter", include_py_files=False)
hiddenimports += collect_submodules("customtkinter")

# tkinter assets (umumnya sudah di-bundle otomatis oleh PyInstaller)
# Tidak perlu tambahan khusus, tapi tetap aman bila runtime berbeda.

block_cipher = None

a = Analysis(
    [ENTRY_SCRIPT],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],          # bisa tambahkan path hook kustom di sini
    hooksconfig={},        # config hook opsional
    runtime_hooks=[],      # runtime hook opsional
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=UPX,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=CONSOLE,
    icon=ICON_PATH
)
