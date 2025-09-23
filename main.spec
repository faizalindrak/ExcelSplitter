# main.spec
# Build: pyinstaller main.spec
# Hasil: dist/ExcelSplitter/ExcelSplitter.exe  (one-file)
# Catatan:
# - console=False => tidak munculkan jendela console
# - Mengumpulkan data/hiddenimports untuk reportlab, openpyxl, pandas, customtkinter
# - Termasuk pywin32 dan xlwings untuk Excel COM automation
# - Jika ingin icon, set ICON_PATH di bawah.
#
# Dependencies yang dibutuhkan:
# pip install pandas openpyxl customtkinter reportlab xlwings pywin32

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

# xlwings (Excel COM automation)
try:
    datas += collect_data_files("xlwings", include_py_files=False)
    hiddenimports += collect_submodules("xlwings")
except:
    pass  # xlwings mungkin tidak terpasang

# pywin32 (Windows COM interface)
try:
    hiddenimports += collect_submodules("win32com")
    hiddenimports += collect_submodules("pythoncom")
    hiddenimports += collect_submodules("pywintypes")
    hiddenimports += collect_submodules("win32api")
    hiddenimports += collect_submodules("win32gui")
    hiddenimports += collect_submodules("win32con")

    # Tambahan hidden imports yang sering dibutuhkan
    hiddenimports += [
        "win32com.client",
        "win32com.client.gencache",
        "win32com.gen_py",
        "pythoncom",
        "pywintypes",
        "win32timezone"
    ]
except:
    pass  # pywin32 mungkin tidak terpasang

# psutil (untuk process management, opsional)
try:
    hiddenimports += collect_submodules("psutil")
except:
    pass

# winreg (Windows registry, biasanya built-in tapi kadang perlu explicit)
try:
    hiddenimports += ["winreg"]
except:
    pass

# Tambahan imports untuk Excel COM yang sering missing
hiddenimports += [
    "pkg_resources.py2_warn",
    "pkg_resources.markers",
    "email.mime.multipart",
    "email.mime.text",
    "email.mime.base",
    "encodings.idna",
    "encodings.utf_8",
    "encodings.cp1252"
]

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
