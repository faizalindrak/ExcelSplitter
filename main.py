# split_gui.py
# GUI untuk split Excel per nilai unik dengan template-based rendering.
# Build exe: pyinstaller split_gui.spec
# Dependencies: PySide6, PySide6-Fluent-Widgets, pandas, openpyxl

import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
import configparser

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from PySide6.QtCore import Qt, Signal, QThread
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QFileDialog
)
from qfluentwidgets import (
    ScrollArea, SimpleCardWidget,
    LineEdit, ComboBox, PushButton, PrimaryPushButton,
    ProgressBar, SpinBox, TextEdit, InfoBar, InfoBarPosition,
    ToolButton, SubtitleLabel, BodyLabel, CaptionLabel,
    setTheme, Theme, FluentIcon as FIF
)

# ==== (Opsional) xlwings untuk PDF via Excel COM ====
try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except Exception:
    XLWINGS_AVAILABLE = False


# ----------------- Helpers -----------------

def safe_file_part(s: str) -> str:
    s = "" if s is None else str(s)
    return re.sub(r'[:\\/\?\*\[\]<>|"]', "_", s).strip() or "Key"

def set_print_titles_and_area(ws, header_rows: int, last_col_idx: int, last_data_row: int):
    ws.print_title_rows = f"1:{header_rows}"
    last_col_letter = get_column_letter(last_col_idx if last_col_idx > 0 else 1)
    last_row = last_data_row if last_data_row >= (header_rows + 1) else (header_rows + 1)
    ws.print_area = f"A1:{last_col_letter}{last_row}"

def find_soffice(explicit_path: str | None = None) -> str | None:
    """Cari soffice.exe dari explicit, env, PATH, atau lokasi umum."""
    def _normalize(p):
        if not p:
            return None
        p = str(Path(p))
        if p.lower().endswith("soffice.exe") and Path(p).exists():
            return p
        prog = Path(p) / "soffice.exe"
        if prog.exists():
            return str(prog)
        prog = Path(p) / "program" / "soffice.exe"
        if prog.exists():
            return str(prog)
        return None

    # explicit
    s = _normalize(explicit_path)
    if s: return s
    # env
    s = _normalize(os.environ.get("LIBREOFFICE_PATH"))
    if s: return s
    # PATH
    s = shutil.which("soffice")
    if s: return s
    # lokasi umum
    common = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        r"C:\PortableApps\LibreOfficePortable\App\libreoffice\program\soffice.exe",
        r"D:\PortableApps\LibreOfficePortable\App\libreoffice\program\soffice.exe",
    ]
    for c in common:
        if Path(c).exists():
            return c
    return None

def export_pdf_via_lo(xlsx_path: Path, soffice_path: str | None = None):
    exe = soffice_path or "soffice"
    cmd = [exe, "--headless", "--convert-to", "pdf", "--outdir", str(xlsx_path.parent), str(xlsx_path)]
    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

def cleanup_excel_com():
    """Clean up Excel COM objects and release resources"""
    try:
        # Try to release any existing COM objects
        import gc
        import pythoncom

        # Force garbage collection
        gc.collect()

        # Uninitialize COM
        try:
            pythoncom.CoUninitialize()
        except:
            pass

        # Re-initialize COM for next use
        try:
            pythoncom.CoInitialize()
        except:
            pass

    except ImportError:
        # pythoncom not available, try basic cleanup
        import gc
        gc.collect()
    except:
        pass

def debug_excel_detection():
    """Debug function to check Excel detection methods"""
    results = []

    results.append(f"XLWINGS_AVAILABLE: {XLWINGS_AVAILABLE}")

    # Test Method 1: xlwings
    if XLWINGS_AVAILABLE:
        try:
            import xlwings as xw
            app = xw.App(visible=False, add_book=False)
            if app is not None:
                app.quit()
                results.append("Method 1 (xlwings): SUCCESS")
            else:
                results.append("Method 1 (xlwings): FAILED - app is None")
        except Exception as e:
            results.append(f"Method 1 (xlwings): FAILED - {str(e)}")
    else:
        results.append("Method 1 (xlwings): SKIPPED - xlwings not available")

    # Test Method 2: win32com.client.Dispatch
    try:
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.Quit()
        del excel
        results.append("Method 2 (win32com Dispatch): SUCCESS")
    except Exception as e:
        results.append(f"Method 2 (win32com Dispatch): FAILED - {str(e)}")

    # Test Method 3: win32com.client.gencache.EnsureDispatch
    try:
        import win32com.client
        excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        excel.Quit()
        del excel
        results.append("Method 3 (win32com EnsureDispatch): SUCCESS")
    except Exception as e:
        results.append(f"Method 3 (win32com EnsureDispatch): FAILED - {str(e)}")

    # Test Method 4: Registry check
    try:
        import winreg
        found_keys = []
        key_paths = [
            r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
            r"SOFTWARE\Microsoft\Office\16.0\Excel\InstallRoot",
            r"SOFTWARE\Microsoft\Office\15.0\Excel\InstallRoot",
            r"SOFTWARE\Microsoft\Office\14.0\Excel\InstallRoot",
            r"SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Excel\InstallRoot",
            r"SOFTWARE\WOW6432Node\Microsoft\Office\15.0\Excel\InstallRoot",
            r"SOFTWARE\WOW6432Node\Microsoft\Office\14.0\Excel\InstallRoot",
        ]

        for key_path in key_paths:
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                winreg.CloseKey(key)
                found_keys.append(key_path)
            except:
                continue

        if found_keys:
            results.append(f"Method 4 (Registry): Found {len(found_keys)} keys: {found_keys}")
        else:
            results.append("Method 4 (Registry): No Excel keys found")

    except ImportError:
        results.append("Method 4 (Registry): SKIPPED - winreg not available")
    except Exception as e:
        results.append(f"Method 4 (Registry): FAILED - {str(e)}")

    return results

def check_excel_availability():
    """Check if Microsoft Excel is available for COM automation"""
    try:
        # Method 1: Try xlwings first (most reliable for xlwings usage)
        if XLWINGS_AVAILABLE:
            try:
                import xlwings as xw
                app = xw.App(visible=False, add_book=False)
                if app is not None:
                    app.quit()
                    return True
            except Exception:
                pass

        # Method 2: Try win32com.client
        try:
            import win32com.client
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.Quit()
            del excel
            return True
        except Exception:
            pass

        # Method 3: Try alternative COM approach
        try:
            import win32com.client
            excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
            excel.Visible = False
            excel.Quit()
            del excel
            return True
        except Exception:
            pass

        # Method 4: Check registry for Excel installation
        try:
            import winreg
            key_paths = [
                r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
                r"SOFTWARE\Microsoft\Office\16.0\Excel\InstallRoot",
                r"SOFTWARE\Microsoft\Office\15.0\Excel\InstallRoot",
                r"SOFTWARE\Microsoft\Office\14.0\Excel\InstallRoot",
                r"SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Excel\InstallRoot",
                r"SOFTWARE\WOW6432Node\Microsoft\Office\15.0\Excel\InstallRoot",
                r"SOFTWARE\WOW6432Node\Microsoft\Office\14.0\Excel\InstallRoot",
            ]

            for key_path in key_paths:
                try:
                    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                    winreg.CloseKey(key)
                    # Found Excel in registry, but still need to test COM
                    try:
                        import win32com.client
                        excel = win32com.client.Dispatch("Excel.Application")
                        excel.Visible = False
                        excel.Quit()
                        return True
                    except:
                        pass
                except:
                    continue
        except ImportError:
            pass  # winreg not available
        except Exception:
            pass

        return False

    except Exception:
        return False
    finally:
        # Clean up after test
        try:
            cleanup_excel_com()
        except:
            pass

def export_pdf_via_xlwings(xlsx_path: Path):
    """Export Excel to PDF using xlwings (requires Excel installed on Windows)"""
    if not XLWINGS_AVAILABLE:
        raise RuntimeError("xlwings belum terpasang. Jalankan: pip install xlwings")

    # Clean up COM objects only (safe cleanup)
    cleanup_excel_com()

    # Check if Excel is available before proceeding
    if not check_excel_availability():
        raise RuntimeError("Microsoft Excel tidak terinstall atau tidak dapat diakses. "
                          "Gunakan PDF Engine 'libreoffice' atau 'none' sebagai alternatif.")

    pdf_path = xlsx_path.with_suffix(".pdf")

    # Use xlwings with invisible Excel application
    app = None
    wb = None
    try:
        # Create Excel application instance (invisible)
        app = xw.App(visible=False, add_book=False)

        # Verify app was created successfully
        if app is None:
            raise RuntimeError("Failed to create Excel application instance")

        # Open the workbook
        wb = app.books.open(str(xlsx_path))

        if wb is None:
            raise RuntimeError(f"Failed to open workbook: {xlsx_path}")

        # Get the active worksheet
        ws = wb.sheets.active

        if ws is None:
            raise RuntimeError("Failed to get active worksheet")

        # Export to PDF with Excel's native formatting
        # This preserves all Excel formatting, styles, colors, etc.
        ws.to_pdf(str(pdf_path))

        # Verify PDF was created
        if not pdf_path.exists():
            raise RuntimeError(f"PDF was not created: {pdf_path}")

    except Exception as e:
        # More detailed error reporting
        error_details = str(e)
        if "apps" in error_details.lower() or "nonetype" in error_details.lower():
            error_details = ("Excel COM interface tidak dapat diakses. "
                           "Pastikan Microsoft Excel terinstall dan tidak sedang digunakan aplikasi lain. "
                           "Alternatif: gunakan PDF Engine 'libreoffice' atau 'none'.")
        raise RuntimeError(f"Gagal export PDF via xlwings: {error_details}")
    finally:
        # Clean up: close workbook and quit Excel
        if wb is not None:
            try:
                wb.close()
            except:
                pass
        if app is not None:
            try:
                app.quit()
            except:
                pass

        # Safe COM cleanup only after each PDF export
        cleanup_excel_com()

        # Give a moment for cleanup to complete
        import time
        time.sleep(0.5)


# ----------------- Split Logic -----------------

def split_excel_with_template(
    source_path: Path, sheet_name: str, key_col, template_path: Path, out_dir: Path,
    header_rows: int, pdf_engine: str = "xlwings", soffice_path: str | None = None,
    prefix: str = "", suffix: str = "", status_cb=None, progress_cb=None
):
    if status_cb is None: status_cb = lambda msg: None
    if progress_cb is None: progress_cb = lambda t, c: None

    if not source_path.exists():
        raise FileNotFoundError(f"Sumber tidak ditemukan: {source_path}")
    if not template_path.exists():
        raise FileNotFoundError(f"Template tidak ditemukan: {template_path}")
    out_dir.mkdir(parents=True, exist_ok=True)

    status_cb("Membaca sumber...")

    # Diagnostic logging
    try:
        file_size = source_path.stat().st_size if source_path.exists() else 0
        status_cb(f"Debug: File path: {source_path}")
        status_cb(f"Debug: File exists: {source_path.exists()}")
        status_cb(f"Debug: File size: {file_size:,} bytes ({file_size/1024/1024:.2f} MB)")
        status_cb(f"Debug: Sheet name: '{sheet_name}'")
        status_cb("Debug: Starting pd.read_excel...")

        # Try reading with timeout and error handling
        import time
        start_time = time.time()

        # First try to read just the header to test file accessibility
        status_cb("Debug: Testing file accessibility...")
        try:
            df_test = pd.read_excel(source_path, sheet_name=sheet_name, nrows=5, dtype=object)
            status_cb(f"Debug: Successfully read {len(df_test)} rows for testing")
        except Exception as test_e:
            status_cb(f"Debug: Test read failed: {str(test_e)}")
            raise test_e

        # Read with header at the correct row (header_rows is 1-indexed)
        df = pd.read_excel(source_path, sheet_name=sheet_name, header=header_rows - 1, dtype=object)

        elapsed = time.time() - start_time
        status_cb(f"Debug: Successfully read {len(df)} rows in {elapsed:.2f} seconds")

    except FileNotFoundError as e:
        status_cb(f"Debug: File not found: {e}")
        raise FileNotFoundError(f"File tidak ditemukan: {source_path}")
    except PermissionError as e:
        status_cb(f"Debug: Permission denied: {e}")
        raise PermissionError(f"Tidak ada akses ke file: {source_path}")
    except Exception as e:
        status_cb(f"Debug: Error reading Excel: {type(e).__name__}: {str(e)}")
        raise e

    # Tentukan kolom kunci
    if isinstance(key_col, int):
        if key_col < 1 or key_col > df.shape[1]:
            raise ValueError("Index kolom kunci di luar jangkauan DataFrame.")
        key_series = df.iloc[:, key_col - 1]
    else:
        if key_col not in df.columns:
            raise ValueError(f"Header kolom kunci '{key_col}' tidak ditemukan.")
        key_series = df[key_col]

    # Debug: Check for categorical data issues
    status_cb(f"Debug: Key column '{key_col}' data type: {key_series.dtype}")
    status_cb(f"Debug: Key column unique values: {len(key_series.unique())}")
    status_cb(f"Debug: Key column has null values: {key_series.isnull().sum()}")

    # Check for categorical columns in the entire DataFrame
    categorical_cols = []
    for col in df.columns:
        if df[col].dtype.name == 'category':
            categorical_cols.append(col)
            status_cb(f"Debug: Found categorical column: {col}")

    if categorical_cols:
        status_cb(f"Debug: Converting {len(categorical_cols)} categorical columns to string...")
        for col in categorical_cols:
            try:
                df[col] = df[col].astype(str)
                status_cb(f"Debug: Converted {col} to string")
            except Exception as conv_e:
                status_cb(f"Debug: Failed to convert {col}: {conv_e}")

    # Selaraskan urutan kolom ke header template jika cocok
    templ_cols = None
    templ_col_start = 1
    try:
        wb_probe = load_workbook(template_path, read_only=True, data_only=True)
        ws_probe = wb_probe.active
        tmp, c, empty_streak = [], 1, 0
        first_col_found = False
        while c <= 500 and empty_streak < 5:
            val = ws_probe.cell(row=header_rows, column=c).value
            if val is None or str(val).strip() == "":
                empty_streak += 1
            else:
                if not first_col_found:
                    templ_col_start = c
                    first_col_found = True
                tmp.append(str(val).strip()); empty_streak = 0
            c += 1
        wb_probe.close()
        if tmp:
            templ_cols = tmp
    except Exception:
        templ_cols = None

    if templ_cols:
        exist = [col for col in templ_cols if col in df.columns]
        if exist:
            df = df[exist]

    # Debug: Check groupby operation
    status_cb("Debug: Starting groupby operation...")
    try:
        groups = df.groupby(key_series, dropna=False, sort=False)
        total, current = len(groups), 0
        status_cb(f"Debug: Groupby successful, found {total} groups")
        progress_cb(total, 0)
    except Exception as groupby_e:
        status_cb(f"Debug: Groupby error: {type(groupby_e).__name__}: {str(groupby_e)}")

        # Try multiple approaches to fix the issue
        if "categorical" in str(groupby_e).lower():
            status_cb("Debug: Attempting to fix categorical issue...")

            # Method 1: Try converting all categorical columns to object
            try:
                status_cb("Debug: Method 1 - Converting all categorical columns to object...")
                df_no_cat = df.copy()
                for col in df_no_cat.columns:
                    if df_no_cat[col].dtype.name == 'category':
                        df_no_cat[col] = df_no_cat[col].astype('object')
                groups = df_no_cat.groupby(key_series, dropna=False, sort=False)
                total, current = len(groups), 0
                status_cb(f"Debug: Method 1 successful, found {total} groups")
                progress_cb(total, 0)
                # Update df to use the fixed version
                df = df_no_cat
            except Exception as method1_e:
                status_cb(f"Debug: Method 1 failed: {method1_e}")

                # Method 2: Try using string conversion for groupby
                try:
                    status_cb("Debug: Method 2 - Using string keys for groupby...")
                    string_keys = key_series.astype(str)
                    groups = df.groupby(string_keys, dropna=False, sort=False)
                    total, current = len(groups), 0
                    status_cb(f"Debug: Method 2 successful, found {total} groups")
                    progress_cb(total, 0)
                except Exception as method2_e:
                    status_cb(f"Debug: Method 2 failed: {method2_e}")
                    raise groupby_e
        else:
            raise groupby_e

    for key_val, group in groups:
        current += 1
        status_cb(f"Proses [{current}/{total}] key={key_val}")
        progress_cb(total, current)

        # 1) Tulis XLSX dari template
        wb = load_workbook(template_path)
        ws = wb.active
        start_row = header_rows + 1

        # Comprehensive cleanup to eliminate all Excel repair warnings
        try:
            # 1. Remove all named ranges (common cause of repair warnings)
            if hasattr(wb, 'defined_names'):
                wb.defined_names = {}

            # 2. Remove external links
            if hasattr(wb, '_external_links') and wb._external_links:
                wb._external_links.clear()

            # 3. Remove external link relationships
            if hasattr(wb, '_external_link_rels'):
                wb._external_link_rels.clear()

            # 4. Remove all drawings completely (they often cause repair issues)
            if hasattr(ws, '_drawings'):
                ws._drawings = []

            # 5. Clean up worksheet-level named ranges
            if hasattr(ws, 'defined_names'):
                ws.defined_names = []

            # 6. Remove external link parts from the package
            try:
                from openpyxl.packaging.relationship import Relationship
                if hasattr(wb, '_rels'):
                    # Remove external link relationships
                    rels_to_remove = []
                    for rel in wb._rels:
                        if ('externalLink' in str(rel.target) or
                            'externalLinks' in str(rel.target) or
                            'drawing' in str(rel.target).lower()):
                            rels_to_remove.append(rel)
                    for rel in rels_to_remove:
                        wb._rels.remove(rel)
            except:
                pass

            # 7. Clean up any hyperlinks that might reference external data
            for row in ws.iter_rows():
                for cell in row:
                    if cell.hyperlink:
                        cell.hyperlink = None

        except Exception as cleanup_e:
            # If cleanup fails, continue anyway
            pass

        values = group.fillna("").values.tolist()

        # Copy formatting from template row to data rows
        # Use the first data row from template as formatting template
        template_row = start_row  # This should be the first data row in template

        for r_off, row_vals in enumerate(values, start=0):
            row_idx = start_row + r_off

            for c_idx, v in enumerate(row_vals, start=templ_col_start):
                ws.cell(row=row_idx, column=c_idx, value=v)

            if r_off > 0:
                try:
                    for col_idx in range(templ_col_start, templ_col_start + len(row_vals)):
                        template_cell = ws.cell(row=template_row, column=col_idx)
                        current_cell = ws.cell(row=row_idx, column=col_idx)

                        # Copy all formatting properties
                        if template_cell.has_style:
                            current_cell.font = template_cell.font.copy()
                            current_cell.fill = template_cell.fill.copy()
                            current_cell.border = template_cell.border.copy()
                            current_cell.alignment = template_cell.alignment.copy()
                            current_cell.number_format = template_cell.number_format
                            current_cell.protection = template_cell.protection.copy()

                except Exception as format_e:
                    # If formatting copy fails, continue without it
                    pass

        last_data_row = start_row + len(values) - 1
        last_col = templ_col_start + group.shape[1] - 1
        set_print_titles_and_area(ws, header_rows, max(1, last_col), last_data_row)

        # Build filename with prefix and suffix
        key_part = safe_file_part(key_val)
        parts = []
        if prefix:
            parts.append(prefix)
        parts.append(key_part)
        if suffix:
            parts.append(suffix)

        out_name = " ".join(parts)
        xlsx_out = out_dir / f"{out_name}.xlsx"
        wb.save(xlsx_out)

        # 2) PDF (opsional)
        eng = (pdf_engine or "none").lower()
        if eng != "none":
            if eng == "libreoffice":
                export_pdf_via_lo(xlsx_out, soffice_path=soffice_path)
            elif eng == "xlwings":
                export_pdf_via_xlwings(xlsx_out)

    status_cb("Selesai.")
    progress_cb(total, total)

    # Final cleanup for Excel COM sessions
    if (pdf_engine or "none").lower() == "xlwings":
        cleanup_excel_com()


# ----------------- GUI -----------------

class SplitWorker(QThread):
    status = Signal(str)
    progress = Signal(int, int)
    finished = Signal()
    error = Signal(str)

    def __init__(self, params):
        super().__init__()
        self.params = params

    def run(self):
        try:
            split_excel_with_template(
                source_path=self.params['source_path'],
                sheet_name=self.params['sheet_name'],
                key_col=self.params['key_col'],
                template_path=self.params['template_path'],
                out_dir=self.params['out_dir'],
                header_rows=self.params['header_rows'],
                pdf_engine=self.params['pdf_engine'],
                soffice_path=self.params['soffice_path'],
                prefix=self.params['prefix'],
                suffix=self.params['suffix'],
                status_cb=self.status.emit,
                progress_cb=lambda t, c: self.progress.emit(t, c)
            )
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))


class AccordionCard(SimpleCardWidget):
    def __init__(self, title, icon=None, parent=None):
        super().__init__(parent)
        self.setBorderRadius(8)
        self._expanded = True

        self._main_layout = QVBoxLayout(self)
        self._main_layout.setContentsMargins(0, 0, 0, 0)
        self._main_layout.setSpacing(0)

        self._header = QWidget()
        self._header.setCursor(Qt.PointingHandCursor)
        self._header.setFixedHeight(48)
        header_layout = QHBoxLayout(self._header)
        header_layout.setContentsMargins(16, 0, 16, 0)

        if icon:
            icon_widget = ToolButton(icon)
            icon_widget.setFixedSize(20, 20)
            icon_widget.setEnabled(False)
            header_layout.addWidget(icon_widget)

        self._title_label = SubtitleLabel(title)
        header_layout.addWidget(self._title_label)
        header_layout.addStretch()

        self._toggle_btn = ToolButton(FIF.CHEVRON_DOWN)
        self._toggle_btn.setFixedSize(20, 20)
        self._toggle_btn.clicked.connect(self.toggle)
        header_layout.addWidget(self._toggle_btn)

        self._main_layout.addWidget(self._header)

        self._content = QWidget()
        self._content_layout = QVBoxLayout(self._content)
        self._content_layout.setContentsMargins(16, 8, 16, 16)
        self._content_layout.setSpacing(12)
        self._main_layout.addWidget(self._content)

        self._header.mousePressEvent = lambda e: self.toggle()

    @property
    def content_layout(self):
        return self._content_layout

    def toggle(self):
        self._expanded = not self._expanded
        self._content.setVisible(self._expanded)
        self._toggle_btn.setIcon(FIF.CHEVRON_DOWN if self._expanded else FIF.CHEVRON_RIGHT)

    def collapse(self):
        self._expanded = False
        self._content.setVisible(False)
        self._toggle_btn.setIcon(FIF.CHEVRON_RIGHT)

    def expand(self):
        self._expanded = True
        self._content.setVisible(True)
        self._toggle_btn.setIcon(FIF.CHEVRON_DOWN)
class SplitApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Splitter")
        self.resize(1000, 750)
        setTheme(Theme.DARK)

        self.is_running = False
        self.worker = None

        self._build_ui()

    def _build_ui(self):
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(0, 0, 0, 0)

        toolbar = QHBoxLayout()
        toolbar.setContentsMargins(16, 12, 16, 4)
        self.btn_save_ini = PushButton(FIF.SAVE, "Save .ini")
        self.btn_save_ini.clicked.connect(self.save_ini)
        self.btn_load_ini = PushButton(FIF.FOLDER, "Load .ini")
        self.btn_load_ini.clicked.connect(self.load_ini)
        self.lbl_loaded_ini = CaptionLabel("")
        toolbar.addWidget(self.btn_save_ini)
        toolbar.addWidget(self.btn_load_ini)
        toolbar.addWidget(self.lbl_loaded_ini)
        toolbar.addStretch()
        root_layout.addLayout(toolbar)

        scroll = ScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_widget = QWidget()
        self.scroll_layout = QVBoxLayout(scroll_widget)
        self.scroll_layout.setContentsMargins(16, 8, 16, 16)
        self.scroll_layout.setSpacing(12)
        scroll.setWidget(scroll_widget)
        root_layout.addWidget(scroll)

        self._build_source_card()
        self._build_template_card()
        self._build_output_card()
        self._build_actions_card()

        self.scroll_layout.addStretch()

    def _build_source_card(self):
        card = AccordionCard("Source", FIF.DOCUMENT)
        layout = card.content_layout

        row1 = QHBoxLayout()
        self.edit_source = LineEdit()
        self.edit_source.setPlaceholderText("Path to source Excel file...")
        self.btn_browse_source = ToolButton(FIF.FOLDER)
        self.btn_browse_source.clicked.connect(self.browse_source)
        row1.addWidget(self.edit_source)
        row1.addWidget(self.btn_browse_source)
        layout.addLayout(row1)

        row2 = QHBoxLayout()
        self.cmb_sheet = ComboBox()
        self.cmb_sheet.setPlaceholderText("Sheet")
        self.cmb_sheet.setMinimumWidth(200)
        self.btn_load_sheets = PushButton("Load Sheets")
        self.btn_load_sheets.clicked.connect(self.load_sheets)
        row2.addWidget(BodyLabel("Sheet:"))
        row2.addWidget(self.cmb_sheet)
        row2.addWidget(self.btn_load_sheets)
        row2.addStretch()
        layout.addLayout(row2)

        row3 = QHBoxLayout()
        self.cmb_key = ComboBox()
        self.cmb_key.setPlaceholderText("Key Column")
        self.cmb_key.setMinimumWidth(200)
        self.btn_load_headers = PushButton("Load Headers")
        self.btn_load_headers.clicked.connect(self.load_headers)
        row3.addWidget(BodyLabel("Key Column:"))
        row3.addWidget(self.cmb_key)
        row3.addWidget(self.btn_load_headers)
        row3.addStretch()
        layout.addLayout(row3)

        self.scroll_layout.addWidget(card)

    def _build_template_card(self):
        card = AccordionCard("Template", FIF.EDIT)
        layout = card.content_layout

        row1 = QHBoxLayout()
        self.edit_template = LineEdit()
        self.edit_template.setPlaceholderText("Path to template Excel file...")
        self.btn_browse_template = ToolButton(FIF.FOLDER)
        self.btn_browse_template.clicked.connect(self.browse_template)
        row1.addWidget(self.edit_template)
        row1.addWidget(self.btn_browse_template)
        layout.addLayout(row1)

        row2 = QHBoxLayout()
        self.spin_header_rows = SpinBox()
        self.spin_header_rows.setRange(1, 100)
        self.spin_header_rows.setValue(5)
        self.spin_header_rows.setFixedWidth(100)
        row2.addWidget(BodyLabel("Header Rows:"))
        row2.addWidget(self.spin_header_rows)
        row2.addStretch()
        layout.addLayout(row2)

        self.scroll_layout.addWidget(card)

    def _build_output_card(self):
        card = AccordionCard("Output", FIF.FOLDER)
        layout = card.content_layout

        row1 = QHBoxLayout()
        self.edit_outdir = LineEdit()
        self.edit_outdir.setPlaceholderText("Output folder...")
        self.btn_browse_outdir = ToolButton(FIF.FOLDER)
        self.btn_browse_outdir.clicked.connect(self.browse_outdir)
        row1.addWidget(self.edit_outdir)
        row1.addWidget(self.btn_browse_outdir)
        layout.addLayout(row1)

        row2 = QHBoxLayout()
        self.cmb_pdf_engine = ComboBox()
        self.cmb_pdf_engine.addItems(["xlwings", "libreoffice", "none"])
        self.cmb_pdf_engine.setCurrentIndex(0)
        self.cmb_pdf_engine.setFixedWidth(180)
        row2.addWidget(BodyLabel("PDF Engine:"))
        row2.addWidget(self.cmb_pdf_engine)
        row2.addStretch()
        layout.addLayout(row2)

        row3 = QHBoxLayout()
        self.edit_lo_path = LineEdit()
        self.edit_lo_path.setPlaceholderText("Path to soffice.exe (optional)...")
        self.btn_browse_soffice = ToolButton(FIF.FOLDER)
        self.btn_browse_soffice.clicked.connect(self.browse_soffice)
        row3.addWidget(BodyLabel("LibreOffice:"))
        row3.addWidget(self.edit_lo_path)
        row3.addWidget(self.btn_browse_soffice)
        layout.addLayout(row3)

        row4 = QHBoxLayout()
        self.edit_prefix = LineEdit()
        self.edit_prefix.setPlaceholderText("Prefix...")
        self.edit_prefix.setFixedWidth(200)
        self.edit_suffix = LineEdit()
        self.edit_suffix.setPlaceholderText("Suffix...")
        self.edit_suffix.setFixedWidth(200)
        row4.addWidget(BodyLabel("Prefix:"))
        row4.addWidget(self.edit_prefix)
        row4.addWidget(BodyLabel("Suffix:"))
        row4.addWidget(self.edit_suffix)
        row4.addStretch()
        layout.addLayout(row4)

        self.scroll_layout.addWidget(card)

    def _build_actions_card(self):
        card = AccordionCard("Actions", FIF.PLAY)
        layout = card.content_layout

        btn_row = QHBoxLayout()
        self.btn_generate = PrimaryPushButton(FIF.PLAY, "Generate")
        self.btn_generate.setFixedHeight(40)
        self.btn_generate.clicked.connect(self.on_run_clicked)
        self.btn_open_output = PushButton(FIF.FOLDER, "Open Output Folder")
        self.btn_open_output.setFixedHeight(40)
        self.btn_open_output.clicked.connect(self.open_output_folder)
        self.btn_open_output.setVisible(False)
        self.btn_debug = PushButton("Debug Excel")
        self.btn_debug.setFixedHeight(40)
        self.btn_debug.clicked.connect(self.debug_excel)
        btn_row.addWidget(self.btn_generate)
        btn_row.addWidget(self.btn_open_output)
        btn_row.addWidget(self.btn_debug)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        self.progress_bar = ProgressBar()
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.txt_log = TextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setMinimumHeight(200)
        layout.addWidget(self.txt_log)

        self.scroll_layout.addWidget(card)
    def log(self, msg):
        self.txt_log.append(msg)

    def set_progress(self, total, current):
        if total <= 0:
            self.progress_bar.setValue(0)
        else:
            self.progress_bar.setValue(int(100 * current / total))

    def set_busy(self, busy):
        self.is_running = busy
        self.btn_generate.setEnabled(not busy)
        self.btn_generate.setText("Generating..." if busy else "Generate")
        self.btn_save_ini.setEnabled(not busy)
        self.btn_load_ini.setEnabled(not busy)
        if not busy:
            self.progress_bar.setValue(0)

    def browse_source(self):
        f, _ = QFileDialog.getOpenFileName(
            self, "Pilih source Excel",
            "", "Excel files (*.xlsx *.xls *.xlsm *.xlsb)"
        )
        if f:
            self.edit_source.setText(f)
            self.log(f"Source: {f}")

    def browse_template(self):
        f, _ = QFileDialog.getOpenFileName(
            self, "Pilih template Excel",
            "", "Excel files (*.xlsx)"
        )
        if f:
            self.edit_template.setText(f)
            self.log(f"Template: {f}")

    def browse_outdir(self):
        d = QFileDialog.getExistingDirectory(self, "Pilih output folder")
        if d:
            self.edit_outdir.setText(d)
            self.log(f"Output: {d}")

    def browse_soffice(self):
        f, _ = QFileDialog.getOpenFileName(
            self, "Pilih soffice.exe (LibreOffice)",
            "", "Executable (soffice.exe);;All files (*.*)"
        )
        if f:
            self.edit_lo_path.setText(f)
            self.log(f"LibreOffice: {f}")

    def load_sheets(self):
        src = self.edit_source.text().strip()
        if not src:
            InfoBar.warning("Perhatian", "Pilih source Excel dulu.", parent=self, duration=3000, position=InfoBarPosition.TOP)
            return
        try:
            xls = pd.ExcelFile(src)
            sheets = xls.sheet_names
            self.cmb_sheet.clear()
            self.cmb_sheet.addItems(sheets)
            if sheets:
                self.cmb_sheet.setCurrentIndex(0)
            self.log(f"Sheets loaded: {', '.join(sheets)}")
        except Exception as e:
            InfoBar.error("Error", str(e), parent=self, duration=5000, position=InfoBarPosition.TOP)

    def load_headers(self):
        src = self.edit_source.text().strip()
        sheet = self.cmb_sheet.currentText().strip()
        if not src or not sheet:
            InfoBar.warning("Perhatian", "Pastikan source & sheet sudah dipilih.", parent=self, duration=3000, position=InfoBarPosition.TOP)
            return
        try:
            header_row_idx = self.spin_header_rows.value() - 1
            df = pd.read_excel(src, sheet_name=sheet, header=header_row_idx, nrows=0)
            headers = list(df.columns.astype(str))
            index_vals = [str(i+1) for i in range(len(headers))]
            values = headers + index_vals
            self.cmb_key.clear()
            self.cmb_key.addItems(values)
            if headers:
                self.cmb_key.setCurrentIndex(0)
            self.log(f"Headers loaded: {headers}")
        except Exception as e:
            InfoBar.error("Error", str(e), parent=self, duration=5000, position=InfoBarPosition.TOP)

    def debug_excel(self):
        try:
            self.log("=== Excel Detection Debug ===")
            results = debug_excel_detection()
            for result in results:
                self.log(result)
            excel_available = check_excel_availability()
            self.log(f"Final check_excel_availability(): {excel_available}")
            self.log("=== Debug selesai ===")
            if excel_available:
                InfoBar.success("Debug Excel", "Excel terdeteksi dan dapat diakses!", parent=self, duration=3000, position=InfoBarPosition.TOP)
            else:
                InfoBar.warning("Debug Excel", "Excel tidak dapat diakses. Lihat log.", parent=self, duration=5000, position=InfoBarPosition.TOP)
        except Exception as e:
            self.log(f"Error saat debug: {str(e)}")
            InfoBar.error("Debug Error", f"Gagal debug: {str(e)}", parent=self, duration=5000, position=InfoBarPosition.TOP)

    def open_output_folder(self):
        out_dir = self.edit_outdir.text().strip()
        if out_dir and Path(out_dir).exists():
            try:
                win_path = str(Path(out_dir).resolve())
                subprocess.run(f'explorer "{win_path}"', shell=True)
            except Exception as e:
                InfoBar.error("Error", f"Failed to open folder: {str(e)}", parent=self, duration=5000, position=InfoBarPosition.TOP)
        else:
            InfoBar.warning("Warning", "Output folder not set or doesn't exist", parent=self, duration=3000, position=InfoBarPosition.TOP)
    def on_run_clicked(self):
        if self.is_running:
            return
        try:
            source_path = Path(self.edit_source.text().strip())
            template_path = Path(self.edit_template.text().strip())
            out_dir = Path(self.edit_outdir.text().strip())
            sheet_name = self.cmb_sheet.currentText().strip()
            key_raw = self.cmb_key.currentText().strip()
            header_rows = self.spin_header_rows.value()
            pdf_engine = self.cmb_pdf_engine.currentText().strip().lower()

            if not source_path.exists():
                InfoBar.error("Error", "Source Excel tidak ditemukan.", parent=self, duration=5000, position=InfoBarPosition.TOP)
                return
            if not template_path.exists():
                InfoBar.error("Error", "Template Excel tidak ditemukan.", parent=self, duration=5000, position=InfoBarPosition.TOP)
                return
            if not self.edit_outdir.text().strip():
                InfoBar.error("Error", "Output folder belum dipilih.", parent=self, duration=5000, position=InfoBarPosition.TOP)
                return
            if not sheet_name:
                InfoBar.error("Error", "Sheet belum dipilih.", parent=self, duration=5000, position=InfoBarPosition.TOP)
                return
            if not key_raw:
                InfoBar.error("Error", "Key Column belum dipilih/diisi.", parent=self, duration=5000, position=InfoBarPosition.TOP)
                return

            try:
                key_col = int(key_raw)
            except ValueError:
                key_col = key_raw

            if pdf_engine == "xlwings":
                if not XLWINGS_AVAILABLE:
                    InfoBar.warning("xlwings", "xlwings belum terpasang. Gunakan 'libreoffice' atau 'none'.", parent=self, duration=5000, position=InfoBarPosition.TOP)
                    return
                elif not check_excel_availability():
                    InfoBar.warning("Excel", "Microsoft Excel tidak dapat diakses via COM.", parent=self, duration=5000, position=InfoBarPosition.TOP)
                    return

            soffice_path = None
            if pdf_engine == "libreoffice":
                lo_explicit = self.edit_lo_path.text().strip()
                soffice_path = find_soffice(lo_explicit)
                if not soffice_path:
                    InfoBar.error("Error", "LibreOffice (soffice.exe) tidak ditemukan.", parent=self, duration=5000, position=InfoBarPosition.TOP)
                    return

            self.set_busy(True)
            self.log("Mulai generate...")

            if pdf_engine == "xlwings":
                self.log("Membersihkan Excel COM sessions...")
                cleanup_excel_com()

            params = {
                'source_path': source_path,
                'sheet_name': sheet_name,
                'key_col': key_col,
                'template_path': template_path,
                'out_dir': out_dir,
                'header_rows': header_rows,
                'pdf_engine': pdf_engine,
                'soffice_path': soffice_path,
                'prefix': self.edit_prefix.text().strip(),
                'suffix': self.edit_suffix.text().strip(),
            }

            self.worker = SplitWorker(params)
            self.worker.status.connect(self.log)
            self.worker.progress.connect(self.set_progress)
            self.worker.finished.connect(self._on_worker_finished)
            self.worker.error.connect(self._on_worker_error)
            self.worker.start()

        except Exception as e:
            InfoBar.error("Error", str(e), parent=self, duration=5000, position=InfoBarPosition.TOP)
            self.set_busy(False)

    def _on_worker_finished(self):
        self.set_busy(False)
        self.log("Selesai.")
        self.btn_open_output.setVisible(True)
        pdf_engine = self.cmb_pdf_engine.currentText().strip().lower()
        if pdf_engine == "xlwings":
            cleanup_excel_com()
        InfoBar.success("Selesai", "Proses selesai.", parent=self, duration=5000, position=InfoBarPosition.TOP)

    def _on_worker_error(self, error_msg):
        self.set_busy(False)
        self.log(f"Error: {error_msg}")
        InfoBar.error("Error", error_msg, parent=self, duration=8000, position=InfoBarPosition.TOP)

    def save_ini(self):
        f, _ = QFileDialog.getSaveFileName(
            self, "Simpan konfigurasi", "", "INI files (*.ini)"
        )
        if not f:
            return
        cfg = configparser.ConfigParser()
        cfg["template"] = {
            "template_path": self.edit_template.text().strip(),
            "header_rows": str(self.spin_header_rows.value()),
        }
        cfg["source"] = {
            "source_path": self.edit_source.text().strip(),
            "sheet_name": self.cmb_sheet.currentText().strip(),
            "key_col": self.cmb_key.currentText().strip()
        }
        cfg["output"] = {
            "output_dir": self.edit_outdir.text().strip(),
            "pdf_engine": self.cmb_pdf_engine.currentText().strip().lower(),
            "libreoffice_path": self.edit_lo_path.text().strip(),
            "prefix": self.edit_prefix.text().strip(),
            "suffix": self.edit_suffix.text().strip()
        }
        with open(f, "w", encoding="utf-8") as fp:
            cfg.write(fp)
        self.log(f"Konfigurasi tersimpan: {f}")

    def load_ini(self):
        f, _ = QFileDialog.getOpenFileName(
            self, "Muat konfigurasi", "", "INI files (*.ini)"
        )
        if not f:
            return
        cfg = configparser.ConfigParser()
        cfg.read(f, encoding="utf-8")
        try:
            self.edit_template.setText(cfg.get("template", "template_path", fallback=""))
            self.spin_header_rows.setValue(cfg.getint("template", "header_rows", fallback=5))
            self.edit_source.setText(cfg.get("source", "source_path", fallback=""))
            sheet = cfg.get("source", "sheet_name", fallback="")
            if sheet:
                self.cmb_sheet.clear()
                self.cmb_sheet.addItem(sheet)
                self.cmb_sheet.setCurrentIndex(0)
            key = cfg.get("source", "key_col", fallback="")
            if key:
                self.cmb_key.clear()
                self.cmb_key.addItem(key)
                self.cmb_key.setCurrentIndex(0)
            self.edit_outdir.setText(cfg.get("output", "output_dir", fallback=""))
            pdf_eng = cfg.get("output", "pdf_engine", fallback="xlwings").lower()
            idx = self.cmb_pdf_engine.findText(pdf_eng)
            if idx >= 0:
                self.cmb_pdf_engine.setCurrentIndex(idx)
            self.edit_lo_path.setText(cfg.get("output", "libreoffice_path", fallback=""))
            self.edit_prefix.setText(cfg.get("output", "prefix", fallback=""))
            self.edit_suffix.setText(cfg.get("output", "suffix", fallback=""))
            self.lbl_loaded_ini.setText(f"Loaded: {Path(f).name}")
            self.log(f"Konfigurasi dimuat: {f}")
        except Exception as e:
            InfoBar.error("Error", f"Format .ini tidak valid: {e}", parent=self, duration=5000, position=InfoBarPosition.TOP)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SplitApp()
    window.show()
    sys.exit(app.exec())