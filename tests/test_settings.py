import os
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path
import tempfile
import unittest


os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
if os.name == "nt" and os.path.isdir(r"C:\Windows\Fonts"):
    os.environ.setdefault("QT_QPA_FONTDIR", r"C:\Windows\Fonts")

try:
    from PySide6.QtCore import QSettings
    from PySide6.QtWidgets import QApplication
    with redirect_stdout(StringIO()):
        import main
except ModuleNotFoundError as exc:
    raise unittest.SkipTest(f"GUI dependencies are not installed: {exc.name}")


class SettingsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def make_settings(self, path: Path):
        settings = QSettings(str(path), QSettings.IniFormat)
        settings.clear()
        return settings

    def test_split_app_persists_fields_with_qsettings(self):
        with tempfile.TemporaryDirectory() as tmp:
            settings_path = Path(tmp) / "settings.ini"
            first = main.SplitApp(settings=self.make_settings(settings_path))
            self.addCleanup(first.deleteLater)

            first.edit_source.setText("source.xlsx")
            first.edit_outdir.setText("out")
            first.edit_prefix.setText("PRE")
            first.edit_suffix.setText("SUF")
            first.spin_source_header_rows.setValue(2)
            first.spin_template_header_rows.setValue(3)
            first.cmb_output_type.setCurrentIndex(first.cmb_output_type.findText("PDF"))
            first.cmb_template_mode.setCurrentIndex(
                first.cmb_template_mode.findText("Use Source as Template")
            )
            first.cmb_pdf_engine.setCurrentIndex(first.cmb_pdf_engine.findText("libreoffice"))
            first.chk_verbose_logging.setChecked(True)
            first.source_headers = ["Name"]
            first.template_headers = ["Worker"]
            first.render_mapping_rows({"Worker": "Name"})
            first.save_settings()

            second = main.SplitApp(settings=QSettings(str(settings_path), QSettings.IniFormat))
            self.addCleanup(second.deleteLater)

            self.assertEqual(second.edit_source.text(), "source.xlsx")
            self.assertEqual(second.edit_outdir.text(), "out")
            self.assertEqual(second.edit_prefix.text(), "PRE")
            self.assertEqual(second.edit_suffix.text(), "SUF")
            self.assertEqual(second.spin_source_header_rows.value(), 2)
            self.assertEqual(second.spin_template_header_rows.value(), 3)
            self.assertEqual(second.cmb_output_type.currentText(), "PDF")
            self.assertEqual(second.cmb_template_mode.currentText(), "Use Source as Template")
            self.assertEqual(second.cmb_pdf_engine.currentText(), "libreoffice")
            self.assertTrue(second.chk_verbose_logging.isChecked())
            self.assertEqual(second.saved_column_mapping, {"Worker": "Name"})

    def test_ini_toolbar_buttons_are_replaced_by_reset_settings(self):
        with tempfile.TemporaryDirectory() as tmp:
            window = main.SplitApp(settings=self.make_settings(Path(tmp) / "settings.ini"))
            self.addCleanup(window.deleteLater)

            self.assertFalse(hasattr(window, "btn_save_ini"))
            self.assertFalse(hasattr(window, "btn_load_ini"))
            self.assertTrue(hasattr(window, "btn_reset_settings"))

    def test_mail_merge_settings_persist_with_qsettings(self):
        with tempfile.TemporaryDirectory() as tmp:
            settings_path = Path(tmp) / "settings.ini"
            first = main.SplitApp(settings=self.make_settings(settings_path))
            self.addCleanup(first.deleteLater)

            first.edit_recipient_path.setText("recipients.xlsx")
            first.cmb_recipient_sheet.addItem("Recipients")
            first.cmb_recipient_sheet.setCurrentIndex(0)
            first.spin_recipient_header_row.setValue(2)
            first.edit_mail_subject.setText("Subject {key}")
            first.edit_mail_html_template.setText("body.html")
            first.chk_attach_excel.setChecked(False)
            first.chk_attach_pdf.setChecked(True)
            first.chk_delay_delivery.setChecked(True)
            first.spin_delay_minutes.setValue(9)
            first.chk_throttle.setChecked(False)
            first.spin_throttle_seconds.setValue(4)
            first.save_settings()

            second = main.SplitApp(settings=QSettings(str(settings_path), QSettings.IniFormat))
            self.addCleanup(second.deleteLater)

            self.assertEqual(second.edit_recipient_path.text(), "recipients.xlsx")
            self.assertEqual(second.cmb_recipient_sheet.currentText(), "Recipients")
            self.assertEqual(second.spin_recipient_header_row.value(), 2)
            self.assertEqual(second.edit_mail_subject.text(), "Subject {key}")
            self.assertEqual(second.edit_mail_html_template.text(), "body.html")
            self.assertFalse(second.chk_attach_excel.isChecked())
            self.assertTrue(second.chk_attach_pdf.isChecked())
            self.assertTrue(second.chk_delay_delivery.isChecked())
            self.assertEqual(second.spin_delay_minutes.value(), 9)
            self.assertFalse(second.chk_throttle.isChecked())
            self.assertEqual(second.spin_throttle_seconds.value(), 4)

    def test_mail_split_folder_settings_persist_with_qsettings(self):
        with tempfile.TemporaryDirectory() as tmp:
            settings_path = Path(tmp) / "settings.ini"
            first = main.SplitApp(settings=self.make_settings(settings_path))
            self.addCleanup(first.deleteLater)

            first.edit_split_folder.setText("C:/out")
            first.edit_detect_prefix.setText("Report")
            first.edit_detect_suffix.setText("Final")
            first.save_settings()

            second = main.SplitApp(settings=QSettings(str(settings_path), QSettings.IniFormat))
            self.addCleanup(second.deleteLater)

            self.assertEqual(second.edit_split_folder.text(), "C:/out")
            self.assertEqual(second.edit_detect_prefix.text(), "Report")
            self.assertEqual(second.edit_detect_suffix.text(), "Final")


if __name__ == "__main__":
    unittest.main()
