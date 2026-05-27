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
    from qfluentwidgets import Theme, isDarkTheme, qconfig
    with redirect_stdout(StringIO()):
        import main
except ModuleNotFoundError as exc:
    raise unittest.SkipTest(f"GUI dependencies are not installed: {exc.name}")


class UISmokeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def make_settings(self, path: Path):
        settings = QSettings(str(path), QSettings.IniFormat)
        settings.clear()
        return settings

    def test_split_app_constructs_with_installed_widgets(self):
        with tempfile.TemporaryDirectory() as tmp:
            window = main.SplitApp(settings=self.make_settings(Path(tmp) / "settings.ini"))
            self.addCleanup(window.deleteLater)

            self.assertEqual(window.windowTitle(), "Excel Splitter")
            self.assertTrue(hasattr(window, "cmb_template_mode"))
            self.assertTrue(hasattr(window, "mapping_card"))
            self.assertTrue(hasattr(window, "btn_auto_map"))

            modes = [
                window.cmb_template_mode.itemText(index)
                for index in range(window.cmb_template_mode.count())
            ]
            self.assertEqual(modes, ["Use Template File", "Use Source as Template"])

    def test_split_app_uses_compact_dashboard_layout(self):
        with tempfile.TemporaryDirectory() as tmp:
            window = main.SplitApp(settings=self.make_settings(Path(tmp) / "settings.ini"))
            self.addCleanup(window.deleteLater)

            self.assertTrue(hasattr(window, "workflow_rail"))
            self.assertTrue(hasattr(window, "main_panel_layout"))
            self.assertTrue(hasattr(window, "footer_bar"))

            self.assertLessEqual(window.edit_source.maximumWidth(), 520)
            self.assertLessEqual(window.edit_template.maximumWidth(), 520)
            self.assertLessEqual(window.edit_outdir.maximumWidth(), 520)
            self.assertLessEqual(window.edit_lo_path.maximumWidth(), 360)
            self.assertLessEqual(window.edit_prefix.maximumWidth(), 180)
            self.assertLessEqual(window.edit_suffix.maximumWidth(), 180)

    def test_template_mode_hides_template_file_controls_for_source_template(self):
        with tempfile.TemporaryDirectory() as tmp:
            window = main.SplitApp(settings=self.make_settings(Path(tmp) / "settings.ini"))
            self.addCleanup(window.deleteLater)

            self.assertFalse(window.edit_template.isHidden())
            self.assertFalse(window.btn_browse_template.isHidden())

            window.cmb_template_mode.setCurrentIndex(
                window.cmb_template_mode.findText("Use Source as Template")
            )
            window.on_template_mode_changed()

            self.assertTrue(window.edit_template.isHidden())
            self.assertTrue(window.btn_browse_template.isHidden())
            self.assertTrue(window.mapping_card.isHidden())

    def test_mapping_status_updates_when_user_selects_source_column(self):
        with tempfile.TemporaryDirectory() as tmp:
            window = main.SplitApp(settings=self.make_settings(Path(tmp) / "settings.ini"))
            self.addCleanup(window.deleteLater)

            window.source_headers = ["Name"]
            window.template_headers = ["Worker"]
            window.render_mapping_rows({})

            self.assertTrue(hasattr(window, "mapping_status_labels"))
            self.assertEqual(window.mapping_status_labels["Worker"].text(), "Missing")

            window.mapping_combos["Worker"].setCurrentIndex(
                window.mapping_combos["Worker"].findText("Name")
            )

            self.assertEqual(window.mapping_status_labels["Worker"].text(), "Mapped")

    def test_dashboard_uses_native_theme_mode_with_matching_surfaces(self):
        with tempfile.TemporaryDirectory() as tmp:
            window = main.SplitApp(settings=self.make_settings(Path(tmp) / "settings.ini"))
            self.addCleanup(window.deleteLater)

            self.assertEqual(window.objectName(), "appRoot")
            self.assertEqual(window.workflow_rail.objectName(), "workflowRail")
            self.assertEqual(qconfig.themeMode.value, Theme.AUTO)
            if isDarkTheme():
                self.assertIn("#202020", window.styleSheet())
                self.assertNotIn("#f5f7fb", window.styleSheet())
            else:
                self.assertIn("#f5f7fb", window.styleSheet())
                self.assertNotIn("#202020", window.styleSheet())

    def test_field_action_buttons_share_input_height_contract(self):
        with tempfile.TemporaryDirectory() as tmp:
            window = main.SplitApp(settings=self.make_settings(Path(tmp) / "settings.ini"))
            self.addCleanup(window.deleteLater)

            self.assertTrue(hasattr(window, "field_action_buttons"))
            self.assertGreaterEqual(len(window.field_action_buttons), 6)
            for button in window.field_action_buttons:
                self.assertEqual(button.minimumHeight(), main.FIELD_CONTROL_HEIGHT)
                self.assertEqual(button.maximumHeight(), main.FIELD_CONTROL_HEIGHT)

    def test_libreoffice_path_row_only_shows_for_libreoffice_pdf_engine(self):
        with tempfile.TemporaryDirectory() as tmp:
            window = main.SplitApp(settings=self.make_settings(Path(tmp) / "settings.ini"))
            self.addCleanup(window.deleteLater)

            window.cmb_output_type.setCurrentIndex(window.cmb_output_type.findText("PDF"))
            window.on_output_type_changed()

            self.assertTrue(hasattr(window, "lo_path_row_widget"))
            self.assertTrue(window.lo_path_row_widget.isHidden())

            window.cmb_pdf_engine.setCurrentIndex(window.cmb_pdf_engine.findText("libreoffice"))
            window.on_pdf_engine_changed()

            self.assertFalse(window.lo_path_row_widget.isHidden())

    def test_header_rows_are_independent_and_output_type_controls_preview(self):
        with tempfile.TemporaryDirectory() as tmp:
            window = main.SplitApp(settings=self.make_settings(Path(tmp) / "settings.ini"))
            self.addCleanup(window.deleteLater)

            self.assertTrue(hasattr(window, "spin_source_header_rows"))
            self.assertTrue(hasattr(window, "spin_template_header_rows"))
            self.assertTrue(hasattr(window, "btn_detect_source_header"))
            self.assertTrue(hasattr(window, "btn_detect_template_header"))
            self.assertTrue(hasattr(window, "cmb_output_type"))
            self.assertTrue(hasattr(window, "lbl_filename_preview"))

            output_types = [
                window.cmb_output_type.itemText(index)
                for index in range(window.cmb_output_type.count())
            ]
            self.assertEqual(output_types, ["Excel", "PDF"])
            self.assertEqual(window.cmb_output_type.currentText(), "Excel")
            self.assertTrue(window.cmb_pdf_engine.isHidden())

            window.cmb_key.clear()
            window.cmb_key.addItem("Dept")
            window.edit_prefix.setText("PRE")
            window.edit_suffix.setText("SUF")
            window.update_filename_preview()
            self.assertIn("PRE <Dept value> SUF.xlsx", window.lbl_filename_preview.text())

            window.cmb_output_type.setCurrentIndex(window.cmb_output_type.findText("PDF"))
            window.on_output_type_changed()
            self.assertFalse(window.cmb_pdf_engine.isHidden())
            self.assertIn("PRE <Dept value> SUF.pdf", window.lbl_filename_preview.text())

    def test_footer_progress_is_hidden_until_generation_starts(self):
        with tempfile.TemporaryDirectory() as tmp:
            window = main.SplitApp(settings=self.make_settings(Path(tmp) / "settings.ini"))
            self.addCleanup(window.deleteLater)

            self.assertTrue(window.progress_bar.isHidden())
            window.set_busy(True)
            self.assertFalse(window.progress_bar.isHidden())


if __name__ == "__main__":
    unittest.main()
