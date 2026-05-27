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


if __name__ == "__main__":
    unittest.main()
