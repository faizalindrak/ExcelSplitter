import os
from contextlib import redirect_stdout
from io import StringIO
import unittest


os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
if os.name == "nt" and os.path.isdir(r"C:\Windows\Fonts"):
    os.environ.setdefault("QT_QPA_FONTDIR", r"C:\Windows\Fonts")

try:
    from PySide6.QtWidgets import QApplication
    with redirect_stdout(StringIO()):
        import main
except ModuleNotFoundError as exc:
    raise unittest.SkipTest(f"GUI dependencies are not installed: {exc.name}")


class UISmokeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_split_app_constructs_with_installed_widgets(self):
        window = main.SplitApp()
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


if __name__ == "__main__":
    unittest.main()
