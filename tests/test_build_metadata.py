from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read_text(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class BuildMetadataTests(unittest.TestCase):
    def test_build_cmd_uses_requirements_for_pip_install(self):
        text = read_text("build.cmd").lower()

        self.assertRegex(text, r"pip[^\r\n]*install[^\r\n]*-r[^\r\n]*requirements\.txt")
        self.assertNotIn("customtkinter", text)

    def test_uv_build_paths_install_requirements_instead_of_syncing(self):
        for script in ("build.cmd", "build-uv.bat"):
            with self.subTest(script=script):
                text = read_text(script).lower()

                self.assertRegex(text, r"uv pip install[^\r\n]*-r[^\r\n]*requirements\.txt")
                self.assertNotIn("uv pip sync", text)

        readme = read_text("README.md").lower()
        self.assertIn("uv pip install --upgrade -r requirements.txt", readme)
        self.assertNotIn("uv pip sync", readme)

    def test_pyinstaller_spec_collects_current_ui_toolkit(self):
        text = read_text("main.spec")
        normalized = text.lower()

        self.assertNotIn("customtkinter", normalized)
        self.assertIn("PySide6", text)
        self.assertRegex(text, r"collect_data_files\([\"']qfluentwidgets[\"']")
        self.assertRegex(text, r"collect_submodules\([\"']qfluentwidgets[\"']")

    def test_readme_lists_pyside6_ui_dependencies(self):
        text = read_text("README.md")
        normalized = text.lower()

        self.assertNotIn("customtkinter", normalized)
        self.assertIn("PySide6", text)
        self.assertIn("PySide6-Fluent-Widgets", text)

    def test_build_scripts_report_onefile_output_path(self):
        for script in ("build.cmd", "build-uv.bat"):
            with self.subTest(script=script):
                text = read_text(script).lower()

                self.assertIn(r"dist\%project_name%.exe", text)
                self.assertNotIn(r"dist\%project_name%\%project_name%.exe", text)

    def test_batch_echo_lines_avoid_parentheses(self):
        for script in ("build.cmd", "build-uv.bat"):
            for line_number, line in enumerate(read_text(script).splitlines(), start=1):
                stripped = line.strip().lower()
                if not stripped.startswith("echo "):
                    continue

                with self.subTest(script=script, line=line_number):
                    self.assertNotRegex(stripped, r"[()]")

    def test_spec_documents_onefile_output_path(self):
        text = read_text("main.spec")

        self.assertIn("dist/ExcelSplitter.exe", text)
        self.assertNotIn("dist/ExcelSplitter/ExcelSplitter.exe", text)

    def test_pyinstaller_spec_embeds_and_bundles_app_icon(self):
        text = read_text("main.spec")

        self.assertIn('ICON_PATH    = "excel-split.ico"', text)
        self.assertIn('("excel-split.ico", ".")', text)
        self.assertIn("icon=ICON_PATH", text)

    def test_spec_does_not_include_invalid_hidden_imports(self):
        text = read_text("main.spec")

        for module in ("win32com.gen_py", "pkg_resources.py2_warn", "pkg_resources.markers"):
            with self.subTest(module=module):
                self.assertNotIn(module, text)

    def test_pdf_metadata_matches_supported_engines(self):
        for path in ("README.md", "requirements.txt", "main.spec", "build.cmd", "build-uv.bat"):
            with self.subTest(path=path):
                self.assertNotIn("reportlab", read_text(path).lower())

        readme = read_text("README.md").lower()
        for engine in ("xlwings", "libreoffice", "none"):
            self.assertIn(engine, readme)


if __name__ == "__main__":
    unittest.main()
