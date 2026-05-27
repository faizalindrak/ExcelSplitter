from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook, load_workbook


with redirect_stdout(StringIO()):
    import main


class HeaderMappingTests(unittest.TestCase):
    def test_normalize_header_ignores_case_spaces_and_punctuation(self):
        self.assertEqual(main.normalize_header(" Employee  Name!"), "employeename")

    def test_auto_map_columns_matches_normalized_headers(self):
        mapping = main.auto_map_columns(
            ["Employee Name", "Dept"],
            ["employee_name", "Dept", "Amount"],
        )

        self.assertEqual(mapping, {"Employee Name": "employee_name", "Dept": "Dept"})

    def test_validate_column_mapping_reports_unmapped_template_columns(self):
        missing = main.validate_column_mapping(
            ["Employee Name", "Dept"],
            {"Employee Name": "Name", "Dept": None},
        )

        self.assertEqual(missing, ["Dept"])


class TemplateFileSplitTests(unittest.TestCase):
    def make_source_workbook(self, path: Path):
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["Name", "Dept", "Amount"])
        ws.append(["Alice", "A", 10])
        ws.append(["Bob", "B", 20])
        ws.append(["Ana", "A", 30])
        wb.save(path)

    def make_template_workbook(self, path: Path):
        wb = Workbook()
        ws = wb.active
        ws.title = "Template"
        ws.append(["Worker", "Team"])
        wb.save(path)

    def test_template_file_mode_uses_manual_mapping_for_different_headers(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "source.xlsx"
            template = tmp_path / "template.xlsx"
            out_dir = tmp_path / "out"
            self.make_source_workbook(source)
            self.make_template_workbook(template)

            main.split_excel_with_template(
                source,
                "Data",
                "Dept",
                template,
                out_dir,
                1,
                pdf_engine="none",
                template_mode="template_file",
                column_mapping={"Worker": "Name", "Team": "Dept"},
            )

            wb = load_workbook(out_dir / "A.xlsx", data_only=True)
            ws = wb.active
            self.assertEqual(ws["A1"].value, "Worker")
            self.assertEqual(ws["B1"].value, "Team")
            self.assertEqual(ws["A2"].value, "Alice")
            self.assertEqual(ws["B2"].value, "A")
            self.assertEqual(ws["A3"].value, "Ana")
            self.assertEqual(ws["B3"].value, "A")

    def test_template_file_mode_requires_complete_manual_mapping(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "source.xlsx"
            template = tmp_path / "template.xlsx"
            out_dir = tmp_path / "out"
            self.make_source_workbook(source)
            self.make_template_workbook(template)

            with self.assertRaisesRegex(ValueError, "Team"):
                main.split_excel_with_template(
                    source,
                    "Data",
                    "Dept",
                    template,
                    out_dir,
                    1,
                    pdf_engine="none",
                    template_mode="template_file",
                    column_mapping={"Worker": "Name"},
                )

    def test_template_file_mode_requires_template_headers_for_mapping(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "source.xlsx"
            template = tmp_path / "template.xlsx"
            out_dir = tmp_path / "out"
            self.make_source_workbook(source)

            wb = Workbook()
            ws = wb.active
            ws.title = "Template"
            ws["A2"] = "data row style only"
            wb.save(template)

            with self.assertRaisesRegex(ValueError, "Header template"):
                main.split_excel_with_template(
                    source,
                    "Data",
                    "Dept",
                    template,
                    out_dir,
                    1,
                    pdf_engine="none",
                    template_mode="template_file",
                    column_mapping={"Worker": "Name"},
                )


class SourceTemplateSplitTests(unittest.TestCase):
    def test_source_template_mode_keeps_selected_sheet_and_matching_key_rows(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "source.xlsx"
            out_dir = tmp_path / "out"

            wb = Workbook()
            ws = wb.active
            ws.title = "Data"
            ws.column_dimensions["A"].width = 22
            ws.append(["Dept", "Name", "Amount"])
            ws.append(["A", "Alice", 10])
            ws.append(["B", "Bob", 20])
            ws.append(["A", "Ana", 30])
            other = wb.create_sheet("Other")
            other["A1"] = "Should not be copied"
            wb.save(source)

            main.split_excel_with_template(
                source,
                "Data",
                "Dept",
                source,
                out_dir,
                1,
                pdf_engine="none",
                template_mode="source_template",
            )

            output = load_workbook(out_dir / "A.xlsx", data_only=True)
            self.assertEqual(output.sheetnames, ["Data"])
            ws_out = output["Data"]
            self.assertEqual(ws_out.column_dimensions["A"].width, 22)
            self.assertEqual([cell.value for cell in ws_out[1]], ["Dept", "Name", "Amount"])
            self.assertEqual(ws_out.max_row, 3)
            self.assertEqual([ws_out["A2"].value, ws_out["B2"].value], ["A", "Alice"])
            self.assertEqual([ws_out["A3"].value, ws_out["B3"].value], ["A", "Ana"])


if __name__ == "__main__":
    unittest.main()
