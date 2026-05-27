from pathlib import Path
import tempfile
import unittest
from datetime import datetime

from openpyxl import Workbook

import mail_merge


class MailMergeCoreTests(unittest.TestCase):
    def test_parse_email_list_splits_semicolon_values(self):
        self.assertEqual(
            mail_merge.parse_email_list("a@example.com; b@example.com ;"),
            ["a@example.com", "b@example.com"],
        )

    def test_parse_email_list_handles_empty_values(self):
        self.assertEqual(mail_merge.parse_email_list(" ; "), [])
        self.assertEqual(mail_merge.parse_email_list(None), [])

    def test_is_valid_email_rejects_obvious_invalid_values(self):
        self.assertTrue(mail_merge.is_valid_email("person@example.com"))
        self.assertFalse(mail_merge.is_valid_email("missing-at.example.com"))
        self.assertFalse(mail_merge.is_valid_email("person@"))
        self.assertFalse(mail_merge.is_valid_email("@example.com"))

    def test_render_placeholders_uses_case_insensitive_keys_and_leaves_unknown(self):
        context = {
            "key": "A",
            "To": "alice@example.com",
            "Dept": "Finance",
        }

        rendered = mail_merge.render_placeholders(
            "Hello {to}, key={KEY}, dept={dept}, keep={missing}",
            context,
        )

        self.assertEqual(
            rendered,
            "Hello alice@example.com, key=A, dept=Finance, keep={missing}",
        )

    def test_read_recipient_headers_from_excel_sheet(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "recipients.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Recipients"
            ws.append(["Report", None, None, None])
            ws.append(["Key", "To", "CC", "Department"])
            wb.save(path)

            headers = mail_merge.read_recipient_headers(path, "Recipients", 2)

            self.assertEqual(headers, ["Key", "To", "CC", "Department"])

    def test_load_recipient_rows_uses_manual_column_mapping(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "recipients.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Recipients"
            ws.append(["Customer", "Email", "Copy", "Team"])
            ws.append(["A", "a@example.com; aa@example.com", "lead@example.com", "Finance"])
            ws.append(["B", "b@example.com", None, "Ops"])
            wb.save(path)

            rows = mail_merge.load_recipient_rows(
                path,
                "Recipients",
                1,
                {
                    "key": "Customer",
                    "to": "Email",
                    "cc": "Copy",
                    "bcc": "",
                },
            )

            self.assertEqual(len(rows), 2)
            self.assertEqual(rows[0].key, "A")
            self.assertEqual(rows[0].to, ["a@example.com", "aa@example.com"])
            self.assertEqual(rows[0].cc, ["lead@example.com"])
            self.assertEqual(rows[0].raw["Team"], "Finance")
            self.assertEqual(rows[1].bcc, [])


if __name__ == "__main__":
    unittest.main()
