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


if __name__ == "__main__":
    unittest.main()
