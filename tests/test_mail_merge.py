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

    def test_build_email_jobs_renders_content_and_selected_attachments(self):
        with tempfile.TemporaryDirectory() as tmp:
            excel_path = Path(tmp) / "A.xlsx"
            pdf_path = Path(tmp) / "A.pdf"
            excel_path.write_text("excel", encoding="utf-8")
            pdf_path.write_text("pdf", encoding="utf-8")

            jobs, warnings = mail_merge.build_email_jobs(
                split_results=[
                    mail_merge.SplitResult("A", excel_path=excel_path, pdf_path=pdf_path, output_file_type="excel_and_pdf")
                ],
                recipients=[
                    mail_merge.RecipientRow(
                        key="A",
                        to=["a@example.com"],
                        cc=[],
                        bcc=[],
                        raw={"Department": "Finance"},
                    )
                ],
                template=mail_merge.EmailTemplate(
                    subject="Report {key} {Department}",
                    body="Hello {to}, file {excel_file}",
                    is_html=False,
                ),
                attachments=mail_merge.AttachmentSelection(attach_excel=True, attach_pdf=True),
            )

            self.assertEqual(warnings, [])
            self.assertEqual(len(jobs), 1)
            self.assertEqual(jobs[0].subject, "Report A Finance")
            self.assertEqual(jobs[0].body, f"Hello a@example.com, file {excel_path}")
            self.assertEqual(jobs[0].attachments, [excel_path, pdf_path])
            self.assertTrue(jobs[0].is_valid)

    def test_build_email_jobs_reports_missing_mapping_and_ignores_extra_mapping_rows(self):
        split_results = [mail_merge.SplitResult("A")]
        recipients = [
            mail_merge.RecipientRow(key="B", to=["b@example.com"], raw={}),
        ]

        jobs, warnings = mail_merge.build_email_jobs(
            split_results=split_results,
            recipients=recipients,
            template=mail_merge.EmailTemplate(subject="Subject", body="Body"),
            attachments=mail_merge.AttachmentSelection(attach_excel=False, attach_pdf=False),
        )

        self.assertEqual(len(jobs), 1)
        self.assertIn("No recipient mapping for key A", jobs[0].validation_errors)
        self.assertEqual(warnings, ["Recipient mapping key B does not match a generated split file"])

    def test_build_email_jobs_validates_email_subject_body_and_attachments(self):
        with tempfile.TemporaryDirectory() as tmp:
            missing_pdf = Path(tmp) / "A.pdf"
            jobs, warnings = mail_merge.build_email_jobs(
                split_results=[mail_merge.SplitResult("A", pdf_path=missing_pdf, output_file_type="pdf")],
                recipients=[mail_merge.RecipientRow(key="A", to=["bad-email"], raw={})],
                template=mail_merge.EmailTemplate(subject=" ", body=" "),
                attachments=mail_merge.AttachmentSelection(attach_excel=False, attach_pdf=True),
            )

            self.assertEqual(warnings, [])
            self.assertIn("Invalid To email: bad-email", jobs[0].validation_errors)
            self.assertIn("Selected PDF attachment is missing for key A", jobs[0].validation_errors)
            self.assertIn("Rendered subject is empty for key A", jobs[0].validation_errors)
            self.assertIn("Rendered body is empty for key A", jobs[0].validation_errors)


if __name__ == "__main__":
    unittest.main()
