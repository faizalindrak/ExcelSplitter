# Mail Merge After Split Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a post-split Mail Merge workflow that loads recipient mappings from Excel, previews one email at a time, and sends through Outlook desktop with delay delivery and throttle controls.

**Architecture:** Add a focused `mail_merge.py` module for mail merge data models, recipient loading, template rendering, validation, job building, and Outlook provider logic. Keep `main.py` responsible for the PySide6 UI, split execution, and wiring generated split results into the Mail Merge panel. Use pure tests for mail-merge logic and UI smoke tests for widget state and carousel behavior.

**Tech Stack:** Python 3.13, PySide6, PySide6-Fluent-Widgets, pandas, openpyxl, pywin32/Outlook COM, unittest.

---

## File Structure

- Create `mail_merge.py`: pure mail merge models/helpers plus provider implementations. No PySide6 imports.
- Modify `main.py`: import mail merge helpers, return split manifests, show the post-split Mail Merge button, build Mail Merge UI, and run sending in a worker thread.
- Create `tests/test_mail_merge.py`: pure unit tests for recipient loading, placeholder rendering, validation, job building, send timing, fake provider, and Outlook provider through injected fakes.
- Modify `tests/test_templating_modes.py`: assert split generation returns `SplitResult` manifests for Excel, PDF, and Excel + PDF.
- Modify `tests/test_ui_smoke.py`: assert Mail Merge button/panel visibility, recipient mapping controls, carousel navigation, and send enablement.
- Modify `tests/test_settings.py`: assert Mail Merge reusable settings persist through `QSettings`.
- Modify `README.md`: document Mail Merge workflow, recipient mapping format, preview, Outlook sending, delay delivery, and throttle.

---

### Task 1: Mail Merge Models And Template Helpers

**Files:**
- Create: `mail_merge.py`
- Test: `tests/test_mail_merge.py`

- [ ] **Step 1: Write failing tests for models, email parsing, and placeholder rendering**

Add `tests/test_mail_merge.py`:

```python
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
rtk .venv\Scripts\python.exe -m unittest tests.test_mail_merge
```

Expected: fail with `ModuleNotFoundError: No module named 'mail_merge'`.

- [ ] **Step 3: Create `mail_merge.py` with data models and helpers**

Add `mail_merge.py`:

```python
import re
import time
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable, Iterable, Protocol

import pandas as pd


EMAIL_PATTERN = re.compile(r"^[^@\s;]+@[^@\s;]+\.[^@\s;]+$")


@dataclass(frozen=True)
class SplitResult:
    key: str
    excel_path: Path | None = None
    pdf_path: Path | None = None
    output_file_type: str = "excel"


@dataclass(frozen=True)
class RecipientRow:
    key: str
    to: list[str]
    cc: list[str] = field(default_factory=list)
    bcc: list[str] = field(default_factory=list)
    raw: dict[str, str] = field(default_factory=dict)


@dataclass(frozen=True)
class EmailTemplate:
    subject: str
    body: str
    is_html: bool = False
    template_path: Path | None = None


@dataclass(frozen=True)
class AttachmentSelection:
    attach_excel: bool = True
    attach_pdf: bool = False


@dataclass(frozen=True)
class SendTimingOptions:
    delay_delivery_enabled: bool = True
    delay_delivery_minutes: int = 5
    throttle_enabled: bool = True
    throttle_seconds: int = 5


@dataclass
class EmailJob:
    key: str
    to: list[str]
    cc: list[str]
    bcc: list[str]
    subject: str
    body: str
    is_html: bool
    attachments: list[Path]
    validation_errors: list[str] = field(default_factory=list)
    validation_warnings: list[str] = field(default_factory=list)

    @property
    def is_valid(self) -> bool:
        return not self.validation_errors


@dataclass(frozen=True)
class SendResult:
    key: str
    to: list[str]
    status: str
    message: str = ""


class MailProvider(Protocol):
    def send(self, job: EmailJob, timing: SendTimingOptions) -> SendResult:
        raise NotImplementedError


def parse_email_list(value) -> list[str]:
    if value is None:
        return []
    return [part.strip() for part in str(value).split(";") if part.strip()]


def is_valid_email(value: str) -> bool:
    return bool(EMAIL_PATTERN.match(value.strip()))


def render_placeholders(template: str, context: dict[str, object]) -> str:
    lookup = {str(key).lower(): "" if value is None else str(value) for key, value in context.items()}

    def replace(match: re.Match) -> str:
        name = match.group(1)
        return lookup.get(name.lower(), match.group(0))

    return re.sub(r"\{([^{}]+)\}", replace, template)
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```powershell
rtk .venv\Scripts\python.exe -m unittest tests.test_mail_merge
```

Expected: 4 tests pass.

- [ ] **Step 5: Commit**

Run:

```powershell
rtk git add mail_merge.py tests/test_mail_merge.py
rtk git commit -m "feat: add mail merge core models"
```

---

### Task 2: Recipient Mapping Excel Loader

**Files:**
- Modify: `mail_merge.py`
- Modify: `tests/test_mail_merge.py`

- [ ] **Step 1: Add failing tests for recipient workbook loading**

Append to `MailMergeCoreTests` in `tests/test_mail_merge.py`:

```python
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
rtk .venv\Scripts\python.exe -m unittest tests.test_mail_merge
```

Expected: fail with `AttributeError` for `read_recipient_headers`.

- [ ] **Step 3: Implement recipient loader**

Append to `mail_merge.py`:

```python
def _clean_cell(value) -> str:
    if value is None:
        return ""
    if pd.isna(value):
        return ""
    return str(value).strip()


def read_recipient_headers(path: Path, sheet_name: str, header_row: int) -> list[str]:
    df = pd.read_excel(path, sheet_name=sheet_name, header=header_row - 1, nrows=0, dtype=object)
    return [str(column).strip() for column in df.columns]


def load_recipient_rows(
    path: Path,
    sheet_name: str,
    header_row: int,
    column_mapping: dict[str, str],
) -> list[RecipientRow]:
    required = ["key", "to"]
    missing = [name for name in required if not column_mapping.get(name)]
    if missing:
        raise ValueError("Recipient mapping missing required columns: " + ", ".join(missing))

    df = pd.read_excel(path, sheet_name=sheet_name, header=header_row - 1, dtype=object)
    rows: list[RecipientRow] = []
    key_col = column_mapping["key"]
    to_col = column_mapping["to"]
    cc_col = column_mapping.get("cc") or ""
    bcc_col = column_mapping.get("bcc") or ""

    for _, record in df.fillna("").iterrows():
        raw = {str(column): _clean_cell(record[column]) for column in df.columns}
        key = _clean_cell(record[key_col])
        if not key:
            continue
        rows.append(
            RecipientRow(
                key=key,
                to=parse_email_list(record[to_col]),
                cc=parse_email_list(record[cc_col]) if cc_col else [],
                bcc=parse_email_list(record[bcc_col]) if bcc_col else [],
                raw=raw,
            )
        )
    return rows
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```powershell
rtk .venv\Scripts\python.exe -m unittest tests.test_mail_merge
```

Expected: 6 tests pass.

- [ ] **Step 5: Commit**

Run:

```powershell
rtk git add mail_merge.py tests/test_mail_merge.py
rtk git commit -m "feat: load mail merge recipients from excel"
```

---

### Task 3: Email Job Builder And Strict Validation

**Files:**
- Modify: `mail_merge.py`
- Modify: `tests/test_mail_merge.py`

- [ ] **Step 1: Add failing tests for job building and validation**

Append to `MailMergeCoreTests`:

```python
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
rtk .venv\Scripts\python.exe -m unittest tests.test_mail_merge
```

Expected: fail with `AttributeError` for `build_email_jobs`.

- [ ] **Step 3: Implement job builder and validation**

Append to `mail_merge.py`:

```python
def _recipient_context(recipient: RecipientRow, split_result: SplitResult) -> dict[str, object]:
    context: dict[str, object] = {}
    context.update(recipient.raw)
    context.update(
        {
            "key": split_result.key,
            "to": "; ".join(recipient.to),
            "cc": "; ".join(recipient.cc),
            "bcc": "; ".join(recipient.bcc),
            "excel_file": split_result.excel_path or "",
            "pdf_file": split_result.pdf_path or "",
        }
    )
    return context


def _validate_addresses(label: str, addresses: Iterable[str]) -> list[str]:
    return [f"Invalid {label} email: {address}" for address in addresses if not is_valid_email(address)]


def _selected_attachments(
    split_result: SplitResult,
    selection: AttachmentSelection,
    errors: list[str],
) -> list[Path]:
    attachments: list[Path] = []
    if selection.attach_excel:
        if split_result.excel_path and split_result.excel_path.exists():
            attachments.append(split_result.excel_path)
        else:
            errors.append(f"Selected Excel attachment is missing for key {split_result.key}")
    if selection.attach_pdf:
        if split_result.pdf_path and split_result.pdf_path.exists():
            attachments.append(split_result.pdf_path)
        else:
            errors.append(f"Selected PDF attachment is missing for key {split_result.key}")
    return attachments


def build_email_jobs(
    split_results: list[SplitResult],
    recipients: list[RecipientRow],
    template: EmailTemplate,
    attachments: AttachmentSelection,
) -> tuple[list[EmailJob], list[str]]:
    recipient_by_key = {row.key: row for row in recipients}
    split_keys = {result.key for result in split_results}
    warnings = [
        f"Recipient mapping key {row.key} does not match a generated split file"
        for row in recipients
        if row.key not in split_keys
    ]
    jobs: list[EmailJob] = []

    for split_result in split_results:
        errors: list[str] = []
        recipient = recipient_by_key.get(split_result.key)
        if recipient is None:
            errors.append(f"No recipient mapping for key {split_result.key}")
            recipient = RecipientRow(key=split_result.key, to=[], cc=[], bcc=[], raw={})

        if not recipient.to:
            errors.append(f"Required To is empty for key {split_result.key}")
        errors.extend(_validate_addresses("To", recipient.to))
        errors.extend(_validate_addresses("CC", recipient.cc))
        errors.extend(_validate_addresses("BCC", recipient.bcc))

        context = _recipient_context(recipient, split_result)
        subject = render_placeholders(template.subject, context).strip()
        body = render_placeholders(template.body, context).strip()
        if not subject:
            errors.append(f"Rendered subject is empty for key {split_result.key}")
        if not body:
            errors.append(f"Rendered body is empty for key {split_result.key}")

        selected_files = _selected_attachments(split_result, attachments, errors)
        jobs.append(
            EmailJob(
                key=split_result.key,
                to=recipient.to,
                cc=recipient.cc,
                bcc=recipient.bcc,
                subject=subject,
                body=body,
                is_html=template.is_html,
                attachments=selected_files,
                validation_errors=errors,
                validation_warnings=[],
            )
        )

    return jobs, warnings


def all_jobs_valid(jobs: list[EmailJob]) -> bool:
    return all(job.is_valid for job in jobs)
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```powershell
rtk .venv\Scripts\python.exe -m unittest tests.test_mail_merge
```

Expected: 9 tests pass.

- [ ] **Step 5: Commit**

Run:

```powershell
rtk git add mail_merge.py tests/test_mail_merge.py
rtk git commit -m "feat: build and validate mail merge jobs"
```

---

### Task 4: Provider Interface, Fake Provider, And Outlook Provider

**Files:**
- Modify: `mail_merge.py`
- Modify: `tests/test_mail_merge.py`

- [ ] **Step 1: Add failing provider tests**

Append to `MailMergeCoreTests`:

```python
    def test_fake_provider_records_sent_jobs(self):
        provider = mail_merge.FakeMailProvider()
        job = mail_merge.EmailJob(
            key="A",
            to=["a@example.com"],
            cc=[],
            bcc=[],
            subject="Subject",
            body="Body",
            is_html=False,
            attachments=[],
        )

        result = provider.send(job, mail_merge.SendTimingOptions())

        self.assertEqual(result.status, "sent")
        self.assertEqual(provider.sent_jobs, [job])

    def test_outlook_provider_sets_recipients_attachments_and_deferred_delivery(self):
        calls = []

        class FakeAttachments:
            def Add(self, path):
                calls.append(("attach", path))

        class FakeRecipients:
            def Add(self, value):
                calls.append(("recipient", value))
                recipient = type("Recipient", (), {})()
                recipient.Type = None
                return recipient

        class FakeMessage:
            def __init__(self):
                self.To = ""
                self.CC = ""
                self.BCC = ""
                self.Subject = ""
                self.Body = ""
                self.HTMLBody = ""
                self.DeferredDeliveryTime = None
                self.Attachments = FakeAttachments()
                self.Recipients = FakeRecipients()

            def Send(self):
                calls.append(("send", self.Subject, self.DeferredDeliveryTime))

        class FakeOutlook:
            def CreateItem(self, item_type):
                calls.append(("create", item_type))
                return FakeMessage()

        now = datetime(2026, 5, 27, 15, 0, 0)
        provider = mail_merge.OutlookMailProvider(
            dispatcher=lambda name: FakeOutlook(),
            now_fn=lambda: now,
        )
        attachment = Path("C:/tmp/A.pdf")
        job = mail_merge.EmailJob(
            key="A",
            to=["a@example.com"],
            cc=["c@example.com"],
            bcc=["b@example.com"],
            subject="Subject A",
            body="<p>Body</p>",
            is_html=True,
            attachments=[attachment],
        )

        result = provider.send(
            job,
            mail_merge.SendTimingOptions(
                delay_delivery_enabled=True,
                delay_delivery_minutes=5,
                throttle_enabled=False,
                throttle_seconds=0,
            ),
        )

        self.assertEqual(result.status, "sent")
        self.assertIn(("create", 0), calls)
        self.assertIn(("attach", str(attachment)), calls)
        self.assertIn(("send", "Subject A", now + mail_merge.timedelta(minutes=5)), calls)
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
rtk .venv\Scripts\python.exe -m unittest tests.test_mail_merge
```

Expected: fail with `AttributeError` for `FakeMailProvider`.

- [ ] **Step 3: Implement providers**

Append to `mail_merge.py`:

```python
class FakeMailProvider:
    def __init__(self):
        self.sent_jobs: list[EmailJob] = []

    def send(self, job: EmailJob, timing: SendTimingOptions) -> SendResult:
        self.sent_jobs.append(job)
        return SendResult(key=job.key, to=job.to, status="sent", message="fake")


class OutlookMailProvider:
    def __init__(
        self,
        dispatcher: Callable[[str], object] | None = None,
        now_fn: Callable[[], datetime] | None = None,
    ):
        self.dispatcher = dispatcher
        self.now_fn = now_fn or datetime.now

    def _dispatch(self):
        if self.dispatcher is not None:
            return self.dispatcher("Outlook.Application")
        import win32com.client
        return win32com.client.Dispatch("Outlook.Application")

    def send(self, job: EmailJob, timing: SendTimingOptions) -> SendResult:
        outlook = self._dispatch()
        message = outlook.CreateItem(0)
        message.To = "; ".join(job.to)
        message.CC = "; ".join(job.cc)
        message.BCC = "; ".join(job.bcc)
        message.Subject = job.subject
        if job.is_html:
            message.HTMLBody = job.body
        else:
            message.Body = job.body
        for attachment in job.attachments:
            message.Attachments.Add(str(attachment))
        if timing.delay_delivery_enabled and timing.delay_delivery_minutes > 0:
            message.DeferredDeliveryTime = self.now_fn() + timedelta(minutes=timing.delay_delivery_minutes)
        message.Send()
        return SendResult(key=job.key, to=job.to, status="sent", message="outlook")
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```powershell
rtk .venv\Scripts\python.exe -m unittest tests.test_mail_merge
```

Expected: 11 tests pass.

- [ ] **Step 5: Commit**

Run:

```powershell
rtk git add mail_merge.py tests/test_mail_merge.py
rtk git commit -m "feat: add outlook mail provider"
```

---

### Task 5: Split Result Manifest

**Files:**
- Modify: `main.py`
- Modify: `tests/test_templating_modes.py`

- [ ] **Step 1: Add failing split manifest tests**

Append to `SourceTemplateSplitTests` in `tests/test_templating_modes.py`:

```python
    def test_split_returns_manifest_for_excel_and_pdf_outputs(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "source.xlsx"
            out_dir = tmp_path / "out"
            wb = Workbook()
            ws = wb.active
            ws.title = "Data"
            ws.append(["Dept", "Name"])
            ws.append(["A", "Alice"])
            wb.save(source)

            original_export = main.export_pdf_via_lo

            def fake_export(xlsx_path, soffice_path=None):
                xlsx_path.with_suffix(".pdf").write_bytes(b"%PDF-1.4\n")

            try:
                main.export_pdf_via_lo = fake_export
                results = main.split_excel_with_template(
                    source,
                    "Data",
                    "Dept",
                    source,
                    out_dir,
                    1,
                    pdf_engine="libreoffice",
                    template_mode="source_template",
                    output_file_type=main.OUTPUT_TYPE_EXCEL_AND_PDF,
                )
            finally:
                main.export_pdf_via_lo = original_export

            self.assertEqual(len(results), 1)
            self.assertEqual(results[0].key, "A")
            self.assertEqual(results[0].excel_path, out_dir / "A.xlsx")
            self.assertEqual(results[0].pdf_path, out_dir / "A.pdf")
            self.assertEqual(results[0].output_file_type, main.OUTPUT_TYPE_EXCEL_AND_PDF)
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```powershell
rtk .venv\Scripts\python.exe -m unittest tests.test_templating_modes.SourceTemplateSplitTests.test_split_returns_manifest_for_excel_and_pdf_outputs
```

Expected: fail because `split_excel_with_template` returns `None`.

- [ ] **Step 3: Return manifests from split logic**

Modify imports near the top of `main.py`:

```python
from mail_merge import SplitResult
```

In `split_excel_with_template`, after `out_dir.mkdir(parents=True, exist_ok=True)`, add:

```python
    split_results: list[SplitResult] = []
```

In both source-template and template-file branches, after PDF export/removal, append:

```python
            pdf_out = xlsx_out.with_suffix(".pdf")
            split_results.append(
                SplitResult(
                    key=str(key_val),
                    excel_path=xlsx_out if xlsx_out.exists() else None,
                    pdf_path=pdf_out if pdf_out.exists() else None,
                    output_file_type=output_file_type,
                )
            )
```

At the end of `split_excel_with_template`, after optional cleanup, return:

```python
    return split_results
```

In `SplitWorker.__init__`, initialize:

```python
        self.results = []
```

In `SplitWorker.run`, assign the return value:

```python
            self.results = split_excel_with_template(
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
                template_mode=self.params.get('template_mode', TEMPLATE_MODE_TEMPLATE_FILE),
                column_mapping=self.params.get('column_mapping'),
                source_header_rows=self.params.get('source_header_rows'),
                template_header_rows=self.params.get('template_header_rows'),
                output_file_type=self.params.get('output_file_type'),
                status_cb=self.status.emit,
                progress_cb=lambda t, c: self.progress.emit(t, c)
            )
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```powershell
rtk .venv\Scripts\python.exe -m unittest tests.test_templating_modes
```

Expected: all templating mode tests pass.

- [ ] **Step 5: Commit**

Run:

```powershell
rtk git add main.py tests/test_templating_modes.py
rtk git commit -m "feat: return split result manifest"
```

---

### Task 6: Mail Merge Entry Button And Panel Shell

**Files:**
- Modify: `main.py`
- Modify: `tests/test_ui_smoke.py`

- [ ] **Step 1: Add failing UI shell tests**

Append to `UISmokeTests` in `tests/test_ui_smoke.py`:

```python
    def test_mail_merge_button_hidden_until_split_results_exist(self):
        with tempfile.TemporaryDirectory() as tmp:
            window = main.SplitApp(settings=self.make_settings(Path(tmp) / "settings.ini"))
            self.addCleanup(window.deleteLater)

            self.assertTrue(hasattr(window, "btn_mail_merge"))
            self.assertTrue(window.btn_mail_merge.isHidden())

            window.current_split_results = [
                main.SplitResult(key="A", excel_path=Path(tmp) / "A.xlsx", output_file_type=main.OUTPUT_TYPE_EXCEL)
            ]
            window.update_mail_merge_entry_state()

            self.assertFalse(window.btn_mail_merge.isHidden())

    def test_show_mail_merge_panel_reveals_panel_with_loaded_count(self):
        with tempfile.TemporaryDirectory() as tmp:
            window = main.SplitApp(settings=self.make_settings(Path(tmp) / "settings.ini"))
            self.addCleanup(window.deleteLater)
            window.current_split_results = [
                main.SplitResult(key="A", excel_path=Path(tmp) / "A.xlsx", output_file_type=main.OUTPUT_TYPE_EXCEL)
            ]

            window.show_mail_merge_panel()

            self.assertFalse(window.mail_merge_card.isHidden())
            self.assertIn("1 split file", window.lbl_mail_merge_summary.text())
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
rtk .venv\Scripts\python.exe -m unittest tests.test_ui_smoke.UISmokeTests.test_mail_merge_button_hidden_until_split_results_exist tests.test_ui_smoke.UISmokeTests.test_show_mail_merge_panel_reveals_panel_with_loaded_count
```

Expected: fail because `btn_mail_merge` does not exist.

- [ ] **Step 3: Add Mail Merge shell UI**

In `SplitApp.__init__`, after `self.field_action_buttons = []`, add:

```python
        self.current_split_results = []
        self.current_mail_jobs = []
        self.current_mail_warnings = []
        self.current_preview_index = 0
```

In `_build_ui`, call a new card builder after `_build_log_card()`:

```python
        self._build_mail_merge_card()
```

In `_build_actions_card`, after `btn_open_output`, add:

```python
        self.btn_mail_merge = PushButton(FIF.MAIL, "Mail Merge")
        self.btn_mail_merge.setFixedHeight(36)
        self.btn_mail_merge.clicked.connect(self.show_mail_merge_panel)
        self.btn_mail_merge.setVisible(False)
        layout.addWidget(self.btn_mail_merge)
```

Add methods to `SplitApp`:

```python
    def _build_mail_merge_card(self):
        self.mail_merge_card, layout = self._panel("Mail Merge", FIF.MAIL)
        self.lbl_mail_merge_summary = CaptionLabel("No split results loaded.")
        layout.addWidget(self.lbl_mail_merge_summary)
        self.mail_merge_card.setVisible(False)
        self.main_panel_layout.addWidget(self.mail_merge_card)

    def update_mail_merge_entry_state(self):
        has_results = bool(self.current_split_results)
        self.btn_mail_merge.setVisible(has_results)

    def show_mail_merge_panel(self):
        count = len(self.current_split_results)
        suffix = "file" if count == 1 else "files"
        self.lbl_mail_merge_summary.setText(f"{count} split {suffix} loaded for mail merge.")
        self.mail_merge_card.setVisible(True)
        self.update_mail_merge_entry_state()
```

In `_on_worker_finished`, before `self.btn_open_output.setVisible(True)`, add:

```python
        self.current_split_results = list(getattr(self.worker, "results", []))
        self.update_mail_merge_entry_state()
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```powershell
rtk .venv\Scripts\python.exe -m unittest tests.test_ui_smoke
```

Expected: UI smoke tests pass.

- [ ] **Step 5: Commit**

Run:

```powershell
rtk git add main.py tests/test_ui_smoke.py
rtk git commit -m "feat: add mail merge entry panel"
```

---

### Task 7: Recipient Mapping UI And Settings

**Files:**
- Modify: `main.py`
- Modify: `tests/test_ui_smoke.py`
- Modify: `tests/test_settings.py`

- [ ] **Step 1: Add failing UI/settings tests**

Append to `UISmokeTests`:

```python
    def test_mail_merge_recipient_controls_exist(self):
        with tempfile.TemporaryDirectory() as tmp:
            window = main.SplitApp(settings=self.make_settings(Path(tmp) / "settings.ini"))
            self.addCleanup(window.deleteLater)

            self.assertTrue(hasattr(window, "edit_recipient_path"))
            self.assertTrue(hasattr(window, "cmb_recipient_sheet"))
            self.assertTrue(hasattr(window, "spin_recipient_header_row"))
            self.assertTrue(hasattr(window, "cmb_recipient_key"))
            self.assertTrue(hasattr(window, "cmb_recipient_to"))
            self.assertTrue(hasattr(window, "cmb_recipient_cc"))
            self.assertTrue(hasattr(window, "cmb_recipient_bcc"))
```

Append to `SettingsTests`:

```python
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
            first.save_settings()

            second = main.SplitApp(settings=QSettings(str(settings_path), QSettings.IniFormat))
            self.addCleanup(second.deleteLater)

            self.assertEqual(second.edit_recipient_path.text(), "recipients.xlsx")
            self.assertEqual(second.cmb_recipient_sheet.currentText(), "Recipients")
            self.assertEqual(second.spin_recipient_header_row.value(), 2)
            self.assertEqual(second.edit_mail_subject.text(), "Subject {key}")
            self.assertEqual(second.edit_mail_html_template.text(), "body.html")
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
rtk .venv\Scripts\python.exe -m unittest tests.test_ui_smoke.UISmokeTests.test_mail_merge_recipient_controls_exist tests.test_settings.SettingsTests.test_mail_merge_settings_persist_with_qsettings
```

Expected: fail because Mail Merge recipient controls do not exist.

- [ ] **Step 3: Add recipient and template controls**

Import from `mail_merge` near the top of `main.py`:

```python
from mail_merge import (
    all_jobs_valid,
    AttachmentSelection,
    EmailJob,
    EmailTemplate,
    SplitResult,
    build_email_jobs,
    load_recipient_rows,
    read_recipient_headers,
)
```

Extend `_build_mail_merge_card` after summary label:

```python
        recipient_grid = QGridLayout()
        recipient_grid.setHorizontalSpacing(10)
        recipient_grid.setVerticalSpacing(8)
        self.edit_recipient_path = self._fixed_width(LineEdit(), PATH_FIELD_WIDTH)
        self.edit_recipient_path.setPlaceholderText("Recipient mapping Excel file")
        self.btn_browse_recipients = self._field_action_button(ToolButton(FIF.FOLDER))
        self.btn_browse_recipients.clicked.connect(self.browse_recipient_mapping)
        self.cmb_recipient_sheet = self._fixed_width(ComboBox(), COMBO_FIELD_WIDTH)
        self.cmb_recipient_sheet.setPlaceholderText("Sheet")
        self.btn_load_recipient_sheets = self._field_action_button(ToolButton(FIF.SYNC))
        self.btn_load_recipient_sheets.setToolTip("Load Recipient Sheets")
        self.btn_load_recipient_sheets.clicked.connect(self.load_recipient_sheets)
        self.spin_recipient_header_row = self._fixed_width(SpinBox(), SMALL_FIELD_WIDTH)
        self.spin_recipient_header_row.setRange(1, 100)
        self.spin_recipient_header_row.setValue(1)
        self.btn_load_recipient_headers = self._field_action_button(ToolButton(FIF.SYNC))
        self.btn_load_recipient_headers.setToolTip("Load Recipient Headers")
        self.btn_load_recipient_headers.clicked.connect(self.load_recipient_headers)
        self.cmb_recipient_key = self._fixed_width(ComboBox(), COMBO_FIELD_WIDTH)
        self.cmb_recipient_to = self._fixed_width(ComboBox(), COMBO_FIELD_WIDTH)
        self.cmb_recipient_cc = self._fixed_width(ComboBox(), COMBO_FIELD_WIDTH)
        self.cmb_recipient_bcc = self._fixed_width(ComboBox(), COMBO_FIELD_WIDTH)
        recipient_grid.addWidget(self._labeled("Recipient Workbook", self.edit_recipient_path), 0, 0)
        recipient_grid.addWidget(self.btn_browse_recipients, 0, 1, Qt.AlignBottom)
        recipient_grid.addWidget(self._labeled("Sheet", self.cmb_recipient_sheet), 0, 2)
        recipient_grid.addWidget(self.btn_load_recipient_sheets, 0, 3, Qt.AlignBottom)
        recipient_grid.addWidget(self._labeled("Header Row", self.spin_recipient_header_row), 0, 4)
        recipient_grid.addWidget(self.btn_load_recipient_headers, 0, 5, Qt.AlignBottom)
        recipient_grid.addWidget(self._labeled("Key", self.cmb_recipient_key), 1, 0)
        recipient_grid.addWidget(self._labeled("To", self.cmb_recipient_to), 1, 1)
        recipient_grid.addWidget(self._labeled("CC", self.cmb_recipient_cc), 1, 2)
        recipient_grid.addWidget(self._labeled("BCC", self.cmb_recipient_bcc), 1, 3)
        recipient_grid.setColumnStretch(6, 1)
        layout.addLayout(recipient_grid)

        content_grid = QGridLayout()
        content_grid.setHorizontalSpacing(10)
        content_grid.setVerticalSpacing(8)
        self.edit_mail_subject = self._fixed_width(LineEdit(), PATH_FIELD_WIDTH)
        self.edit_mail_subject.setPlaceholderText("Subject, e.g. Report {key}")
        self.edit_mail_body = TextEdit()
        self.edit_mail_body.setMinimumHeight(90)
        self.edit_mail_html_template = self._fixed_width(LineEdit(), PATH_FIELD_WIDTH)
        self.edit_mail_html_template.setPlaceholderText("Optional HTML template")
        self.btn_browse_mail_html = self._field_action_button(ToolButton(FIF.FOLDER))
        self.btn_browse_mail_html.clicked.connect(self.browse_mail_html_template)
        content_grid.addWidget(self._labeled("Subject", self.edit_mail_subject), 0, 0)
        content_grid.addWidget(self._labeled("Body", self.edit_mail_body), 1, 0)
        content_grid.addWidget(self._labeled("HTML Template", self.edit_mail_html_template), 2, 0)
        content_grid.addWidget(self.btn_browse_mail_html, 2, 1, Qt.AlignBottom)
        layout.addLayout(content_grid)
```

Add methods:

```python
    def browse_recipient_mapping(self):
        f, _ = QFileDialog.getOpenFileName(
            self, "Pilih recipient mapping Excel",
            "", "Excel files (*.xlsx *.xls *.xlsm)"
        )
        if f:
            self.edit_recipient_path.setText(f)
            self.load_recipient_sheets()

    def load_recipient_sheets(self):
        path = self.edit_recipient_path.text().strip()
        if not path:
            InfoBar.warning("Perhatian", "Pilih recipient mapping Excel dulu.", parent=self, duration=3000, position=InfoBarPosition.TOP)
            return
        try:
            with pd.ExcelFile(path) as xls:
                sheets = list(xls.sheet_names)
            self.cmb_recipient_sheet.clear()
            self.cmb_recipient_sheet.addItems(sheets)
            if sheets:
                self.cmb_recipient_sheet.setCurrentIndex(0)
                self.load_recipient_headers()
            self.save_settings()
        except Exception as e:
            InfoBar.error("Error", str(e), parent=self, duration=5000, position=InfoBarPosition.TOP)

    def load_recipient_headers(self):
        path = self.edit_recipient_path.text().strip()
        sheet = self.cmb_recipient_sheet.currentText().strip()
        if not path or not sheet:
            return
        try:
            headers = read_recipient_headers(Path(path), sheet, self.spin_recipient_header_row.value())
            choices = [""] + headers
            for combo in [self.cmb_recipient_key, self.cmb_recipient_to, self.cmb_recipient_cc, self.cmb_recipient_bcc]:
                combo.clear()
                combo.addItems(choices)
            self._select_combo_text(self.cmb_recipient_key, "Key")
            self._select_combo_text(self.cmb_recipient_to, "To")
            self._select_combo_text(self.cmb_recipient_cc, "CC")
            self._select_combo_text(self.cmb_recipient_bcc, "BCC")
            self.save_settings()
        except Exception as e:
            InfoBar.error("Error", str(e), parent=self, duration=5000, position=InfoBarPosition.TOP)

    def _select_combo_text(self, combo, text):
        idx = combo.findText(text)
        if idx >= 0:
            combo.setCurrentIndex(idx)

    def browse_mail_html_template(self):
        f, _ = QFileDialog.getOpenFileName(self, "Pilih HTML template", "", "HTML files (*.html *.htm)")
        if f:
            self.edit_mail_html_template.setText(f)
```

Extend `_connect_settings_signals`, `save_settings`, `load_settings`, and `reset_settings` to include the new fields:

```python
        for edit in [self.edit_recipient_path, self.edit_mail_subject, self.edit_mail_html_template]:
            edit.editingFinished.connect(self.save_settings)
        self.edit_mail_body.textChanged.connect(self.save_settings)
        for combo in [self.cmb_recipient_sheet, self.cmb_recipient_key, self.cmb_recipient_to, self.cmb_recipient_cc, self.cmb_recipient_bcc]:
            combo.currentTextChanged.connect(self.save_settings)
        self.spin_recipient_header_row.valueChanged.connect(self.save_settings)
```

Use QSettings keys:

```python
        self.settings.setValue("mail_recipient_path", self.edit_recipient_path.text().strip())
        self.settings.setValue("mail_recipient_sheet", self.cmb_recipient_sheet.currentText().strip())
        self.settings.setValue("mail_recipient_header_row", self.spin_recipient_header_row.value())
        self.settings.setValue("mail_recipient_key_col", self.cmb_recipient_key.currentText().strip())
        self.settings.setValue("mail_recipient_to_col", self.cmb_recipient_to.currentText().strip())
        self.settings.setValue("mail_recipient_cc_col", self.cmb_recipient_cc.currentText().strip())
        self.settings.setValue("mail_recipient_bcc_col", self.cmb_recipient_bcc.currentText().strip())
        self.settings.setValue("mail_subject", self.edit_mail_subject.text().strip())
        self.settings.setValue("mail_body", self.edit_mail_body.toPlainText())
        self.settings.setValue("mail_html_template", self.edit_mail_html_template.text().strip())
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```powershell
rtk .venv\Scripts\python.exe -m unittest tests.test_ui_smoke tests.test_settings
```

Expected: UI and settings tests pass.

- [ ] **Step 5: Commit**

Run:

```powershell
rtk git add main.py tests/test_ui_smoke.py tests/test_settings.py
rtk git commit -m "feat: add mail merge recipient controls"
```

---

### Task 8: Attachment Options, Timing Controls, Job Build, And Carousel Preview

**Files:**
- Modify: `main.py`
- Modify: `tests/test_ui_smoke.py`

- [ ] **Step 1: Add failing UI carousel tests**

Append to `UISmokeTests`:

```python
    def test_mail_merge_preview_carousel_moves_between_jobs(self):
        with tempfile.TemporaryDirectory() as tmp:
            window = main.SplitApp(settings=self.make_settings(Path(tmp) / "settings.ini"))
            self.addCleanup(window.deleteLater)
            window.current_mail_jobs = [
                main.EmailJob("A", ["a@example.com"], [], [], "Subject A", "Body A", False, []),
                main.EmailJob("B", ["b@example.com"], [], [], "Subject B", "Body B", False, []),
            ]
            window.current_preview_index = 0

            window.render_mail_preview()
            self.assertIn("1 / 2", window.lbl_mail_preview_count.text())
            self.assertIn("Subject A", window.lbl_mail_preview_subject.text())

            window.next_mail_preview()

            self.assertIn("2 / 2", window.lbl_mail_preview_count.text())
            self.assertIn("Subject B", window.lbl_mail_preview_subject.text())

    def test_mail_merge_send_disabled_when_validation_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            window = main.SplitApp(settings=self.make_settings(Path(tmp) / "settings.ini"))
            self.addCleanup(window.deleteLater)
            window.current_mail_jobs = [
                main.EmailJob("A", [], [], [], "", "", False, [], validation_errors=["Required To is empty for key A"])
            ]

            window.render_mail_preview()

            self.assertFalse(window.btn_send_mail_merge.isEnabled())
            self.assertIn("1 issue", window.lbl_mail_validation_summary.text())
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
rtk .venv\Scripts\python.exe -m unittest tests.test_ui_smoke.UISmokeTests.test_mail_merge_preview_carousel_moves_between_jobs tests.test_ui_smoke.UISmokeTests.test_mail_merge_send_disabled_when_validation_fails
```

Expected: fail because preview widgets do not exist.

- [ ] **Step 3: Add controls and preview behavior**

Import `CheckBox` from `qfluentwidgets` in the existing widget import block:

```python
    CheckBox,
```

Extend `_build_mail_merge_card`:

```python
        options_grid = QGridLayout()
        options_grid.setHorizontalSpacing(10)
        options_grid.setVerticalSpacing(8)
        self.chk_attach_excel = CheckBox("Attach Excel")
        self.chk_attach_excel.setChecked(True)
        self.chk_attach_pdf = CheckBox("Attach PDF")
        self.chk_delay_delivery = CheckBox("Delay delivery")
        self.chk_delay_delivery.setChecked(True)
        self.spin_delay_minutes = self._fixed_width(SpinBox(), SMALL_FIELD_WIDTH)
        self.spin_delay_minutes.setRange(0, 1440)
        self.spin_delay_minutes.setValue(5)
        self.chk_throttle = CheckBox("Throttle")
        self.chk_throttle.setChecked(True)
        self.spin_throttle_seconds = self._fixed_width(SpinBox(), SMALL_FIELD_WIDTH)
        self.spin_throttle_seconds.setRange(0, 3600)
        self.spin_throttle_seconds.setValue(5)
        self.btn_build_mail_preview = PushButton("Build Preview")
        self.btn_build_mail_preview.clicked.connect(self.build_mail_preview)
        options_grid.addWidget(self.chk_attach_excel, 0, 0)
        options_grid.addWidget(self.chk_attach_pdf, 0, 1)
        options_grid.addWidget(self.chk_delay_delivery, 0, 2)
        options_grid.addWidget(self._labeled("Minutes", self.spin_delay_minutes), 0, 3)
        options_grid.addWidget(self.chk_throttle, 0, 4)
        options_grid.addWidget(self._labeled("Seconds", self.spin_throttle_seconds), 0, 5)
        options_grid.addWidget(self.btn_build_mail_preview, 0, 6)
        layout.addLayout(options_grid)

        self.lbl_mail_validation_summary = CaptionLabel("Build preview to validate emails.")
        self.lbl_mail_preview_count = BodyLabel("0 / 0")
        self.lbl_mail_preview_key = BodyLabel("")
        self.lbl_mail_preview_recipients = BodyLabel("")
        self.lbl_mail_preview_subject = BodyLabel("")
        self.txt_mail_preview_body = TextEdit()
        self.txt_mail_preview_body.setReadOnly(True)
        self.txt_mail_preview_body.setMinimumHeight(100)
        self.lbl_mail_preview_attachments = CaptionLabel("")
        self.lbl_mail_preview_errors = CaptionLabel("")
        nav = QHBoxLayout()
        self.btn_prev_mail_preview = PushButton("Previous")
        self.btn_next_mail_preview = PushButton("Next")
        self.btn_prev_mail_preview.clicked.connect(self.prev_mail_preview)
        self.btn_next_mail_preview.clicked.connect(self.next_mail_preview)
        nav.addWidget(self.btn_prev_mail_preview)
        nav.addWidget(self.lbl_mail_preview_count)
        nav.addWidget(self.btn_next_mail_preview)
        nav.addStretch()
        layout.addWidget(self.lbl_mail_validation_summary)
        layout.addLayout(nav)
        layout.addWidget(self.lbl_mail_preview_key)
        layout.addWidget(self.lbl_mail_preview_recipients)
        layout.addWidget(self.lbl_mail_preview_subject)
        layout.addWidget(self.txt_mail_preview_body)
        layout.addWidget(self.lbl_mail_preview_attachments)
        layout.addWidget(self.lbl_mail_preview_errors)
        self.btn_send_mail_merge = PrimaryPushButton(FIF.SEND, "Send")
        self.btn_send_mail_merge.setEnabled(False)
        layout.addWidget(self.btn_send_mail_merge)
```

Add methods:

```python
    def current_attachment_selection(self):
        return AttachmentSelection(
            attach_excel=self.chk_attach_excel.isChecked(),
            attach_pdf=self.chk_attach_pdf.isChecked(),
        )

    def current_email_template(self):
        html_path = self.edit_mail_html_template.text().strip()
        body = self.edit_mail_body.toPlainText()
        is_html = False
        if html_path and Path(html_path).exists():
            body = Path(html_path).read_text(encoding="utf-8")
            is_html = True
        return EmailTemplate(
            subject=self.edit_mail_subject.text().strip(),
            body=body,
            is_html=is_html,
            template_path=Path(html_path) if html_path else None,
        )

    def build_mail_preview(self):
        rows = load_recipient_rows(
            Path(self.edit_recipient_path.text().strip()),
            self.cmb_recipient_sheet.currentText().strip(),
            self.spin_recipient_header_row.value(),
            {
                "key": self.cmb_recipient_key.currentText().strip(),
                "to": self.cmb_recipient_to.currentText().strip(),
                "cc": self.cmb_recipient_cc.currentText().strip(),
                "bcc": self.cmb_recipient_bcc.currentText().strip(),
            },
        )
        self.current_mail_jobs, self.current_mail_warnings = build_email_jobs(
            split_results=self.current_split_results,
            recipients=rows,
            template=self.current_email_template(),
            attachments=self.current_attachment_selection(),
        )
        self.current_preview_index = 0
        self.render_mail_preview()

    def render_mail_preview(self):
        total = len(self.current_mail_jobs)
        if total == 0:
            self.lbl_mail_preview_count.setText("0 / 0")
            self.btn_send_mail_merge.setEnabled(False)
            return
        self.current_preview_index = max(0, min(self.current_preview_index, total - 1))
        job = self.current_mail_jobs[self.current_preview_index]
        issue_count = sum(len(item.validation_errors) for item in self.current_mail_jobs)
        ready_count = sum(1 for item in self.current_mail_jobs if item.is_valid)
        issue_label = "issue" if issue_count == 1 else "issues"
        self.lbl_mail_validation_summary.setText(
            f"{ready_count} emails ready, {issue_count} {issue_label} found"
        )
        self.lbl_mail_preview_count.setText(f"{self.current_preview_index + 1} / {total}")
        self.lbl_mail_preview_key.setText(f"Key: {job.key}")
        self.lbl_mail_preview_recipients.setText(
            f"To: {'; '.join(job.to)} | CC: {'; '.join(job.cc)} | BCC: {'; '.join(job.bcc)}"
        )
        self.lbl_mail_preview_subject.setText(f"Subject: {job.subject}")
        self.txt_mail_preview_body.setPlainText(job.body)
        self.lbl_mail_preview_attachments.setText(
            "Attachments: " + "; ".join(str(path) for path in job.attachments)
        )
        self.lbl_mail_preview_errors.setText("Errors: " + "; ".join(job.validation_errors))
        self.btn_send_mail_merge.setEnabled(all_jobs_valid(self.current_mail_jobs))

    def next_mail_preview(self):
        if self.current_preview_index < len(self.current_mail_jobs) - 1:
            self.current_preview_index += 1
        self.render_mail_preview()

    def prev_mail_preview(self):
        if self.current_preview_index > 0:
            self.current_preview_index -= 1
        self.render_mail_preview()
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```powershell
rtk .venv\Scripts\python.exe -m unittest tests.test_ui_smoke
```

Expected: UI smoke tests pass.

- [ ] **Step 5: Commit**

Run:

```powershell
rtk git add main.py tests/test_ui_smoke.py
rtk git commit -m "feat: add mail merge preview carousel"
```

---

### Task 9: Send Worker, Progress, Cancel, And Timing

**Files:**
- Modify: `main.py`
- Modify: `tests/test_mail_merge.py`
- Modify: `tests/test_ui_smoke.py`

- [ ] **Step 1: Add failing send timing test**

Append to `MailMergeCoreTests`:

```python
    def test_send_jobs_throttles_between_jobs_and_can_cancel(self):
        provider = mail_merge.FakeMailProvider()
        sleeps = []
        statuses = []
        jobs = [
            mail_merge.EmailJob("A", ["a@example.com"], [], [], "A", "Body", False, []),
            mail_merge.EmailJob("B", ["b@example.com"], [], [], "B", "Body", False, []),
        ]

        results = mail_merge.send_jobs(
            jobs,
            provider,
            mail_merge.SendTimingOptions(
                delay_delivery_enabled=False,
                delay_delivery_minutes=0,
                throttle_enabled=True,
                throttle_seconds=5,
            ),
            status_cb=statuses.append,
            sleep_fn=sleeps.append,
        )

        self.assertEqual([result.key for result in results], ["A", "B"])
        self.assertEqual(sleeps, [5])
        self.assertEqual(statuses, ["Sending 1/2: A", "Sending 2/2: B"])
```

Append to `UISmokeTests`:

```python
    def test_current_send_timing_reads_delay_and_throttle_controls(self):
        with tempfile.TemporaryDirectory() as tmp:
            window = main.SplitApp(settings=self.make_settings(Path(tmp) / "settings.ini"))
            self.addCleanup(window.deleteLater)
            window.chk_delay_delivery.setChecked(True)
            window.spin_delay_minutes.setValue(7)
            window.chk_throttle.setChecked(True)
            window.spin_throttle_seconds.setValue(3)

            timing = window.current_send_timing()

            self.assertTrue(timing.delay_delivery_enabled)
            self.assertEqual(timing.delay_delivery_minutes, 7)
            self.assertTrue(timing.throttle_enabled)
            self.assertEqual(timing.throttle_seconds, 3)
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
rtk .venv\Scripts\python.exe -m unittest tests.test_mail_merge.MailMergeCoreTests.test_send_jobs_throttles_between_jobs_and_can_cancel tests.test_ui_smoke.UISmokeTests.test_current_send_timing_reads_delay_and_throttle_controls
```

Expected: mail merge test fails until `send_jobs` exists, and UI test fails until `current_send_timing` exists.

- [ ] **Step 3: Add send worker and UI wiring**

Append `send_jobs` to `mail_merge.py`:

```python
def send_jobs(
    jobs: list[EmailJob],
    provider: MailProvider,
    timing: SendTimingOptions,
    status_cb: Callable[[str], None] | None = None,
    stop_requested: Callable[[], bool] | None = None,
    sleep_fn: Callable[[float], None] = time.sleep,
) -> list[SendResult]:
    status_cb = status_cb or (lambda message: None)
    stop_requested = stop_requested or (lambda: False)
    results: list[SendResult] = []
    for index, job in enumerate(jobs):
        if stop_requested():
            results.append(SendResult(key=job.key, to=job.to, status="cancelled", message="cancelled before send"))
            break
        status_cb(f"Sending {index + 1}/{len(jobs)}: {job.key}")
        results.append(provider.send(job, timing))
        if timing.throttle_enabled and timing.throttle_seconds > 0 and index < len(jobs) - 1:
            sleep_fn(timing.throttle_seconds)
    return results
```

Import from `mail_merge`:

```python
from mail_merge import OutlookMailProvider, SendTimingOptions, send_jobs
```

Add below `SplitWorker` in `main.py`:

```python
class MailMergeWorker(QThread):
    status = Signal(str)
    finished = Signal(object)
    error = Signal(str)

    def __init__(self, jobs, timing, provider=None):
        super().__init__()
        self.jobs = jobs
        self.timing = timing
        self.provider = provider or OutlookMailProvider()
        self._cancel_requested = False

    def cancel(self):
        self._cancel_requested = True

    def run(self):
        try:
            results = send_jobs(
                self.jobs,
                self.provider,
                self.timing,
                status_cb=self.status.emit,
                stop_requested=lambda: self._cancel_requested,
            )
            self.finished.emit(results)
        except Exception as e:
            self.error.emit(str(e))
```

Add methods to `SplitApp`:

```python
    def current_send_timing(self):
        return SendTimingOptions(
            delay_delivery_enabled=self.chk_delay_delivery.isChecked(),
            delay_delivery_minutes=self.spin_delay_minutes.value(),
            throttle_enabled=self.chk_throttle.isChecked(),
            throttle_seconds=self.spin_throttle_seconds.value(),
        )

    def on_send_mail_merge_clicked(self):
        self.render_mail_preview()
        if not all_jobs_valid(self.current_mail_jobs):
            InfoBar.error("Mail Merge", "Fix validation issues before sending.", parent=self, duration=5000, position=InfoBarPosition.TOP)
            return
        self.btn_send_mail_merge.setEnabled(False)
        self.mail_worker = MailMergeWorker(self.current_mail_jobs, self.current_send_timing())
        self.mail_worker.status.connect(self.log)
        self.mail_worker.finished.connect(self.on_mail_merge_finished)
        self.mail_worker.error.connect(self.on_mail_merge_error)
        self.mail_worker.start()

    def cancel_mail_merge_send(self):
        if hasattr(self, "mail_worker") and self.mail_worker is not None:
            self.mail_worker.cancel()

    def on_mail_merge_finished(self, results):
        for result in results:
            self.log(f"Mail {result.status}: {result.key} {'; '.join(result.to)} {result.message}")
        self.btn_send_mail_merge.setEnabled(all_jobs_valid(self.current_mail_jobs))
        InfoBar.success("Mail Merge", "Send process finished.", parent=self, duration=5000, position=InfoBarPosition.TOP)

    def on_mail_merge_error(self, error_msg):
        self.log(f"Mail Merge error: {error_msg}")
        self.btn_send_mail_merge.setEnabled(all_jobs_valid(self.current_mail_jobs))
        InfoBar.error("Mail Merge", error_msg, parent=self, duration=8000, position=InfoBarPosition.TOP)
```

Wire buttons in `_build_mail_merge_card`:

```python
        self.btn_send_mail_merge.clicked.connect(self.on_send_mail_merge_clicked)
        self.btn_cancel_mail_merge = PushButton("Cancel Send")
        self.btn_cancel_mail_merge.clicked.connect(self.cancel_mail_merge_send)
        layout.addWidget(self.btn_cancel_mail_merge)
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```powershell
rtk .venv\Scripts\python.exe -m unittest tests.test_mail_merge tests.test_ui_smoke
```

Expected: mail merge and UI smoke tests pass.

- [ ] **Step 5: Commit**

Run:

```powershell
rtk git add main.py tests/test_mail_merge.py tests/test_ui_smoke.py
rtk git commit -m "feat: send mail merge jobs through outlook"
```

---

### Task 10: README, Full Verification, And Build

**Files:**
- Modify: `README.md`
- Modify: `dist/ExcelSplitter.exe`

- [ ] **Step 1: Update README**

Add a "Mail Merge" section after "Output Files":

```markdown
### Mail Merge

After a successful split, click **Mail Merge** to send generated files by email.

Recipient mapping is loaded from an Excel worksheet with one row per split key:

- `Key`: matches the split key value
- `To`: required recipient address list
- `CC`: optional
- `BCC`: optional

Multiple email addresses in `To`, `CC`, and `BCC` use semicolon separators.

Mail Merge supports in-app subject/body placeholders such as `{key}`, `{to}`, and columns from the recipient mapping worksheet. An optional `.html` file can be used as the email body template.

Before sending, the app shows a carousel preview so each email can be checked one by one. Strict validation blocks sending if recipients, attachments, subject, body, or Outlook availability are invalid.

The first sending provider is Microsoft Outlook desktop. Delay delivery sets Outlook's deferred delivery time, and throttle controls how quickly the app hands messages to Outlook.
```

- [ ] **Step 2: Run full automated verification**

Run:

```powershell
rtk .venv\Scripts\python.exe -m unittest discover
rtk python -m py_compile main.py mail_merge.py
rtk git diff --check
```

Expected:

- all tests pass
- compile commands exit 0
- diff check has no output

- [ ] **Step 3: Build executable**

Run:

```powershell
rtk cmd /c build.cmd
```

Expected: build exits 0 and prints `Build sukses: dist\ExcelSplitter.exe`.

- [ ] **Step 4: Commit docs and build**

Run:

```powershell
rtk git add README.md dist/ExcelSplitter.exe
rtk git commit -m "docs: document mail merge workflow"
```

- [ ] **Step 5: Final status**

Run:

```powershell
rtk git status -sb
rtk git log --oneline -5
```

Expected: working tree clean and latest commits include the mail merge implementation commits.

---

## Self-Review Notes

- Spec coverage: tasks cover post-split entry, recipient Excel mapping, one-row-per-key recipients, subject/body placeholders, optional HTML body template, attachment selection, carousel preview, strict validation, Outlook sending, delay delivery, throttle, progress/cancel, QSettings, tests, README, and build.
- Scope: Thunderbird, Gmail web, SMTP, CSV mapping, `.docx` body templates, and persisted split manifests are intentionally excluded from implementation tasks.
- Type consistency: plan uses `SplitResult`, `RecipientRow`, `EmailTemplate`, `AttachmentSelection`, `SendTimingOptions`, `EmailJob`, `SendResult`, `FakeMailProvider`, `OutlookMailProvider`, and `MailMergeWorker` consistently across tasks.
