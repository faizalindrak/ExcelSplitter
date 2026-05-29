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


def _build_email_job(
    split_result: SplitResult,
    recipient: RecipientRow,
    template: EmailTemplate,
    attachments: AttachmentSelection,
    initial_errors: list[str] | None = None,
) -> EmailJob:
    errors = list(initial_errors or [])
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
    return EmailJob(
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


def build_email_jobs(
    split_results: list[SplitResult],
    recipients: list[RecipientRow],
    template: EmailTemplate,
    attachments: AttachmentSelection,
) -> tuple[list[EmailJob], list[str]]:
    if not split_results:
        return [
            _build_email_job(SplitResult(key=recipient.key), recipient, template, attachments)
            for recipient in recipients
        ], []

    recipient_by_key = {row.key: row for row in recipients}
    split_keys = {result.key for result in split_results}
    warnings = [
        f"Recipient mapping key {row.key} does not match a generated split file"
        for row in recipients
        if row.key not in split_keys
    ]
    jobs: list[EmailJob] = []

    for split_result in split_results:
        recipient = recipient_by_key.get(split_result.key)
        errors: list[str] = []
        if recipient is None:
            errors.append(f"No recipient mapping for key {split_result.key}")
            recipient = RecipientRow(key=split_result.key, to=[], cc=[], bcc=[], raw={})

        jobs.append(
            _build_email_job(split_result, recipient, template, attachments, initial_errors=errors)
        )

    return jobs, warnings


def all_jobs_valid(jobs: list[EmailJob]) -> bool:
    return all(job.is_valid for job in jobs)


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


def detect_key_from_filename(stem: str, prefix: str = "", suffix: str = "") -> str:
    original = stem.strip()
    result = stem
    if prefix and result.startswith(prefix + " "):
        result = result[len(prefix) + 1:]
    if suffix:
        if result.endswith(" " + suffix):
            result = result[:-len(suffix) - 1]
        elif result == suffix:
            result = ""
    result = result.strip()
    if not result:
        return original
    return result


def discover_split_results_from_folder(
    folder: Path,
    prefix: str = "",
    suffix: str = "",
    recurse: bool = False,
) -> list[SplitResult]:
    if not folder.exists() or not folder.is_dir():
        return []
    
    pattern = "**/*" if recurse else "*"
    files_by_key: dict[str, dict[str, Path]] = {}
    
    for path in sorted(folder.glob(pattern)):
        if not path.is_file():
            continue
        if path.name.startswith("~$"):
            continue
        ext = path.suffix.lower()
        if ext not in [".xlsx", ".pdf"]:
            continue
        
        key = detect_key_from_filename(path.stem, prefix, suffix)
        if key not in files_by_key:
            files_by_key[key] = {}
        
        if ext == ".xlsx" and "xlsx" not in files_by_key[key]:
            files_by_key[key]["xlsx"] = path
        elif ext == ".pdf" and "pdf" not in files_by_key[key]:
            files_by_key[key]["pdf"] = path
    
    results: list[SplitResult] = []
    for key in sorted(files_by_key.keys()):
        files = files_by_key[key]
        excel_path = files.get("xlsx")
        pdf_path = files.get("pdf")
        
        if excel_path and pdf_path:
            output_type = "excel_and_pdf"
        elif pdf_path:
            output_type = "pdf"
        else:
            output_type = "excel"
        
        results.append(SplitResult(
            key=key,
            excel_path=excel_path,
            pdf_path=pdf_path,
            output_file_type=output_type,
        ))
    
    return results
