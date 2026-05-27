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
