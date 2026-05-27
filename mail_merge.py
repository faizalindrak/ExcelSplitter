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
