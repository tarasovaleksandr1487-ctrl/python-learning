#!/usr/bin/env python3
"""
Analyze rent and utility payment SMS/iMessage text from a safe copy of
macOS Messages.

Default source database:
    ~/Library/Messages/chat.db

Default copied database:
    /Users/alex/Projects/rent_sms_analysis/chat.db

Default workbook output:
    /Users/alex/Projects/rent_sms_analysis/rent_utilities_v4.xlsx

Usage:
    python3 rent_sms_analyzer.py --list-senders
    python3 rent_sms_analyzer.py
    python3 rent_sms_analyzer.py --sender "+48 500265786"

Safety:
    - Opens the original Messages database read-only.
    - Creates a separate SQLite backup copy first.
    - Reads only the copied database for analysis.
    - Does not delete, edit, or send any messages.

If macOS blocks access to Messages:
    1. Open System Settings.
    2. Go to Privacy & Security -> Full Disk Access.
    3. Add or enable Terminal. If you run this from another app, add that app.
    4. Quit and reopen Terminal, then run the script again.
"""

from __future__ import annotations

import argparse
import hashlib
import math
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import unicodedata
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Iterable, Sequence
from urllib.parse import quote
from xml.sax.saxutils import escape


DEFAULT_SOURCE_DB = "~/Library/Messages/chat.db"
DEFAULT_ANALYSIS_DIR = "/Users/alex/Projects/rent_sms_analysis"
DEFAULT_OUTPUT_NAME = "rent_utilities_v4.xlsx"
DEFAULT_SENDER = "+48 500265786"
DEFAULT_ATTACHMENTS_DIR_NAME = "attachments"
APPLE_EPOCH = datetime(2001, 1, 1, tzinfo=timezone.utc)
LOW_AMOUNT_PLN_THRESHOLD = 50
MIN_PAYMENT_CONFIDENCE = 50
MAX_UTILITY_AMOUNT = 5000
OCR_TEXT_LIMIT = 32000

BUNDLED_BIN_DIR = Path.home() / ".cache/codex-runtimes/codex-primary-runtime/dependencies/bin"
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".heic", ".heif", ".tif", ".tiff", ".bmp", ".gif", ".webp"}
PDF_EXTENSIONS = {".pdf"}
OCR_EXTENSIONS = IMAGE_EXTENSIONS | PDF_EXTENSIONS


FULL_DISK_ACCESS_HELP = """
macOS may be blocking access to Messages.

To grant Terminal Full Disk Access:
  1. Open System Settings.
  2. Go to Privacy & Security -> Full Disk Access.
  3. Add or enable Terminal. If you use iTerm, VS Code, PyCharm, or Codex
     to run this script, add that app too.
  4. Quit and reopen the app you granted access to, then run the script again.
"""


SERVICE_PATTERNS = [
    ("gas", re.compile(r"\b(?:gas|gaz\w*|pgnig)\b")),
    (
        "electricity",
        re.compile(
            r"\b(?:electricity|electric|power|energy|energia|energii|prad\w*|energa|tauron|pge)\b"
        ),
    ),
    ("water", re.compile(r"\b(?:water|woda|wode|wody|wodociag\w*)\b")),
    ("heating", re.compile(r"\b(?:heat|heating|ogrzew\w*|ciepl\w*)\b")),
    ("rent", re.compile(r"\b(?:rent|czynsz\w*|najem|wynajem|mieszkan\w*)\b")),
    (
        "trash",
        re.compile(r"\b(?:trash|garbage|waste|rubbish|smiec\w*|odpady|odpadow|wywoz\w*)\b"),
    ),
    (
        "internet",
        re.compile(r"\b(?:internet|wi\s*fi|wifi|router|swiatlowod\w*|orange|upc|play|netia)\b"),
    ),
    (
        "other services",
        re.compile(
            r"\b(?:media|utilities|utility|oplat\w*|uslug\w*|service\w*|maintenance|admin|administrac\w*|rachunek|rachunki)\b"
        ),
    ),
]

SERVICE_DISPLAY = {
    "gas": "Gas",
    "electricity": "Electricity",
    "water": "Water",
    "heating": "Heating",
    "rent": "Rent",
    "trash": "Trash",
    "internet": "Internet",
    "other services": "Other Services",
}

PAYMENT_CONTEXT_PATTERNS = [
    ("kwota do zaplaty", re.compile(r"\bkwota\s+do\s+zaplaty\b")),
    ("do zaplaty", re.compile(r"\bdo\s+zaplaty\b")),
    ("rachunek", re.compile(r"\brachun\w*\b")),
    ("faktura", re.compile(r"\bfaktur\w*\b")),
    ("oplata", re.compile(r"\boplat\w*\b")),
    ("naleznosc", re.compile(r"\bnalezn\w*\b")),
    ("kwota", re.compile(r"\bkwot\w*\b")),
    ("razem", re.compile(r"\brazem\b")),
    ("do uregulowania", re.compile(r"\bdo\s+uregulowania\b")),
]

PAYABLE_CONTEXT_PATTERNS = [
    ("kwota do zaplaty", re.compile(r"\bkwota\s+do\s+zaplaty\b")),
    ("do zaplaty", re.compile(r"\bdo\s+zaplaty\b")),
    ("do uregulowania", re.compile(r"\bdo\s+uregulowania\b")),
    ("naleznosc", re.compile(r"\bnalezn\w*\b")),
    ("razem", re.compile(r"\brazem\b")),
    ("PLN/zl", re.compile(r"\b(?:pln|zl|zloty|zlotych)\b")),
]

NON_PAYMENT_NUMBER_PATTERNS = [
    ("numer klienta", re.compile(r"\bnumer\s+klienta\b")),
    ("numer faktury", re.compile(r"\bnumer\s+faktur\w*\b")),
    ("numer licznika", re.compile(r"\bnumer\s+licznik\w*\b")),
    ("PPE", re.compile(r"\bppe\b")),
    ("PPG", re.compile(r"\bppg\b")),
    ("account number", re.compile(r"\baccount\s+number\b")),
    ("QR code", re.compile(r"\bqr\s+code\b|\bkod\s+qr\b")),
]

FEE_CONTEXT_PATTERNS = [
    ("fee", re.compile(r"\bfees?\b")),
    ("charge", re.compile(r"\bcharges?\b")),
    ("oplata", re.compile(r"\boplat\w*\b")),
]

CURRENCY_WORD = r"(?:pln|zl|zloty|zlotych|eur|euro|usd|\$|gbp|\u20ac|\u00a3)"
AMOUNT_RE = re.compile(
    rf"(?<![\w])(?P<currency_before>{CURRENCY_WORD})?\s*"
    rf"(?P<amount>(?:\d{{1,3}}(?:[ .]\d{{3}})+|\d+)(?:[,.]\d{{1,2}})?|\d+[,.]\d{{1,2}})"
    rf"\s*(?P<currency_after>{CURRENCY_WORD})?(?![\w])",
    re.IGNORECASE,
)

TRANSLATION = {
    ord("\u0142"): "l",
    ord("\u0141"): "l",
    ord("\u00a0"): " ",
}


@dataclass(frozen=True)
class MessageRow:
    message_id: int
    date: datetime | None
    direction: str
    handle_id: str
    handle_service: str
    chat_display_names: str
    chat_identifiers: str
    text: str


@dataclass(frozen=True)
class AttachmentRow:
    attachment_id: int
    message_id: int
    message_date: datetime | None
    direction: str
    sender: str
    transfer_name: str
    mime_type: str
    uti: str
    total_bytes: int | None
    original_filename: str
    resolved_path: str
    copied_path: str
    is_invoice_related: bool
    copy_status: str
    message_text: str


@dataclass(frozen=True)
class ServiceHit:
    service: str
    start: int
    end: int


@dataclass(frozen=True)
class AmountHit:
    amount: float
    amount_text: str
    currency: str | None
    currency_source: str
    start: int
    end: int
    payment_context_labels: tuple[str, ...]
    fee_context_labels: tuple[str, ...]


@dataclass(frozen=True)
class ScoredAmount:
    hit: AmountHit
    services: tuple[str, ...]
    confidence_score: int
    reason: str


@dataclass(frozen=True)
class PaymentRow:
    message_date: datetime | None
    service_type: str
    amount: float
    currency: str
    currency_source: str
    confidence_score: int
    extraction_reason: str
    source_type: str
    direction: str
    sender: str
    message_id: int
    original_text: str
    attachment_id: int | None = None
    attachment_path: str = ""
    invoice_date: str = ""
    invoice_number: str = ""


@dataclass(frozen=True)
class OcrResultRow:
    attachment_id: int
    message_id: int
    message_date: datetime | None
    source_path: str
    copied_path: str
    file_type: str
    ocr_status: str
    invoice_date: str
    invoice_number: str
    utility_type: str
    amount: float | None
    currency: str
    confidence_score: int
    extraction_reason: str
    ocr_text: str


@dataclass(frozen=True)
class ValidationRow:
    message_date: datetime | None
    message_id: int
    original_text: str
    extracted_amount: float | None
    currency: str
    service_type: str
    confidence_score: int
    reason: str


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    source_db = Path(args.source_db).expanduser()
    analysis_dir = Path(args.analysis_dir).expanduser()
    copied_db = analysis_dir / "chat.db"
    attachments_dir = Path(args.attachments_dir).expanduser()
    output_path = Path(args.output).expanduser() if args.output else analysis_dir / DEFAULT_OUTPUT_NAME

    sender = args.sender or DEFAULT_SENDER

    try:
        copy_messages_database(source_db, copied_db)
        with open_sqlite_readonly(copied_db) as conn:
            messages = fetch_messages(conn)
            attachments = fetch_attachments(conn)

        if args.list_senders:
            print_sender_list(messages, args.limit)
            print(f"\nCopied database: {copied_db}")
            return 0

        matched_messages = [
            message
            for message in messages
            if matches_sender(message, sender or "")
            and (args.include_my_replies or message.direction == "Incoming")
        ]

        matched_attachments = [
            attachment
            for attachment in attachments
            if matches_sender(attachment_message_row(attachment), sender)
            and (args.include_my_replies or attachment.direction == "Incoming")
        ]
        copied_attachments = copy_invoice_attachments(matched_attachments, attachments_dir)
        ocr_results = [
            perform_ocr_for_attachment(
                attachment,
                default_currency=args.default_currency,
                max_pdf_pages=args.ocr_max_pages,
            )
            for attachment in copied_attachments
            if attachment.is_invoice_related and is_ocr_candidate(attachment)
        ]

        payments = []
        validations = []
        for message in matched_messages:
            payment, validation = extract_payment_from_message(
                message,
                default_currency=args.default_currency,
                service_window=args.service_window,
            )
            validations.append(validation)
            if payment:
                payments.append(payment)

        payments.extend(ocr_payments_from_results(ocr_results, copied_attachments, args.default_currency))

        sheets = build_workbook_sheets(matched_messages, copied_attachments, ocr_results, payments)
        write_xlsx(output_path, sheets)

        print(f"Copied Messages database to: {copied_db}")
        print(f"Matched messages: {len(matched_messages)}")
        print(f"Matched attachments: {len(copied_attachments)}")
        print(f"OCR results: {len(ocr_results)}")
        print(f"Extracted payment rows: {len(payments)}")
        print(f"Wrote workbook to: {output_path}")
        print("No messages were deleted, edited, or sent.")
        return 0
    except PermissionError as exc:
        print(f"Permission error: {exc}", file=sys.stderr)
        print(FULL_DISK_ACCESS_HELP, file=sys.stderr)
        return 1
    except FileNotFoundError as exc:
        print(f"File not found: {exc}", file=sys.stderr)
        return 1
    except sqlite3.OperationalError as exc:
        print(f"SQLite error: {exc}", file=sys.stderr)
        if "unable to open" in str(exc).lower() or "authorization" in str(exc).lower():
            print(FULL_DISK_ACCESS_HELP, file=sys.stderr)
        return 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Safely analyze rent and utility payment texts from a copy of macOS Messages."
    )
    parser.add_argument(
        "--sender",
        default=DEFAULT_SENDER,
        help="Landlord phone number, handle, or visible chat/contact text to match.",
    )
    parser.add_argument(
        "--list-senders",
        action="store_true",
        help="List likely message senders/chats from the copied database, then exit.",
    )
    parser.add_argument(
        "--include-my-replies",
        action="store_true",
        help="Include your outgoing replies in the analysis as well as incoming landlord messages.",
    )
    parser.add_argument("--source-db", default=DEFAULT_SOURCE_DB, help="Path to original chat.db.")
    parser.add_argument(
        "--analysis-dir",
        default=DEFAULT_ANALYSIS_DIR,
        help="Folder where the copied chat.db and workbook should be saved.",
    )
    parser.add_argument(
        "--attachments-dir",
        default=str(Path(DEFAULT_ANALYSIS_DIR) / DEFAULT_ATTACHMENTS_DIR_NAME),
        help="Folder where invoice-related attachments should be copied.",
    )
    parser.add_argument("--output", help="Workbook output path. Defaults to rent_utilities_v4.xlsx.")
    parser.add_argument(
        "--default-currency",
        default="PLN",
        help="Currency to assume when a nearby amount has no explicit currency. Default: PLN.",
    )
    parser.add_argument(
        "--service-window",
        type=int,
        default=90,
        help="Characters between a service keyword and amount that still count as related.",
    )
    parser.add_argument(
        "--ocr-max-pages",
        type=int,
        default=4,
        help="Maximum PDF pages to OCR per attachment. Default: 4.",
    )
    parser.add_argument("--limit", type=int, default=50, help="Rows to show with --list-senders.")
    return parser


def copy_messages_database(source_db: Path, copied_db: Path) -> None:
    if not source_db.exists():
        raise FileNotFoundError(f"Messages database not found: {source_db}")

    copied_db.parent.mkdir(parents=True, exist_ok=True)
    temp_db = copied_db.with_name(f"{copied_db.name}.tmp")
    if temp_db.exists():
        temp_db.unlink()

    source_uri = sqlite_uri(source_db, mode="ro")
    source_conn = sqlite3.connect(source_uri, uri=True)
    try:
        dest_conn = sqlite3.connect(str(temp_db))
        try:
            source_conn.backup(dest_conn)
        finally:
            dest_conn.close()
    finally:
        source_conn.close()

    os.replace(temp_db, copied_db)
    try:
        os.chmod(copied_db, 0o600)
    except OSError:
        pass


def sqlite_uri(path: Path, mode: str) -> str:
    return f"file:{quote(str(path), safe='/')}?mode={mode}"


def open_sqlite_readonly(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(sqlite_uri(db_path, mode="ro"), uri=True)
    conn.row_factory = sqlite3.Row
    return conn


def fetch_messages(conn: sqlite3.Connection) -> list[MessageRow]:
    message_columns = table_columns(conn, "message")
    text_expr = "m.text" if "text" in message_columns else "NULL"
    attributed_expr = "m.attributedBody" if "attributedBody" in message_columns else "NULL"

    query = f"""
        SELECT
            m.ROWID AS message_id,
            m.date AS message_date,
            m.is_from_me AS is_from_me,
            {text_expr} AS text,
            {attributed_expr} AS attributed_body,
            COALESCE(h.id, '') AS handle_id,
            COALESCE(h.service, '') AS handle_service,
            COALESCE(GROUP_CONCAT(DISTINCT c.display_name), '') AS chat_display_names,
            COALESCE(GROUP_CONCAT(DISTINCT c.chat_identifier), '') AS chat_identifiers
        FROM message m
        LEFT JOIN handle h ON h.ROWID = m.handle_id
        LEFT JOIN chat_message_join cmj ON cmj.message_id = m.ROWID
        LEFT JOIN chat c ON c.ROWID = cmj.chat_id
        GROUP BY m.ROWID
        ORDER BY m.date ASC, m.ROWID ASC
    """

    rows = []
    for row in conn.execute(query):
        text = clean_message_text(row["text"]) or decode_attributed_body(row["attributed_body"])
        rows.append(
            MessageRow(
                message_id=int(row["message_id"]),
                date=convert_apple_date(row["message_date"]),
                direction="Outgoing" if row["is_from_me"] else "Incoming",
                handle_id=row["handle_id"] or "",
                handle_service=row["handle_service"] or "",
                chat_display_names=row["chat_display_names"] or "",
                chat_identifiers=row["chat_identifiers"] or "",
                text=text,
            )
        )
    return rows


def fetch_attachments(conn: sqlite3.Connection) -> list[AttachmentRow]:
    attachment_columns = table_columns(conn, "attachment")
    message_columns = table_columns(conn, "message")

    filename_expr = "a.filename" if "filename" in attachment_columns else "NULL"
    transfer_expr = "a.transfer_name" if "transfer_name" in attachment_columns else "NULL"
    mime_expr = "a.mime_type" if "mime_type" in attachment_columns else "NULL"
    uti_expr = "a.uti" if "uti" in attachment_columns else "NULL"
    total_bytes_expr = "a.total_bytes" if "total_bytes" in attachment_columns else "NULL"
    text_expr = "m.text" if "text" in message_columns else "NULL"
    attributed_expr = "m.attributedBody" if "attributedBody" in message_columns else "NULL"

    query = f"""
        SELECT
            a.ROWID AS attachment_id,
            maj.message_id AS message_id,
            m.date AS message_date,
            m.is_from_me AS is_from_me,
            {text_expr} AS text,
            {attributed_expr} AS attributed_body,
            COALESCE(h.id, '') AS handle_id,
            {filename_expr} AS filename,
            {transfer_expr} AS transfer_name,
            {mime_expr} AS mime_type,
            {uti_expr} AS uti,
            {total_bytes_expr} AS total_bytes
        FROM message_attachment_join maj
        JOIN attachment a ON a.ROWID = maj.attachment_id
        JOIN message m ON m.ROWID = maj.message_id
        LEFT JOIN handle h ON h.ROWID = m.handle_id
        ORDER BY m.date ASC, m.ROWID ASC, a.ROWID ASC
    """

    rows = []
    for row in conn.execute(query):
        original_filename = clean_message_text(row["filename"])
        transfer_name = clean_message_text(row["transfer_name"])
        mime_type = clean_message_text(row["mime_type"])
        uti = clean_message_text(row["uti"])
        message_text = clean_message_text(row["text"]) or decode_attributed_body(row["attributed_body"])
        resolved_path = resolve_messages_attachment_path(original_filename)
        rows.append(
            AttachmentRow(
                attachment_id=int(row["attachment_id"]),
                message_id=int(row["message_id"]),
                message_date=convert_apple_date(row["message_date"]),
                direction="Outgoing" if row["is_from_me"] else "Incoming",
                sender=row["handle_id"] or "",
                transfer_name=transfer_name,
                mime_type=mime_type,
                uti=uti,
                total_bytes=int(row["total_bytes"]) if row["total_bytes"] is not None else None,
                original_filename=original_filename,
                resolved_path=str(resolved_path) if resolved_path else "",
                copied_path="",
                is_invoice_related=is_invoice_related_attachment(
                    filename=original_filename,
                    transfer_name=transfer_name,
                    mime_type=mime_type,
                    uti=uti,
                    message_text=message_text,
                ),
                copy_status="Not copied",
                message_text=message_text,
            )
        )
    return rows


def attachment_message_row(attachment: AttachmentRow) -> MessageRow:
    return MessageRow(
        message_id=attachment.message_id,
        date=attachment.message_date,
        direction=attachment.direction,
        handle_id=attachment.sender,
        handle_service="",
        chat_display_names="",
        chat_identifiers="",
        text=attachment.message_text,
    )


def resolve_messages_attachment_path(filename: str) -> Path | None:
    if not filename:
        return None
    path_text = filename.replace("file://", "")
    path = Path(path_text).expanduser()
    if path.is_absolute():
        return path
    return Path.home() / "Library" / "Messages" / path


def is_invoice_related_attachment(
    filename: str,
    transfer_name: str,
    mime_type: str,
    uti: str,
    message_text: str,
) -> bool:
    searchable = normalize_for_match(" ".join([filename, transfer_name, mime_type, uti, message_text]))
    if is_probable_image_or_pdf(filename, transfer_name, mime_type, uti):
        return True
    invoice_words = (
        "invoice",
        "faktura",
        "rachunek",
        "oplata",
        "naleznosc",
        "gaz",
        "woda",
        "prad",
        "energia",
        "ogrzew",
    )
    return any(re.search(rf"\b{word}\w*\b", searchable) for word in invoice_words)


def is_probable_image_or_pdf(filename: str, transfer_name: str, mime_type: str, uti: str) -> bool:
    suffixes = [Path(value).suffix.casefold() for value in (filename, transfer_name) if value]
    if any(suffix in OCR_EXTENSIONS for suffix in suffixes):
        return True
    mime_norm = mime_type.casefold()
    uti_norm = uti.casefold()
    return (
        mime_norm.startswith("image/")
        or mime_norm == "application/pdf"
        or "public.image" in uti_norm
        or "com.adobe.pdf" in uti_norm
    )


def is_ocr_candidate(attachment: AttachmentRow) -> bool:
    return bool(attachment.copied_path) and is_probable_image_or_pdf(
        attachment.original_filename,
        attachment.transfer_name,
        attachment.mime_type,
        attachment.uti,
    )


def copy_invoice_attachments(
    attachments: Sequence[AttachmentRow],
    attachments_dir: Path,
) -> list[AttachmentRow]:
    attachments_dir.mkdir(parents=True, exist_ok=True)
    copied = []
    for attachment in attachments:
        if not attachment.is_invoice_related:
            copied.append(attachment)
            continue

        source_path = Path(attachment.resolved_path).expanduser() if attachment.resolved_path else None
        if not source_path or not source_path.exists():
            copied.append(replace_attachment_copy_fields(attachment, "", "Source file not found"))
            continue

        destination = unique_attachment_destination(attachment, source_path, attachments_dir)
        try:
            shutil.copy2(source_path, destination)
            copied.append(
                replace_attachment_copy_fields(attachment, str(destination), "Copied")
            )
        except OSError as exc:
            copied.append(replace_attachment_copy_fields(attachment, "", f"Copy failed: {exc}"))
    return copied


def replace_attachment_copy_fields(
    attachment: AttachmentRow,
    copied_path: str,
    copy_status: str,
) -> AttachmentRow:
    return AttachmentRow(
        attachment_id=attachment.attachment_id,
        message_id=attachment.message_id,
        message_date=attachment.message_date,
        direction=attachment.direction,
        sender=attachment.sender,
        transfer_name=attachment.transfer_name,
        mime_type=attachment.mime_type,
        uti=attachment.uti,
        total_bytes=attachment.total_bytes,
        original_filename=attachment.original_filename,
        resolved_path=attachment.resolved_path,
        copied_path=copied_path,
        is_invoice_related=attachment.is_invoice_related,
        copy_status=copy_status,
        message_text=attachment.message_text,
    )


def unique_attachment_destination(
    attachment: AttachmentRow,
    source_path: Path,
    attachments_dir: Path,
) -> Path:
    display_name = attachment.transfer_name or source_path.name or f"attachment_{attachment.attachment_id}"
    safe_name = safe_filename(display_name)
    if not Path(safe_name).suffix and source_path.suffix:
        safe_name += source_path.suffix
    digest = hashlib.sha1(f"{attachment.message_id}-{attachment.attachment_id}-{source_path}".encode()).hexdigest()[:10]
    return attachments_dir / f"msg{attachment.message_id}_att{attachment.attachment_id}_{digest}_{safe_name}"


def safe_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._ -]+", "_", value).strip(" .")
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned[:160] or "attachment"


def table_columns(conn: sqlite3.Connection, table_name: str) -> set[str]:
    return {row["name"] for row in conn.execute(f"PRAGMA table_info({table_name})")}


def clean_message_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        value = value.decode("utf-8", errors="ignore")
    text = str(value).replace("\ufffc", "").strip()
    return normalize_spaces(text)


def decode_attributed_body(value: object) -> str:
    if not value:
        return ""
    if not isinstance(value, bytes):
        return clean_message_text(value)

    decoded = value.decode("utf-8", errors="ignore")
    decoded = re.sub(r"[\x00-\x08\x0b-\x1f\x7f]+", "\n", decoded)
    chunks = [normalize_spaces(chunk) for chunk in decoded.splitlines()]
    chunks = [chunk for chunk in chunks if len(chunk) >= 2]

    bad_tokens = (
        "streamtyped",
        "NSString",
        "NSObject",
        "NSDictionary",
        "NSMutable",
        "NSNumber",
        "NSAttributed",
        "__kIM",
    )

    candidates = []
    for chunk in chunks:
        if any(token in chunk for token in bad_tokens):
            continue
        printable_ratio = sum(1 for char in chunk if char.isprintable()) / max(len(chunk), 1)
        if printable_ratio >= 0.8:
            candidates.append(chunk)

    if not candidates:
        return ""

    def score(chunk: str) -> int:
        bonus = 0
        if re.search(r"\d", chunk):
            bonus += 8
        if re.search(CURRENCY_WORD, normalize_for_match(chunk), re.IGNORECASE):
            bonus += 12
        if " " in chunk:
            bonus += 4
        return len(chunk) + bonus

    return max(candidates, key=score).strip()


def normalize_spaces(text: str) -> str:
    return re.sub(r"[ \t\r\f\v]+", " ", text).strip()


def convert_apple_date(value: object) -> datetime | None:
    if value is None:
        return None
    try:
        raw = int(value)
    except (TypeError, ValueError):
        return None

    absolute = abs(raw)
    if absolute > 10_000_000_000_000:
        seconds = raw / 1_000_000_000
    elif absolute > 10_000_000_000:
        seconds = raw / 1_000_000
    else:
        seconds = raw

    try:
        return (APPLE_EPOCH + timedelta(seconds=seconds)).astimezone()
    except (OverflowError, OSError):
        return None


def print_sender_list(messages: Sequence[MessageRow], limit: int) -> None:
    grouped = {}
    for message in messages:
        key = (
            message.handle_id,
            message.handle_service,
            message.chat_display_names,
            message.chat_identifiers,
        )
        stats = grouped.setdefault(
            key,
            {
                "count": 0,
                "incoming": 0,
                "outgoing": 0,
                "first": message.date,
                "last": message.date,
            },
        )
        stats["count"] += 1
        stats["incoming" if message.direction == "Incoming" else "outgoing"] += 1
        if message.date and (stats["first"] is None or message.date < stats["first"]):
            stats["first"] = message.date
        if message.date and (stats["last"] is None or message.date > stats["last"]):
            stats["last"] = message.date

    sorted_items = sorted(grouped.items(), key=lambda item: item[1]["count"], reverse=True)
    print("Potential senders/chats:")
    for index, (key, stats) in enumerate(sorted_items[:limit], start=1):
        handle_id, handle_service, chat_names, chat_ids = key
        label_parts = [part for part in (handle_id, chat_names, chat_ids) if part]
        label = " | ".join(label_parts) if label_parts else "(unknown)"
        service = f" [{handle_service}]" if handle_service else ""
        first = format_datetime(stats["first"])
        last = format_datetime(stats["last"])
        print(
            f"{index:>3}. {label}{service} - {stats['count']} messages "
            f"({stats['incoming']} incoming, {stats['outgoing']} outgoing), {first} to {last}"
        )


def matches_sender(message: MessageRow, sender_query: str) -> bool:
    query = sender_query.strip()
    if not query:
        return False

    query_norm = normalize_for_match(query)
    query_digits = digits_only(query)
    fields = [
        message.handle_id,
        message.handle_service,
        message.chat_display_names,
        message.chat_identifiers,
    ]

    for field in fields:
        field_norm = normalize_for_match(field)
        if query_norm and query_norm in field_norm:
            return True

        field_digits = digits_only(field)
        if len(query_digits) >= 4 and field_digits:
            if query_digits in field_digits or field_digits.endswith(query_digits):
                return True
            if len(field_digits) >= 4 and field_digits in query_digits:
                return True

    return False


def digits_only(text: object) -> str:
    return re.sub(r"\D", "", str(text or ""))


def normalize_for_match(text: object) -> str:
    value = str(text or "").casefold().translate(TRANSLATION)
    value = unicodedata.normalize("NFKD", value)
    value = "".join(char for char in value if not unicodedata.combining(char))
    return normalize_spaces(value)


def extract_payment_from_message(
    message: MessageRow,
    default_currency: str,
    service_window: int,
) -> tuple[PaymentRow | None, ValidationRow]:
    original_text = message.text
    normalized_text = normalize_for_match(original_text)
    service_hits = find_service_hits(normalized_text)
    message_payment_labels = labels_near_range(
        normalized_text,
        0,
        len(normalized_text),
        PAYMENT_CONTEXT_PATTERNS,
        window=0,
    )
    if not service_hits and not message_payment_labels:
        return no_payment_result(
            message,
            "No utility/rent service keyword or payment keyword found.",
        )

    amount_hits, rejected_amounts = find_amount_hits(normalized_text, default_currency=default_currency)
    if not amount_hits:
        detail = summarize_rejections(rejected_amounts)
        reason = "No accepted amount candidates."
        if detail:
            reason = f"{reason} {detail}"
        return no_payment_result(message, reason)

    all_services = unique_ordered([hit.service for hit in service_hits])
    if not all_services and message_payment_labels:
        all_services = ["other services"]

    scored_amounts = []
    for amount_hit in amount_hits:
        selected_services = choose_services_for_amount(
            amount_hit=amount_hit,
            service_hits=service_hits,
            all_services=all_services,
            amount_count=len(amount_hits),
            service_window=service_window,
        )
        if not selected_services and amount_hit.payment_context_labels:
            selected_services = ["other services"]
        if not selected_services:
            continue

        scored_amounts.append(
            score_amount_candidate(
                amount_hit=amount_hit,
                selected_services=selected_services,
                default_currency=default_currency,
            )
        )

    scored_amounts = [
        scored for scored in scored_amounts if scored.confidence_score >= MIN_PAYMENT_CONFIDENCE
    ]
    if not scored_amounts:
        detail = summarize_rejections(rejected_amounts)
        reason = "Amount candidates were present, but none had enough payment context."
        if detail:
            reason = f"{reason} Rejected candidates: {detail}"
        return no_payment_result(message, reason)

    best = max(
        scored_amounts,
        key=lambda scored: (
            scored.confidence_score,
            scored.hit.currency is not None,
            scored.hit.amount,
        ),
    )
    service_label = display_service_label(best.services)
    currency = best.hit.currency or default_currency.upper()
    extraction_reason = best.reason
    ignored_detail = summarize_rejections(rejected_amounts, limit=4)
    if ignored_detail:
        extraction_reason = f"{extraction_reason} Ignored other numbers: {ignored_detail}."

    payment = PaymentRow(
        message_date=message.date,
        service_type=service_label,
        amount=round(best.hit.amount, 2),
        currency=currency,
        currency_source=best.hit.currency_source,
        confidence_score=best.confidence_score,
        extraction_reason=extraction_reason,
        source_type="Message",
        direction=message.direction,
        sender=message.handle_id,
        message_id=message.message_id,
        original_text=original_text,
    )
    validation = ValidationRow(
        message_date=message.date,
        message_id=message.message_id,
        original_text=original_text,
        extracted_amount=payment.amount,
        currency=payment.currency,
        service_type=payment.service_type,
        confidence_score=payment.confidence_score,
        reason=payment.extraction_reason,
    )
    return payment, validation


def no_payment_result(message: MessageRow, reason: str) -> tuple[None, ValidationRow]:
    return None, ValidationRow(
        message_date=message.date,
        message_id=message.message_id,
        original_text=message.text,
        extracted_amount=None,
        currency="",
        service_type="",
        confidence_score=0,
        reason=reason,
    )


def score_amount_candidate(
    amount_hit: AmountHit,
    selected_services: Sequence[str],
    default_currency: str,
) -> ScoredAmount:
    score = 0
    reasons = []

    currency = amount_hit.currency or default_currency.upper()
    if amount_hit.currency:
        score += 45
        reasons.append(f"explicit {currency} currency")
    else:
        score += 5
        reasons.append(f"{default_currency.upper()} assumed")

    if amount_hit.payment_context_labels:
        keyword_bonus = 35
        if "do zaplaty" in amount_hit.payment_context_labels:
            keyword_bonus += 20
        score += keyword_bonus
        reasons.append(
            "near payment keyword "
            + ", ".join(f"'{label}'" for label in amount_hit.payment_context_labels)
        )

    specific_services = [service for service in selected_services if service != "other services"]
    if specific_services:
        score += 25
        reasons.append(
            "near service keyword "
            + ", ".join(f"'{SERVICE_DISPLAY.get(service, service)}'" for service in specific_services)
        )
    elif selected_services:
        score += 10
        reasons.append("classified as Other Services from payment wording")

    if currency == "PLN" and amount_hit.amount < LOW_AMOUNT_PLN_THRESHOLD:
        score += 10
        reasons.append("below 50 PLN but explicitly marked as a fee")
    elif amount_hit.amount >= LOW_AMOUNT_PLN_THRESHOLD:
        score += 10
        reasons.append("amount is at least 50")

    if amount_hit.fee_context_labels:
        score += 10

    confidence_score = max(0, min(100, score))
    reason = "Selected as the highest-confidence amount: " + "; ".join(reasons) + "."
    return ScoredAmount(
        hit=amount_hit,
        services=tuple(selected_services),
        confidence_score=confidence_score,
        reason=reason,
    )


def find_service_hits(normalized_text: str) -> list[ServiceHit]:
    hits = []
    for service, pattern in SERVICE_PATTERNS:
        for match in pattern.finditer(normalized_text):
            hits.append(ServiceHit(service=service, start=match.start(), end=match.end()))
    return sorted(hits, key=lambda hit: (hit.start, hit.end))


def labels_near_range(
    normalized_text: str,
    start: int,
    end: int,
    patterns: Sequence[tuple[str, re.Pattern[str]]],
    window: int,
) -> tuple[str, ...]:
    context_start = max(0, start - window)
    context_end = min(len(normalized_text), end + window)
    context = normalized_text[context_start:context_end]
    labels = []
    for label, pattern in patterns:
        if pattern.search(context):
            labels.append(label)
    return tuple(unique_ordered(labels))


def text_context(
    text: str,
    start: int,
    end: int,
    window: int,
) -> str:
    return text[max(0, start - window) : min(len(text), end + window)]


def has_non_payment_number_context(normalized_text: str, start: int, end: int) -> tuple[str, ...]:
    non_payment = nearest_context_labels(
        normalized_text=normalized_text,
        start=start,
        end=end,
        patterns=NON_PAYMENT_NUMBER_PATTERNS,
        window=70,
    )
    if not non_payment:
        return ()

    payable = nearest_context_labels(
        normalized_text=normalized_text,
        start=start,
        end=end,
        patterns=PAYABLE_CONTEXT_PATTERNS,
        window=90,
    )
    payable_distance = min((distance for _, distance in payable), default=10_000)
    return tuple(
        label
        for label, distance in non_payment
        if distance <= 35 and distance <= payable_distance
    )


def nearest_context_labels(
    normalized_text: str,
    start: int,
    end: int,
    patterns: Sequence[tuple[str, re.Pattern[str]]],
    window: int,
) -> tuple[tuple[str, int], ...]:
    context_start = max(0, start - window)
    context_end = min(len(normalized_text), end + window)
    context = normalized_text[context_start:context_end]
    amount_start = start - context_start
    amount_end = end - context_start
    labels = []
    for label, pattern in patterns:
        best_distance = None
        for match in pattern.finditer(context):
            distance = distance_between_ranges(amount_start, amount_end, match.start(), match.end())
            if best_distance is None or distance < best_distance:
                best_distance = distance
        if best_distance is not None:
            labels.append((label, best_distance))
    labels.sort(key=lambda item: item[1])
    return tuple(labels)


def summarize_rejections(rejections: Sequence[str], limit: int = 6) -> str:
    if not rejections:
        return ""
    shown = list(rejections[:limit])
    extra_count = len(rejections) - len(shown)
    detail = "; ".join(shown)
    if extra_count > 0:
        detail = f"{detail}; plus {extra_count} more"
    return detail


def find_amount_hits(normalized_text: str, default_currency: str) -> tuple[list[AmountHit], list[str]]:
    hits = []
    rejected = []
    for match in AMOUNT_RE.finditer(normalized_text):
        before = match.group("currency_before")
        after = match.group("currency_after")
        currency_token = before or after
        currency = normalize_currency(currency_token)
        amount_text = match.group("amount")
        amount = parse_amount(amount_text)
        if amount is None:
            rejected.append(f"{amount_text}: could not parse as a number")
            continue
        payment_context_labels = labels_near_range(
            normalized_text=normalized_text,
            start=match.start(),
            end=match.end(),
            patterns=PAYMENT_CONTEXT_PATTERNS,
            window=80,
        )
        fee_context_labels = labels_near_range(
            normalized_text=normalized_text,
            start=match.start(),
            end=match.end(),
            patterns=FEE_CONTEXT_PATTERNS,
            window=80,
        )
        rejection_reason = amount_rejection_reason(
            normalized_text=normalized_text,
            amount_text=amount_text,
            amount=amount,
            currency=currency,
            default_currency=default_currency,
            start=match.start(),
            end=match.end(),
            fee_context_labels=fee_context_labels,
        )
        if rejection_reason:
            rejected.append(f"{amount_text}: {rejection_reason}")
            continue
        hits.append(
            AmountHit(
                amount=amount,
                amount_text=amount_text,
                currency=currency,
                currency_source="explicit" if currency else f"assumed {default_currency.upper()}",
                start=match.start(),
                end=match.end(),
                payment_context_labels=payment_context_labels,
                fee_context_labels=fee_context_labels,
            )
        )
    return hits, rejected


def normalize_currency(token: str | None) -> str | None:
    if not token:
        return None
    token_norm = normalize_for_match(token)
    if token_norm in {"pln", "zl", "zloty", "zlotych"}:
        return "PLN"
    if token_norm in {"eur", "euro", "\u20ac"}:
        return "EUR"
    if token_norm in {"usd", "$"}:
        return "USD"
    if token_norm in {"gbp", "\u00a3"}:
        return "GBP"
    return token_norm.upper()


def parse_amount(amount_text: str) -> float | None:
    value = amount_text.replace(" ", "").replace("\u00a0", "")
    if not value:
        return None

    comma = value.rfind(",")
    dot = value.rfind(".")
    if comma >= 0 and dot >= 0:
        if comma > dot:
            value = value.replace(".", "").replace(",", ".")
        else:
            value = value.replace(",", "")
    elif comma >= 0:
        integer, _, fraction = value.partition(",")
        value = integer + "." + fraction if len(fraction) <= 2 else integer + fraction
    elif dot >= 0:
        integer, _, fraction = value.partition(".")
        value = integer + fraction if len(fraction) == 3 else value

    try:
        amount = float(value)
    except ValueError:
        return None
    if not math.isfinite(amount):
        return None
    return amount


def amount_rejection_reason(
    normalized_text: str,
    amount_text: str,
    amount: float,
    currency: str | None,
    default_currency: str,
    start: int,
    end: int,
    fee_context_labels: Sequence[str],
) -> str:
    if amount <= 0 or amount > 1_000_000:
        return "outside expected payment range"

    digits = digits_only(amount_text)
    explicit_currency = currency is not None
    effective_currency = currency or default_currency.upper()
    non_payment_labels = has_non_payment_number_context(normalized_text, start, end)
    if non_payment_labels:
        return "near non-payment identifier " + ", ".join(non_payment_labels)
    if effective_currency == "PLN" and amount > MAX_UTILITY_AMOUNT:
        return f"above {MAX_UTILITY_AMOUNT} PLN utility limit"
    if not explicit_currency:
        if is_standalone_year(amount_text, amount):
            return "ignored as standalone year"
        if len(digits) >= 7:
            return "too many digits for an unmarked payment amount"
        if looks_like_day_month(amount_text):
            return "looks like a day/month date"

    if effective_currency == "PLN" and amount < LOW_AMOUNT_PLN_THRESHOLD and not fee_context_labels:
        return "below 50 PLN and not explicitly marked as a fee"

    context = normalized_text[max(0, start - 24) : min(len(normalized_text), end + 24)]
    account_words = ("konto", "account", "iban", "nr", "numer", "tel", "phone", "kod", "code")
    if not explicit_currency and any(re.search(rf"\b{word}\b", context) for word in account_words):
        return "looks like an account, phone, or code number"

    return ""


def is_standalone_year(amount_text: str, amount: float) -> bool:
    return re.fullmatch(r"\d{4}", amount_text) is not None and 2000 <= int(amount) <= 2100


def looks_like_day_month(amount_text: str) -> bool:
    match = re.fullmatch(r"(\d{1,2})\.(\d{1,2})", amount_text)
    if not match:
        return False
    left = int(match.group(1))
    right = int(match.group(2))
    return 1 <= left <= 31 and 1 <= right <= 12


def choose_services_for_amount(
    amount_hit: AmountHit,
    service_hits: Sequence[ServiceHit],
    all_services: Sequence[str],
    amount_count: int,
    service_window: int,
) -> list[str]:
    nearby = [
        hit
        for hit in service_hits
        if distance_between_ranges(amount_hit.start, amount_hit.end, hit.start, hit.end) <= service_window
    ]

    if nearby:
        nearby = sorted(
            nearby,
            key=lambda hit: distance_between_ranges(amount_hit.start, amount_hit.end, hit.start, hit.end),
        )
        if amount_count == 1:
            text_order = sorted(nearby, key=lambda hit: (hit.start, hit.end))
            return prefer_specific_services(unique_ordered([hit.service for hit in text_order]))
        return prefer_specific_services([nearby[0].service])

    if len(all_services) == 1:
        return list(all_services)

    if amount_count == 1 and all_services:
        return prefer_specific_services(list(all_services))

    return []


def distance_between_ranges(start_a: int, end_a: int, start_b: int, end_b: int) -> int:
    if end_a < start_b:
        return start_b - end_a
    if end_b < start_a:
        return start_a - end_b
    return 0


def prefer_specific_services(services: Sequence[str]) -> list[str]:
    unique = unique_ordered(services)
    specific = [service for service in unique if service != "other services"]
    return specific or unique


def unique_ordered(values: Iterable[str]) -> list[str]:
    seen = set()
    result = []
    for value in values:
        if value not in seen:
            seen.add(value)
            result.append(value)
    return result


def display_service_label(services: Sequence[str]) -> str:
    return " + ".join(SERVICE_DISPLAY.get(service, service.title()) for service in services)


def perform_ocr_for_attachment(
    attachment: AttachmentRow,
    default_currency: str,
    max_pdf_pages: int = 4,
) -> OcrResultRow:
    copied_path = Path(attachment.copied_path)
    file_type = attachment_file_type(attachment)
    if not copied_path.exists():
        return empty_ocr_result(
            attachment,
            file_type=file_type,
            status="Copied attachment file not found",
        )

    if file_type == "PDF":
        ocr_text, status = extract_pdf_text(copied_path, max_pages=max_pdf_pages)
    elif file_type == "Image":
        ocr_text, status = ocr_image(copied_path)
    else:
        return empty_ocr_result(
            attachment,
            file_type=file_type,
            status="Not an OCR-supported image or PDF",
        )

    invoice_date = extract_invoice_date(ocr_text)
    invoice_number = extract_invoice_number(ocr_text)
    utility_type = extract_utility_type(ocr_text)
    amount, currency, confidence_score, extraction_reason = extract_best_payment_amount_from_text(
        ocr_text,
        default_currency=default_currency,
        utility_type=utility_type,
    )

    return OcrResultRow(
        attachment_id=attachment.attachment_id,
        message_id=attachment.message_id,
        message_date=attachment.message_date,
        source_path=attachment.resolved_path,
        copied_path=attachment.copied_path,
        file_type=file_type,
        ocr_status=status,
        invoice_date=invoice_date,
        invoice_number=invoice_number,
        utility_type=utility_type,
        amount=amount,
        currency=currency,
        confidence_score=confidence_score,
        extraction_reason=extraction_reason,
        ocr_text=ocr_text[:OCR_TEXT_LIMIT],
    )


def empty_ocr_result(attachment: AttachmentRow, file_type: str, status: str) -> OcrResultRow:
    return OcrResultRow(
        attachment_id=attachment.attachment_id,
        message_id=attachment.message_id,
        message_date=attachment.message_date,
        source_path=attachment.resolved_path,
        copied_path=attachment.copied_path,
        file_type=file_type,
        ocr_status=status,
        invoice_date="",
        invoice_number="",
        utility_type="",
        amount=None,
        currency="",
        confidence_score=0,
        extraction_reason=status,
        ocr_text="",
    )


def attachment_file_type(attachment: AttachmentRow) -> str:
    values = [attachment.copied_path, attachment.original_filename, attachment.transfer_name]
    suffixes = [Path(value).suffix.casefold() for value in values if value]
    mime = attachment.mime_type.casefold()
    uti = attachment.uti.casefold()
    if any(suffix in PDF_EXTENSIONS for suffix in suffixes) or mime == "application/pdf" or "pdf" in uti:
        return "PDF"
    if any(suffix in IMAGE_EXTENSIONS for suffix in suffixes) or mime.startswith("image/") or "image" in uti:
        return "Image"
    return "Other"


def extract_pdf_text(pdf_path: Path, max_pages: int) -> tuple[str, str]:
    text_parts = []
    try:
        from pypdf import PdfReader  # type: ignore

        reader = PdfReader(str(pdf_path))
        for page in reader.pages[:max_pages]:
            text_parts.append(page.extract_text() or "")
        text = normalize_ocr_text("\n".join(text_parts))
        if len(text) >= 30:
            return text, "Extracted embedded PDF text"
    except Exception as exc:
        text_parts.append(f"PDF text-layer extraction failed: {exc}")

    tesseract = find_command("tesseract")
    pdftoppm = find_command("pdftoppm")
    if not tesseract or not pdftoppm:
        missing = []
        if not tesseract:
            missing.append("tesseract")
        if not pdftoppm:
            missing.append("pdftoppm")
        return "", "OCR skipped: missing " + " and ".join(missing)

    try:
        with tempfile.TemporaryDirectory(prefix="rent_sms_pdf_ocr_") as temp_dir:
            prefix = str(Path(temp_dir) / "page")
            run_command(
                [
                    pdftoppm,
                    "-r",
                    "200",
                    "-png",
                    "-f",
                    "1",
                    "-l",
                    str(max_pages),
                    str(pdf_path),
                    prefix,
                ],
                timeout=180,
            )
            page_images = sorted(Path(temp_dir).glob("page-*.png"))
            if not page_images:
                return "", "OCR failed: PDF rendering produced no page images"

            ocr_pages = []
            for page_image in page_images:
                page_text, page_status = ocr_image(page_image)
                if page_text:
                    ocr_pages.append(page_text)
                elif page_status:
                    ocr_pages.append(f"[{page_status}]")
            return normalize_ocr_text("\n".join(ocr_pages)), "OCR completed from rendered PDF pages"
    except Exception as exc:
        return "", f"OCR failed: {exc}"


def ocr_image(image_path: Path) -> tuple[str, str]:
    tesseract = find_command("tesseract")
    if not tesseract:
        return "", "OCR skipped: tesseract not found"

    source_path = image_path
    cleanup_dir = None
    if image_path.suffix.casefold() in {".heic", ".heif", ".webp", ".gif"}:
        converted, cleanup_dir = convert_image_for_ocr(image_path)
        if converted:
            source_path = converted

    try:
        result = run_tesseract(tesseract, source_path, languages="pol+eng")
        if result.returncode != 0:
            result = run_tesseract(tesseract, source_path, languages=None)
        if result.returncode != 0:
            stderr = normalize_spaces(result.stderr.decode("utf-8", errors="ignore"))
            return "", f"OCR failed: {stderr or 'tesseract returned an error'}"
        return normalize_ocr_text(result.stdout.decode("utf-8", errors="ignore")), "OCR completed"
    finally:
        if cleanup_dir:
            shutil.rmtree(cleanup_dir, ignore_errors=True)


def convert_image_for_ocr(image_path: Path) -> tuple[Path | None, str | None]:
    sips = find_command("sips")
    if not sips:
        return None, None
    temp_dir = tempfile.mkdtemp(prefix="rent_sms_image_ocr_")
    converted = Path(temp_dir) / "converted.png"
    result = run_command(
        [sips, "-s", "format", "png", str(image_path), "--out", str(converted)],
        timeout=60,
        check=False,
    )
    if result.returncode == 0 and converted.exists():
        return converted, temp_dir
    shutil.rmtree(temp_dir, ignore_errors=True)
    return None, None


def run_tesseract(tesseract: str, image_path: Path, languages: str | None) -> subprocess.CompletedProcess[bytes]:
    command = [tesseract, str(image_path), "stdout", "--psm", "6"]
    if languages:
        command.extend(["-l", languages])
    return run_command(command, timeout=120, check=False)


def run_command(
    command: Sequence[str],
    timeout: int,
    check: bool = True,
) -> subprocess.CompletedProcess[bytes]:
    result = subprocess.run(
        list(command),
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        timeout=timeout,
        check=False,
    )
    if check and result.returncode != 0:
        stderr = result.stderr.decode("utf-8", errors="ignore").strip()
        raise RuntimeError(stderr or f"{command[0]} failed")
    return result


def find_command(command: str) -> str | None:
    search_path = os.environ.get("PATH", "")
    if BUNDLED_BIN_DIR.exists():
        search_path = f"{BUNDLED_BIN_DIR}{os.pathsep}{search_path}"
    return shutil.which(command, path=search_path)


def normalize_ocr_text(text: str) -> str:
    text = text.replace("\r", "\n")
    text = re.sub(r"[ \t\f\v]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_invoice_date(text: str) -> str:
    if not text:
        return ""
    candidates = []
    for match in re.finditer(r"(?<!\d)(\d{1,2})[./-](\d{1,2})[./-]((?:20|21)\d{2})(?!\d)", text):
        parsed = make_date(int(match.group(3)), int(match.group(2)), int(match.group(1)))
        if parsed:
            candidates.append((score_date_context(text, match.start(), match.end()), parsed))
    for match in re.finditer(r"(?<!\d)((?:20|21)\d{2})[./-](\d{1,2})[./-](\d{1,2})(?!\d)", text):
        parsed = make_date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
        if parsed:
            candidates.append((score_date_context(text, match.start(), match.end()), parsed))
    if not candidates:
        return ""
    return max(candidates, key=lambda item: item[0])[1]


def make_date(year: int, month: int, day: int) -> str:
    if not (2000 <= year <= 2100):
        return ""
    try:
        return datetime(year, month, day).strftime("%Y-%m-%d")
    except ValueError:
        return ""


def score_date_context(text: str, start: int, end: int) -> int:
    normalized = normalize_for_match(text[max(0, start - 60) : min(len(text), end + 60)])
    score = 0
    if re.search(r"\b(?:data|date)\b", normalized):
        score += 20
    if re.search(r"\b(?:wystawienia|sprzedazy|faktur\w*|invoice)\b", normalized):
        score += 25
    if re.search(r"\b(?:termin|platnosci|due)\b", normalized):
        score += 8
    return score


def extract_invoice_number(text: str) -> str:
    if not text:
        return ""
    patterns = [
        r"\b(?:faktura|invoice)\s*(?:vat)?\s*(?:nr|no\.?|number|#)?\s*[:#]?\s*([A-Z0-9][A-Z0-9/_\-.]{2,})",
        r"\b(?:nr|numer)\s+faktur\w*\s*[:#]?\s*([A-Z0-9][A-Z0-9/_\-.]{2,})",
        r"\binvoice\s*(?:number|no\.?|#)\s*[:#]?\s*([A-Z0-9][A-Z0-9/_\-.]{2,})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            value = match.group(1).strip(" .,:;")
            if not value.casefold().startswith(("data", "date")):
                return value
    return ""


def extract_utility_type(text: str) -> str:
    normalized = normalize_for_match(text)
    utility_services = {"gas", "water", "electricity", "heating"}
    services = [
        hit.service for hit in find_service_hits(normalized) if hit.service in utility_services
    ]
    return display_service_label(unique_ordered(services)) if services else ""


def extract_best_payment_amount_from_text(
    text: str,
    default_currency: str,
    utility_type: str,
) -> tuple[float | None, str, int, str]:
    normalized = normalize_for_match(text)
    amount_hits, rejected_amounts = find_amount_hits(normalized, default_currency=default_currency)
    if not amount_hits:
        detail = summarize_rejections(rejected_amounts)
        reason = "No accepted OCR amount candidates."
        if detail:
            reason = f"{reason} {detail}"
        return None, "", 0, reason

    service_hits = find_service_hits(normalized)
    if utility_type:
        allowed = {service.casefold() for service in utility_type.split(" + ")}
        reverse_display = {value.casefold(): key for key, value in SERVICE_DISPLAY.items()}
        allowed_services = {reverse_display.get(value, value) for value in allowed}
        service_hits = [hit for hit in service_hits if hit.service in allowed_services]

    all_services = unique_ordered([hit.service for hit in service_hits]) or ["other services"]

    scored = []
    for amount_hit in amount_hits:
        selected_services = choose_services_for_amount(
            amount_hit=amount_hit,
            service_hits=service_hits,
            all_services=all_services,
            amount_count=len(amount_hits),
            service_window=160,
        ) or all_services
        scored.append(
            score_invoice_amount_candidate(
                text=text,
                normalized_text=normalized,
                amount_hit=amount_hit,
                selected_services=selected_services,
                default_currency=default_currency,
            )
        )

    scored = [candidate for candidate in scored if candidate.confidence_score >= MIN_PAYMENT_CONFIDENCE]
    if not scored:
        detail = summarize_rejections(rejected_amounts)
        reason = "OCR amount candidates were present, but none looked like the final payable amount."
        if detail:
            reason = f"{reason} Rejected candidates: {detail}"
        return None, "", 0, reason

    best = max(
        scored,
        key=lambda candidate: (
            candidate.confidence_score,
            invoice_payable_context_rank(candidate.hit.payment_context_labels),
            candidate.hit.currency is not None,
            candidate.hit.start,
        ),
    )
    currency = best.hit.currency or default_currency.upper()
    reason = best.reason
    ignored_detail = summarize_rejections(rejected_amounts, limit=6)
    if ignored_detail:
        reason = f"{reason} Ignored other numbers: {ignored_detail}."
    return round(best.hit.amount, 2), currency, best.confidence_score, reason


def score_invoice_amount_candidate(
    text: str,
    normalized_text: str,
    amount_hit: AmountHit,
    selected_services: Sequence[str],
    default_currency: str,
) -> ScoredAmount:
    score = 0
    reasons = []
    currency = amount_hit.currency or default_currency.upper()

    if amount_hit.currency:
        score += 35
        reasons.append(f"explicit {currency} currency")
    else:
        score += 4
        reasons.append(f"{default_currency.upper()} assumed")

    payable_labels = labels_near_range(
        normalized_text=normalized_text,
        start=amount_hit.start,
        end=amount_hit.end,
        patterns=PAYABLE_CONTEXT_PATTERNS,
        window=120,
    )
    rank = invoice_payable_context_rank(payable_labels)
    if payable_labels:
        score += 35 + rank
        reasons.append("near final-payable keyword " + ", ".join(f"'{label}'" for label in payable_labels))

    line = line_containing_offset(text, amount_hit.start)
    line_norm = normalize_for_match(line)
    if any(pattern.search(line_norm) for _, pattern in PAYABLE_CONTEXT_PATTERNS):
        score += 20
        reasons.append("payable keyword appears on the same OCR line")

    specific_services = [service for service in selected_services if service != "other services"]
    if specific_services:
        score += 12
        reasons.append("near utility keyword " + ", ".join(f"'{SERVICE_DISPLAY.get(service, service)}'" for service in specific_services))

    if amount_hit.amount <= MAX_UTILITY_AMOUNT:
        score += 12
        reasons.append(f"amount is within {MAX_UTILITY_AMOUNT} PLN utility limit")

    confidence_score = max(0, min(100, score))
    reason = "Selected as final payable OCR amount: " + "; ".join(reasons) + "."
    return ScoredAmount(
        hit=amount_hit,
        services=tuple(selected_services),
        confidence_score=confidence_score,
        reason=reason,
    )


def invoice_payable_context_rank(labels: Sequence[str]) -> int:
    priority = {
        "kwota do zaplaty": 60,
        "do zaplaty": 55,
        "do uregulowania": 52,
        "naleznosc": 45,
        "razem": 35,
        "PLN/zl": 20,
    }
    return max((priority.get(label, 0) for label in labels), default=0)


def line_containing_offset(text: str, normalized_offset: int) -> str:
    # Normalization can shift offsets slightly; a nearby raw slice is enough for line-level context.
    raw_offset = min(max(normalized_offset, 0), len(text))
    start = text.rfind("\n", 0, raw_offset) + 1
    end = text.find("\n", raw_offset)
    if end == -1:
        end = len(text)
    return text[start:end]


def ocr_payments_from_results(
    ocr_results: Sequence[OcrResultRow],
    attachments: Sequence[AttachmentRow],
    default_currency: str,
) -> list[PaymentRow]:
    attachment_by_id = {attachment.attachment_id: attachment for attachment in attachments}
    payments = []
    for result in ocr_results:
        if result.amount is None:
            continue
        attachment = attachment_by_id.get(result.attachment_id)
        message_date = parse_iso_date(result.invoice_date) or result.message_date
        service_type = result.utility_type or "Other Services"
        payments.append(
            PaymentRow(
                message_date=message_date,
                service_type=service_type,
                amount=result.amount,
                currency=result.currency or default_currency.upper(),
                currency_source="OCR",
                confidence_score=result.confidence_score,
                extraction_reason=result.extraction_reason,
                source_type="OCR Attachment",
                direction=attachment.direction if attachment else "Incoming",
                sender=attachment.sender if attachment else "",
                message_id=result.message_id,
                original_text=result.ocr_text[:1000],
                attachment_id=result.attachment_id,
                attachment_path=result.copied_path,
                invoice_date=result.invoice_date,
                invoice_number=result.invoice_number,
            )
        )
    return payments


def parse_iso_date(value: str) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d").astimezone()
    except ValueError:
        return None


def build_workbook_sheets(
    messages: Sequence[MessageRow],
    attachments: Sequence[AttachmentRow],
    ocr_results: Sequence[OcrResultRow],
    payments: Sequence[PaymentRow],
) -> list[tuple[str, list[list[object]]]]:
    message_rows = [
        [
            "Message Date",
            "Direction",
            "Sender/Handle",
            "Handle Service",
            "Chat Name",
            "Chat Identifier",
            "Message ID",
            "Original Message Text",
        ]
    ]
    for message in messages:
        message_rows.append(
            [
                format_datetime(message.date),
                message.direction,
                message.handle_id,
                message.handle_service,
                message.chat_display_names,
                message.chat_identifiers,
                message.message_id,
                message.text,
            ]
        )

    attachment_rows = [
        [
            "Message Date",
            "Message ID",
            "Attachment ID",
            "Direction",
            "Sender",
            "Transfer Name",
            "MIME Type",
            "UTI",
            "Bytes",
            "Original Attachment Path",
            "Resolved Attachment Path",
            "Copied Attachment Path",
            "Invoice Related",
            "Copy Status",
            "Message Text",
        ]
    ]
    for attachment in attachments:
        attachment_rows.append(
            [
                format_datetime(attachment.message_date),
                attachment.message_id,
                attachment.attachment_id,
                attachment.direction,
                attachment.sender,
                attachment.transfer_name,
                attachment.mime_type,
                attachment.uti,
                attachment.total_bytes if attachment.total_bytes is not None else "",
                attachment.original_filename,
                attachment.resolved_path,
                attachment.copied_path,
                "Yes" if attachment.is_invoice_related else "No",
                attachment.copy_status,
                attachment.message_text,
            ]
        )

    ocr_rows = [
        [
            "Message Date",
            "Message ID",
            "Attachment ID",
            "File Type",
            "OCR Status",
            "Invoice Date",
            "Invoice Number",
            "Utility Type",
            "Amount",
            "Currency",
            "Confidence Score",
            "Extraction Reason",
            "Copied Attachment Path",
            "Original Attachment Path",
            "OCR Text",
        ]
    ]
    for result in ocr_results:
        ocr_rows.append(
            [
                format_datetime(result.message_date),
                result.message_id,
                result.attachment_id,
                result.file_type,
                result.ocr_status,
                result.invoice_date,
                result.invoice_number,
                result.utility_type,
                result.amount if result.amount is not None else "",
                result.currency,
                result.confidence_score,
                result.extraction_reason,
                result.copied_path,
                result.source_path,
                result.ocr_text,
            ]
        )

    payment_rows = [
        [
            "Date",
            "Source",
            "Service Type",
            "Amount",
            "Currency",
            "Currency Source",
            "Confidence Score",
            "Invoice Date",
            "Invoice Number",
            "Direction",
            "Sender/Handle",
            "Message ID",
            "Attachment ID",
            "Attachment Path",
            "Extraction Reason",
            "Source Text",
        ]
    ]
    for payment in payments:
        payment_rows.append(
            [
                format_datetime(payment.message_date),
                payment.source_type,
                payment.service_type,
                payment.amount,
                payment.currency,
                payment.currency_source,
                payment.confidence_score,
                payment.invoice_date,
                payment.invoice_number,
                payment.direction,
                payment.sender,
                payment.message_id,
                payment.attachment_id if payment.attachment_id is not None else "",
                payment.attachment_path,
                payment.extraction_reason,
                payment.original_text,
            ]
        )

    return [
        ("Messages", message_rows),
        ("Attachments", attachment_rows),
        ("OCR Results", ocr_rows),
        ("Payments", payment_rows),
        ("Validation", ocr_validation_rows(ocr_results)),
        ("Monthly Summary", monthly_summary_rows(payments)),
        ("Yearly Summary", yearly_summary_rows(payments)),
    ]


def ocr_validation_rows(ocr_results: Sequence[OcrResultRow]) -> list[list[object]]:
    rows = [
        [
            "Message Date",
            "Message ID",
            "Attachment ID",
            "OCR Text Snippet",
            "Extracted Amount",
            "Currency",
            "Confidence",
            "Extraction Reason",
            "Copied Attachment Path",
        ]
    ]
    for result in ocr_results:
        rows.append(
            [
                format_datetime(result.message_date),
                result.message_id,
                result.attachment_id,
                ocr_text_snippet(result.ocr_text, result.amount),
                result.amount if result.amount is not None else "",
                result.currency,
                result.confidence_score,
                result.extraction_reason,
                result.copied_path,
            ]
        )
    return rows


def ocr_text_snippet(text: str, amount: float | None, window: int = 180) -> str:
    if not text:
        return ""
    if amount is None:
        return normalize_spaces(text[: window * 2])

    normalized = normalize_for_match(text)
    for match in AMOUNT_RE.finditer(normalized):
        parsed = parse_amount(match.group("amount"))
        if parsed is not None and abs(parsed - amount) < 0.01:
            snippet = text_context(text, match.start(), match.end(), window=window)
            return normalize_spaces(snippet)
    return normalize_spaces(text[: window * 2])


def yearly_summary_rows(payments: Sequence[PaymentRow]) -> list[list[object]]:
    summary = defaultdict(lambda: {"count": 0, "total": 0.0})
    for payment in payments:
        year = payment.message_date.strftime("%Y") if payment.message_date else "Unknown"
        key = (year, payment.service_type, payment.currency)
        summary[key]["count"] += 1
        summary[key]["total"] += payment.amount

    rows = [["Year", "Service Type", "Currency", "Payment Count", "Total Amount"]]
    for (year, service, currency), values in sorted(summary.items()):
        rows.append([year, service, currency, values["count"], round(values["total"], 2)])
    return rows


def monthly_summary_rows(payments: Sequence[PaymentRow]) -> list[list[object]]:
    summary = defaultdict(lambda: {"count": 0, "total": 0.0})
    for payment in payments:
        month = payment.message_date.strftime("%Y-%m") if payment.message_date else "Unknown"
        key = (month, payment.service_type, payment.currency)
        summary[key]["count"] += 1
        summary[key]["total"] += payment.amount

    rows = [["Month", "Service Type", "Currency", "Payment Count", "Total Amount"]]
    for (month, service, currency), values in sorted(summary.items()):
        rows.append([month, service, currency, values["count"], round(values["total"], 2)])
    return rows


def by_service_rows(payments: Sequence[PaymentRow]) -> list[list[object]]:
    summary = defaultdict(lambda: {"count": 0, "total": 0.0, "first": None, "last": None})
    for payment in payments:
        key = (payment.service_type, payment.currency)
        item = summary[key]
        item["count"] += 1
        item["total"] += payment.amount
        if payment.message_date and (item["first"] is None or payment.message_date < item["first"]):
            item["first"] = payment.message_date
        if payment.message_date and (item["last"] is None or payment.message_date > item["last"]):
            item["last"] = payment.message_date

    rows = [["Service Type", "Currency", "Payment Count", "Total Amount", "First Date", "Last Date"]]
    for (service, currency), values in sorted(summary.items()):
        rows.append(
            [
                service,
                currency,
                values["count"],
                round(values["total"], 2),
                format_datetime(values["first"]),
                format_datetime(values["last"]),
            ]
        )
    return rows


def format_datetime(value: datetime | None) -> str:
    if value is None:
        return ""
    return value.strftime("%Y-%m-%d %H:%M:%S")


def write_xlsx(path: Path, sheets: Sequence[tuple[str, Sequence[Sequence[object]]]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_name(f"{path.name}.tmp")
    if temp_path.exists():
        temp_path.unlink()

    with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types_xml(len(sheets)))
        archive.writestr("_rels/.rels", package_rels_xml())
        archive.writestr("xl/workbook.xml", workbook_xml([name for name, _ in sheets]))
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(sheets)))
        archive.writestr("xl/styles.xml", styles_xml())
        for index, (_, rows) in enumerate(sheets, start=1):
            archive.writestr(f"xl/worksheets/sheet{index}.xml", worksheet_xml(rows))

    os.replace(temp_path, path)


def content_types_xml(sheet_count: int) -> str:
    sheet_overrides = "\n".join(
        f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for index in range(1, sheet_count + 1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
{sheet_overrides}
</Types>"""


def package_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""


def workbook_xml(sheet_names: Sequence[str]) -> str:
    sheet_xml = "\n".join(
        f'<sheet name="{xml_attr(safe_sheet_name(name))}" sheetId="{index}" r:id="rId{index}"/>'
        for index, name in enumerate(sheet_names, start=1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
{sheet_xml}
</sheets>
</workbook>"""


def workbook_rels_xml(sheet_count: int) -> str:
    rels = []
    for index in range(1, sheet_count + 1):
        rels.append(
            f'<Relationship Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
    rels.append(
        f'<Relationship Id="rId{sheet_count + 1}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{chr(10).join(rels)}
</Relationships>"""


def styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def worksheet_xml(rows: Sequence[Sequence[object]]) -> str:
    max_columns = max((len(row) for row in rows), default=1)
    max_rows = max(len(rows), 1)
    dimension = f"A1:{column_name(max_columns)}{max_rows}"
    row_xml = "\n".join(
        worksheet_row_xml(row_index, row) for row_index, row in enumerate(rows, start=1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="{dimension}"/>
<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<sheetData>
{row_xml}
</sheetData>
</worksheet>"""


def worksheet_row_xml(row_index: int, row: Sequence[object]) -> str:
    cells = []
    for column_index, value in enumerate(row, start=1):
        if value is None or value == "":
            continue
        cell_ref = f"{column_name(column_index)}{row_index}"
        cells.append(cell_xml(cell_ref, value))
    return f'<row r="{row_index}">{"".join(cells)}</row>'


def cell_xml(cell_ref: str, value: object) -> str:
    if isinstance(value, bool):
        return f'<c r="{cell_ref}" t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(float(value)):
        return f'<c r="{cell_ref}"><v>{value}</v></c>'

    text = xml_text(str(value))
    preserve = ' xml:space="preserve"' if needs_space_preserved(text) else ""
    return f'<c r="{cell_ref}" t="inlineStr"><is><t{preserve}>{text}</t></is></c>'


def column_name(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def safe_sheet_name(name: str) -> str:
    value = re.sub(r"[\[\]:*?/\\]", " ", name).strip() or "Sheet"
    return value[:31]


def xml_text(value: str) -> str:
    cleaned = "".join(char if is_valid_xml_char(char) else " " for char in value)
    return escape(cleaned)


def xml_attr(value: str) -> str:
    return escape(value, {'"': "&quot;"})


def is_valid_xml_char(char: str) -> bool:
    codepoint = ord(char)
    return (
        codepoint in (0x9, 0xA, 0xD)
        or 0x20 <= codepoint <= 0xD7FF
        or 0xE000 <= codepoint <= 0xFFFD
        or 0x10000 <= codepoint <= 0x10FFFF
    )


def needs_space_preserved(value: str) -> bool:
    return bool(value) and (value[0].isspace() or value[-1].isspace() or "\n" in value)


if __name__ == "__main__":
    raise SystemExit(main())
