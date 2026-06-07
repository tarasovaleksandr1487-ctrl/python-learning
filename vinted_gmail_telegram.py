#!/usr/bin/env python3
"""
Poll Gmail for unread Vinted sold-item emails and notify a Telegram bot.

Required environment variables:
  TELEGRAM_BOT_TOKEN
  TELEGRAM_CHAT_ID
  GMAIL_EMAIL
  GMAIL_APP_PASSWORD

Optional environment variables:
  GMAIL_IMAP_HOST        default: imap.gmail.com
  GMAIL_IMAP_PORT        default: 993
  GMAIL_MAILBOX          default: INBOX
  VINTED_FROM_QUERY      default: vinted
  VINTED_SOLD_KEYWORDS   comma-separated; default includes English and Polish terms
  STATE_FILE             default: .vinted_gmail_telegram_state.json
  MAX_EMAILS             default: 20
"""

from __future__ import annotations

import email
import html
import imaplib
import json
import os
import re
import sys
from dataclasses import dataclass
from email.header import decode_header
from email.message import Message
from pathlib import Path
from typing import Iterable
from urllib.parse import quote
from urllib.request import Request, urlopen


DEFAULT_KEYWORDS = (
    "sold",
    "has sold",
    "just sold",
    "item sold",
    "sprzedane",
    "sprzedala",
    "sprzedała",
    "sprzedales",
    "sprzedałeś",
    "kupiono",
)


@dataclass(frozen=True)
class Config:
    telegram_bot_token: str
    telegram_chat_id: str
    gmail_email: str
    gmail_app_password: str
    gmail_imap_host: str
    gmail_imap_port: int
    gmail_mailbox: str
    vinted_from_query: str
    sold_keywords: tuple[str, ...]
    state_file: Path
    max_emails: int


@dataclass(frozen=True)
class SaleDetails:
    item_title: str
    sale_price: str
    buyer_name: str


@dataclass(frozen=True)
class CandidateEmail:
    uid: str
    subject: str
    sender: str
    message_id: str
    snippet: str
    gmail_link: str
    sale_details: SaleDetails


def getenv_required(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def load_config() -> Config:
    keywords = tuple(
        keyword.strip().lower()
        for keyword in os.getenv("VINTED_SOLD_KEYWORDS", ",".join(DEFAULT_KEYWORDS)).split(",")
        if keyword.strip()
    )

    return Config(
        telegram_bot_token=getenv_required("TELEGRAM_BOT_TOKEN"),
        telegram_chat_id=getenv_required("TELEGRAM_CHAT_ID"),
        gmail_email=getenv_required("GMAIL_EMAIL"),
        gmail_app_password=getenv_required("GMAIL_APP_PASSWORD"),
        gmail_imap_host=os.getenv("GMAIL_IMAP_HOST", "imap.gmail.com"),
        gmail_imap_port=int(os.getenv("GMAIL_IMAP_PORT", "993")),
        gmail_mailbox=os.getenv("GMAIL_MAILBOX", "INBOX"),
        vinted_from_query=os.getenv("VINTED_FROM_QUERY", "vinted"),
        sold_keywords=keywords,
        state_file=Path(os.getenv("STATE_FILE", ".vinted_gmail_telegram_state.json")),
        max_emails=int(os.getenv("MAX_EMAILS", "20")),
    )


def decode_mime_header(value: str | None) -> str:
    if not value:
        return ""

    parts: list[str] = []
    for decoded, charset in decode_header(value):
        if isinstance(decoded, bytes):
            parts.append(decoded.decode(charset or "utf-8", errors="replace"))
        else:
            parts.append(decoded)
    return "".join(parts).strip()


def text_from_message(message: Message) -> str:
    if message.is_multipart():
        plain_parts: list[str] = []
        html_parts: list[str] = []

        for part in message.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if part.get_filename():
                continue

            content_type = part.get_content_type()
            payload = decode_part_payload(part)
            if not payload:
                continue

            if content_type == "text/plain":
                plain_parts.append(payload)
            elif content_type == "text/html":
                html_parts.append(strip_html(payload))

        return "\n".join(plain_parts or html_parts)

    payload = decode_part_payload(message)
    if message.get_content_type() == "text/html":
        return strip_html(payload)
    return payload


def decode_part_payload(part: Message) -> str:
    payload = part.get_payload(decode=True)
    if payload is None:
        raw_payload = part.get_payload()
        return raw_payload if isinstance(raw_payload, str) else ""

    charset = part.get_content_charset() or "utf-8"
    return payload.decode(charset, errors="replace")


def strip_html(value: str) -> str:
    value = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", value)
    value = re.sub(r"(?s)<br\s*/?>", "\n", value)
    value = re.sub(r"(?s)</p\s*>", "\n", value)
    value = re.sub(r"(?s)<.*?>", " ", value)
    return html.unescape(value)


def compact_snippet(value: str, limit: int = 500) -> str:
    snippet = re.sub(r"\s+", " ", value).strip()
    if len(snippet) <= limit:
        return snippet
    return snippet[: limit - 1].rstrip() + "..."


def has_sold_signal(subject: str, body: str, keywords: Iterable[str]) -> bool:
    searchable = f"{subject}\n{body}".lower()
    return any(keyword in searchable for keyword in keywords)


def first_regex_match(patterns: Iterable[str], value: str) -> str:
    for pattern in patterns:
        match = re.search(pattern, value, flags=re.IGNORECASE | re.MULTILINE)
        if match:
            return compact_snippet(match.group(1), limit=120)
    return ""


def extract_sale_price(subject: str, body: str) -> str:
    searchable = f"{subject}\n{body}"
    price_patterns = (
        r"(?:price|sale price|sold for|amount|total|paid|buyer paid|cena|kwota|sprzedano za|kupiono za)\s*[:\-]?\s*((?:€|£|\$|zł|PLN|EUR|USD|GBP)\s*\d+(?:[.,]\d{1,2})?|\d+(?:[.,]\d{1,2})?\s*(?:€|£|\$|zł|PLN|EUR|USD|GBP))",
        r"((?:€|£|\$|zł|PLN|EUR|USD|GBP)\s*\d+(?:[.,]\d{1,2})?|\d+(?:[.,]\d{1,2})?\s*(?:€|£|\$|zł|PLN|EUR|USD|GBP))",
    )
    return first_regex_match(price_patterns, searchable)


def extract_buyer_name(body: str) -> str:
    buyer_patterns = (
        r"(?:buyer|sold to|purchased by|kupujący|kupujacy|sprzedane użytkownikowi|sprzedane uzytkownikowi)\s*[:\-]?\s*([^\n\r]+)",
        r"(?:ship|send|wyślij|wyslij).{0,40}(?:to|do)\s+([A-ZŁŚŻŹĆŃÓĘĄ][^\n\r,.;]{1,60})",
    )
    buyer = first_regex_match(buyer_patterns, body)
    return remove_common_noise(buyer)


def extract_item_title(subject: str, body: str) -> str:
    searchable = f"{subject}\n{body}"
    title_patterns = (
        r"^\s*(?:item|title|listing|product|przedmiot|tytuł|tytul|ogłoszenie|ogloszenie)\s*[:\-]\s*[\"“”']?([^\n\r\"“”']{2,120})",
        r"(?:your item sold|you sold|has sold|sprzedałaś|sprzedalas|sprzedałeś|sprzedales|sprzedano|kupiono)\s*[:\-]\s*[\"“”']?([^\n\r\"“”']{2,120})",
        r"(?:you sold|has sold|sold|sprzedałaś|sprzedalas|sprzedałeś|sprzedales|sprzedano|kupiono)\s+[\"“”']?([^\n\r\"“”']{2,120})",
        r"[\"“”']([^\"“”'\n\r]{2,120})[\"“”']",
    )
    title = first_regex_match(title_patterns, searchable)
    return remove_common_noise(title) or subject


def remove_common_noise(value: str) -> str:
    if not value:
        return ""

    noise_patterns = (
        r"\b(on vinted|via vinted|from vinted|w vinted)\b.*$",
        r"\b(for|za)\s+(?:€|£|\$|zł|PLN|EUR|USD|GBP)?\s*\d+.*$",
        r"\b(?:€|£|\$|zł|PLN|EUR|USD|GBP)\s*\d+.*$",
    )
    cleaned = value
    for pattern in noise_patterns:
        cleaned = re.sub(pattern, "", cleaned, flags=re.IGNORECASE).strip(" \t:-,.")
    return cleaned


def extract_sale_details(subject: str, body: str) -> SaleDetails:
    return SaleDetails(
        item_title=extract_item_title(subject, body),
        sale_price=extract_sale_price(subject, body),
        buyer_name=extract_buyer_name(body),
    )


def load_seen(state_file: Path) -> set[str]:
    if not state_file.exists():
        return set()

    with state_file.open("r", encoding="utf-8") as file:
        data = json.load(file)
    return set(data.get("notified_message_ids", []))


def save_seen(state_file: Path, seen: set[str]) -> None:
    state_file.parent.mkdir(parents=True, exist_ok=True)
    with state_file.open("w", encoding="utf-8") as file:
        json.dump({"notified_message_ids": sorted(seen)}, file, indent=2)


def gmail_search_link(message_id: str) -> str:
    if message_id:
        query = f"rfc822msgid:{message_id}"
    else:
        query = "from:vinted unread"
    return f"https://mail.google.com/mail/u/0/#search/{quote(query, safe='')}"


def fetch_unread_vinted_emails(config: Config) -> list[CandidateEmail]:
    with imaplib.IMAP4_SSL(config.gmail_imap_host, config.gmail_imap_port) as mailbox:
        mailbox.login(config.gmail_email, config.gmail_app_password)
        mailbox.select(config.gmail_mailbox, readonly=True)

        search_query = f'(UNSEEN FROM "{config.vinted_from_query}")'
        status, data = mailbox.uid("search", None, search_query)
        if status != "OK":
            raise RuntimeError(f"Gmail search failed: {status}")

        uids = data[0].split()[-config.max_emails :]
        candidates: list[CandidateEmail] = []

        for raw_uid in uids:
            uid = raw_uid.decode("ascii", errors="replace")
            status, fetch_data = mailbox.uid("fetch", raw_uid, "(RFC822)")
            if status != "OK" or not fetch_data:
                continue

            raw_message = next(
                (item[1] for item in fetch_data if isinstance(item, tuple) and item[1]),
                None,
            )
            if not raw_message:
                continue

            message = email.message_from_bytes(raw_message)
            subject = decode_mime_header(message.get("Subject"))
            sender = decode_mime_header(message.get("From"))
            message_id = (message.get("Message-ID") or "").strip()
            body = text_from_message(message)

            if not has_sold_signal(subject, body, config.sold_keywords):
                continue

            candidates.append(
                CandidateEmail(
                    uid=uid,
                    subject=subject or "(no subject)",
                    sender=sender,
                    message_id=message_id or f"imap-uid:{uid}",
                    snippet=compact_snippet(body),
                    gmail_link=gmail_search_link(message_id),
                    sale_details=extract_sale_details(subject, body),
                )
            )

        return candidates


def send_telegram_message(config: Config, candidate: CandidateEmail) -> None:
    details = candidate.sale_details
    price = details.sale_price or "Not found"
    buyer = details.buyer_name or "Not available"

    text = (
        "🎉 Vinted sale!\n\n"
        f"📦 Item: {details.item_title}\n"
        f"💰 Price: {price}\n"
        f"👤 Buyer: {buyer}\n"
        f"✉️ Subject: {candidate.subject}\n"
        f"🔗 Email: {candidate.gmail_link}\n\n"
        f"📝 {candidate.snippet}"
    )

    payload = json.dumps(
        {
            "chat_id": config.telegram_chat_id,
            "text": text,
            "disable_web_page_preview": True,
        }
    ).encode("utf-8")

    request = Request(
        f"https://api.telegram.org/bot{config.telegram_bot_token}/sendMessage",
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    with urlopen(request, timeout=20) as response:
        if response.status >= 400:
            raise RuntimeError(f"Telegram request failed with HTTP {response.status}")


def main() -> int:
    config = load_config()
    seen = load_seen(config.state_file)
    candidates = fetch_unread_vinted_emails(config)
    new_candidates = [candidate for candidate in candidates if candidate.message_id not in seen]

    for candidate in new_candidates:
        send_telegram_message(config, candidate)
        seen.add(candidate.message_id)

    save_seen(config.state_file, seen)
    print(f"Checked Gmail: sent {len(new_candidates)} Telegram notification(s).")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1)
