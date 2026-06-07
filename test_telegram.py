#!/usr/bin/env python3
"""Send a test Telegram message for the Vinted automation."""

from __future__ import annotations

import json
import os
import sys
from urllib.request import Request, urlopen


def main() -> int:
    bot_token = os.getenv("TELEGRAM_BOT_TOKEN")
    chat_id = os.getenv("TELEGRAM_CHAT_ID")

    missing = [
        name
        for name, value in (
            ("TELEGRAM_BOT_TOKEN", bot_token),
            ("TELEGRAM_CHAT_ID", chat_id),
        )
        if not value
    ]
    if missing:
        print(f"Missing environment variable(s): {', '.join(missing)}", file=sys.stderr)
        return 1

    payload = json.dumps(
        {
            "chat_id": chat_id,
            "text": "Test from Vinted automation",
            "disable_web_page_preview": True,
        }
    ).encode("utf-8")

    request = Request(
        f"https://api.telegram.org/bot{bot_token}/sendMessage",
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    with urlopen(request, timeout=20) as response:
        data = json.load(response)

    if not data.get("ok"):
        print(f"Telegram API error: {data.get('description', data)}", file=sys.stderr)
        return 1

    message_id = data.get("result", {}).get("message_id", "unknown")
    print(f"Sent Telegram test message. message_id={message_id}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
