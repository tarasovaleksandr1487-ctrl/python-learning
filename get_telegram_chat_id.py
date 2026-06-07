#!/usr/bin/env python3
"""Print chat IDs from recent Telegram bot updates."""

from __future__ import annotations

import json
import os
import sys
from urllib.request import urlopen


def main() -> int:
    bot_token = os.getenv("TELEGRAM_BOT_TOKEN")
    if not bot_token:
        print("Error: set TELEGRAM_BOT_TOKEN first.", file=sys.stderr)
        return 1

    url = f"https://api.telegram.org/bot{bot_token}/getUpdates"
    with urlopen(url, timeout=20) as response:
        data = json.load(response)

    if not data.get("ok"):
        print(f"Telegram API error: {data}", file=sys.stderr)
        return 1

    updates = data.get("result", [])
    if not updates:
        print("No updates found. Send a message to your bot, then run this script again.")
        return 0

    seen_chat_ids: set[int] = set()
    for update in updates:
        message = (
            update.get("message")
            or update.get("edited_message")
            or update.get("channel_post")
            or update.get("edited_channel_post")
        )
        if not message or "chat" not in message:
            continue

        chat = message["chat"]
        chat_id = chat["id"]
        if chat_id in seen_chat_ids:
            continue

        seen_chat_ids.add(chat_id)
        title = chat.get("title") or chat.get("username") or chat.get("first_name") or "unknown"
        print(f"chat_id={chat_id} type={chat.get('type', 'unknown')} name={title}")

    if not seen_chat_ids:
        print("Updates were found, but none contained a chat ID.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
