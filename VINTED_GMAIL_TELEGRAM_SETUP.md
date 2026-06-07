# Vinted Gmail to Telegram Automation

This automation checks unread Gmail messages from Vinted, looks for sold-item wording, and sends a Telegram bot message with the subject, Gmail search link, and email snippet.

It does not scrape Vinted.

## 1. Gmail setup

Use a Gmail app password, not your normal Google password.

1. Enable IMAP in Gmail settings.
2. Enable 2-Step Verification on your Google account.
3. Create an app password for Mail.

## 2. Environment variables

```sh
export TELEGRAM_BOT_TOKEN="123456:your_bot_token"
export TELEGRAM_CHAT_ID="123456789"
export GMAIL_EMAIL="your.email@gmail.com"
export GMAIL_APP_PASSWORD="your_gmail_app_password"
```

Optional:

```sh
export VINTED_FROM_QUERY="vinted"
export VINTED_SOLD_KEYWORDS="sold,has sold,just sold,item sold,sprzedane,sprzedała,sprzedałeś,kupiono"
export STATE_FILE="/Users/alex/Desktop/python/.vinted_gmail_telegram_state.json"
export MAX_EMAILS="20"
```

## 3. Run it once

```sh
python3 /Users/alex/Desktop/python/vinted_gmail_telegram.py
```

## 4. Run it every 5 minutes on macOS cron

Cron does not automatically receive the environment variables from your Terminal session. After exporting the variables in step 2, run this to save your current values into a private env file:

```sh
umask 077
{
  printf 'export TELEGRAM_BOT_TOKEN=%q\n' "$TELEGRAM_BOT_TOKEN"
  printf 'export TELEGRAM_CHAT_ID=%q\n' "$TELEGRAM_CHAT_ID"
  printf 'export GMAIL_EMAIL=%q\n' "$GMAIL_EMAIL"
  printf 'export GMAIL_APP_PASSWORD=%q\n' "$GMAIL_APP_PASSWORD"
  printf 'export VINTED_FROM_QUERY=%q\n' "${VINTED_FROM_QUERY:-vinted}"
  printf 'export VINTED_SOLD_KEYWORDS=%q\n' "${VINTED_SOLD_KEYWORDS:-sold,has sold,just sold,item sold,sprzedane,sprzedała,sprzedałeś,kupiono}"
  printf 'export STATE_FILE=%q\n' "${STATE_FILE:-/Users/alex/Desktop/python/.vinted_gmail_telegram_state.json}"
  printf 'export MAX_EMAILS=%q\n' "${MAX_EMAILS:-20}"
} > /Users/alex/Desktop/python/vinted_gmail_telegram.env
```

Install the cron job:

```sh
CRON_LINE="*/5 * * * * /bin/zsh -lc 'source /Users/alex/Desktop/python/vinted_gmail_telegram.env && cd /Users/alex/Desktop/python && /usr/bin/python3 /Users/alex/Desktop/python/vinted_gmail_telegram.py >> /Users/alex/Desktop/python/vinted_gmail_telegram.log 2>&1'"
(crontab -l 2>/dev/null | grep -v 'vinted_gmail_telegram.py'; echo "$CRON_LINE") | crontab -
```

Check that it was installed:

```sh
crontab -l
```

Watch the log after a few minutes:

```sh
tail -f /Users/alex/Desktop/python/vinted_gmail_telegram.log
```

Remove the cron job:

```sh
crontab -l 2>/dev/null | grep -v 'vinted_gmail_telegram.py' | crontab -
```

The script does not mark messages as read. It remembers which matching unread message IDs it already notified in `STATE_FILE`.
