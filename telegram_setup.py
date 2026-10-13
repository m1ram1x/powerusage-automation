"""Show chat IDs after /start; never print a bot token or send a message."""

import argparse
import sys
from pathlib import Path

from telegram_alerts import TelegramError, api_call, credentials


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--telegram-env", type=Path)
    args = parser.parse_args(argv)
    try:
        token, _ = credentials(args.telegram_env, require_chat=False)
        updates = api_call(token, "getUpdates", {"timeout": 0, "limit": 100})
        if not isinstance(updates, list):
            raise TelegramError("Telegram returned an unexpected chat list")
        chats = {}
        for update in updates:
            for key in ("message", "edited_message", "channel_post", "my_chat_member"):
                chat = update.get(key, {}).get("chat", {})
                if "id" in chat:
                    chats[chat["id"]] = chat.get("type", "unknown")
        for chat_id, kind in sorted(chats.items()):
            print(f"Chat ID: {chat_id} ({kind})")
        if not chats:
            print("No recent chats: open YOUR bot, press Start and run this command again.")
        return 0
    except TelegramError as exc:
        print(f"TELEGRAM: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
