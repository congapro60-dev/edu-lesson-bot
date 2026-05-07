from __future__ import annotations

import argparse
from datetime import datetime

import requests

from app.config import load_settings, require_values


class TelegramNotifier:
    def __init__(self, bot_token: str, chat_id: str) -> None:
        self.bot_token = bot_token
        self.chat_id = chat_id
        self.api_base = f"https://api.telegram.org/bot{bot_token}"

    def send_message(self, text: str) -> dict:
        response = requests.post(
            f"{self.api_base}/sendMessage",
            json={
                "chat_id": self.chat_id,
                "text": text,
                "disable_web_page_preview": True,
            },
            timeout=30,
        )
        if not response.ok:
            raise RuntimeError(
                "Telegram API request failed: "
                f"HTTP {response.status_code} - {response.text}"
            )
        return response.json()


def build_notifier() -> TelegramNotifier:
    settings = load_settings()
    require_values(settings, ["telegram_bot_token", "telegram_chat_id"])
    return TelegramNotifier(
        bot_token=settings.telegram_bot_token,
        chat_id=settings.telegram_chat_id,
    )


def send_test_message() -> None:
    notifier = build_notifier()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    notifier.send_message(
        "✅ Test Telegram thành công.\n"
        f"Lesson Bot đã kết nối được với Telegram lúc {now}."
    )
    print("Telegram test message sent successfully.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Telegram notification helper")
    parser.add_argument("--test", action="store_true", help="Send a test message")
    args = parser.parse_args()

    if args.test:
        send_test_message()
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
