from __future__ import annotations

import argparse
import os
import threading

import uvicorn

from app.bot_api_server import app as fastapi_app
from app.telegram_bot import TelegramPollingBot


def _run_api_server() -> None:
    # Railway injects PORT; fall back to API_PORT or 8000
    port = int(os.getenv("PORT", os.getenv("API_PORT", "8000")))
    uvicorn.run(fastapi_app, host="0.0.0.0", port=port, log_level="info")


def main() -> None:
    parser = argparse.ArgumentParser(description="Lesson Bot entrypoint")
    parser.add_argument("--run", action="store_true", help="Start API server + Telegram bot")
    args = parser.parse_args()

    if args.run:
        # API server runs in a daemon thread so it dies when the bot exits
        api_thread = threading.Thread(target=_run_api_server, daemon=True, name="api-server")
        api_thread.start()
        TelegramPollingBot().run()
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
