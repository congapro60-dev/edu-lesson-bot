from __future__ import annotations

import argparse

from app.config import load_settings


def print_next_steps() -> None:
    settings = load_settings()
    print("Lesson Bot scaffold is ready.")
    print(f"Default week: {settings.default_week}")
    print(f"Default year: {settings.default_year}")
    print("Next: configure .env and run Telegram test.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Lesson Bot automation entrypoint")
    parser.add_argument("--status", action="store_true", help="Print current scaffold status")
    args = parser.parse_args()

    if args.status:
        print_next_steps()
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
