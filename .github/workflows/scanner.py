from __future__ import annotations

import os
import sys
import textwrap
from pathlib import Path

import requests


def _run_scanner() -> tuple[Path, Path, Path]:
    package_dir = Path(__file__).resolve().parent / "scanner"
    package_parent = str(package_dir.parent)
    if package_parent not in sys.path:
        sys.path.insert(0, package_parent)

    from scanner.main import run

    return run()


def _send_telegram(markdown_path: Path, excel_path: Path) -> None:
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    chat_id = os.getenv("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        print("Telegram secrets are not configured; skipping notification.")
        return

    report_text = markdown_path.read_text(encoding="utf-8")
    message = textwrap.shorten(report_text.replace("\n", "\n"), width=3500, placeholder="\n\n...")
    api_base = f"https://api.telegram.org/bot{token}"

    response = requests.post(
        f"{api_base}/sendMessage",
        data={
            "chat_id": chat_id,
            "text": message,
            "disable_web_page_preview": "true",
        },
        timeout=30,
    )
    response.raise_for_status()

    with excel_path.open("rb") as file_handle:
        response = requests.post(
            f"{api_base}/sendDocument",
            data={"chat_id": chat_id},
            files={"document": (excel_path.name, file_handle)},
            timeout=60,
        )
    response.raise_for_status()
    print("Telegram notification sent.")


def main() -> None:
    excel_path, markdown_path, _history_path = _run_scanner()
    _send_telegram(markdown_path, excel_path)


if __name__ == "__main__":
    main()
