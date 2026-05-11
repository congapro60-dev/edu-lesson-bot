from __future__ import annotations

import argparse
import re
import time
import unicodedata
from dataclasses import dataclass, replace

import requests

from app.config import load_settings, require_values
from app.drive_audit import WeekAudit, audit_week, format_week_audit
from app.drive_client import GoogleDriveClient
from app.lesson_generator import GeneratedLessonFiles, generate_moet_docx, generate_tds_docx
from app.telegram_notify import build_notifier


@dataclass(frozen=True)
class BotCommand:
    action: str
    program: str
    grade: int
    week: int
    upload: bool = False


@dataclass(frozen=True)
class ParsedRequest:
    action: str | None = None
    program: str | None = None
    grade: int | None = None
    week: int | None = None

    def merge(self, other: ParsedRequest) -> ParsedRequest:
        return ParsedRequest(
            action=other.action or self.action,
            program=other.program or self.program,
            grade=other.grade or self.grade,
            week=other.week or self.week,
        )

    def missing_fields(self) -> list[str]:
        missing: list[str] = []
        if not self.action:
            missing.append("action")
        if not self.program:
            missing.append("program")
        if not self.grade:
            missing.append("grade")
        if not self.week:
            missing.append("week")
        return missing

    def to_command(self) -> BotCommand | None:
        if self.missing_fields():
            return None
        assert self.action is not None
        assert self.program is not None
        assert self.grade is not None
        assert self.week is not None
        return BotCommand(
            action=self.action,
            program=self.program,
            grade=self.grade,
            week=self.week,
            upload=self.action == "generate",
        )


HELP_TEXT = """Bot giáo án đang hoạt động.

Bạn có thể nhắn tự nhiên, ví dụ:
- Kiểm tra giáo án lớp 12 MOET tuần 1 đã đủ chưa?
- Kiểm tra TDS lớp 10 tuần 22
- Soạn giáo án MOET lớp 11 tuần 2
- Tạo giáo án TDS lớp 12 tuần 3

Nếu câu lệnh chưa đủ thông tin, bot sẽ hỏi lại đúng phần còn thiếu.
Cần xác định đủ 4 ý: kiểm tra hay soạn, hệ TDS/MOET, lớp 10/11/12, tuần số mấy.
"""


ACTION_LABELS = {"audit": "kiểm tra Drive", "generate": "soạn và upload giáo án"}
PROGRAM_LABELS = {"tds": "TDS", "moet": "MOET"}


def normalize_text(text: str) -> str:
    normalized = unicodedata.normalize("NFC", text)
    return re.sub(r"\s+", " ", normalized.lower()).strip()


def extract_grade(text: str) -> int | None:
    patterns = [
        r"(?:lớp|lop|g|grade|khối|khoi)\s*(10|11|12)\b",
        r"\b(10|11|12)\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return int(match.group(1))
    return None


def extract_week(text: str) -> int | None:
    match = re.search(r"(?:tuần|tuan|week)\s*0*(\d{1,2})\b", text)
    if not match:
        return None
    week = int(match.group(1))
    if 1 <= week <= 40:
        return week
    return None


def extract_program(text: str) -> str | None:
    if "moet" in text or "mơ et" in text or "mộet" in text:
        return "moet"
    if "tds" in text:
        return "tds"
    return None


def extract_action(text: str) -> str | None:
    generate_words = ["soạn", "soan", "tạo", "tao", "generate", "làm", "lam"]
    audit_words = ["kiểm tra", "kiem tra", "check", "audit", "xem", "đã có", "da co", "đủ", "du"]
    if any(word in text for word in generate_words):
        return "generate"
    if any(word in text for word in audit_words):
        return "audit"
    return None


def parse_request(message: str) -> ParsedRequest | BotCommand:
    text = normalize_text(message)
    if text in {"/start", "start", "help", "/help", "hướng dẫn", "huong dan"}:
        return BotCommand(action="help", program="", grade=0, week=0)

    return ParsedRequest(
        action=extract_action(text),
        program=extract_program(text),
        grade=extract_grade(text),
        week=extract_week(text),
    )


def parse_command(message: str) -> BotCommand | None:
    parsed = parse_request(message)
    if isinstance(parsed, BotCommand):
        return parsed
    return parsed.to_command()


def split_message(text: str, max_length: int = 3500) -> list[str]:
    chunks: list[str] = []
    current = ""
    for line in text.splitlines():
        if len(current) + len(line) + 1 > max_length:
            if current:
                chunks.append(current)
            current = line
        else:
            current = f"{current}\n{line}" if current else line
    if current:
        chunks.append(current)
    return chunks or [text]


def build_drive_file_link(file_id: str | None) -> str:
    return f"https://drive.google.com/file/d/{file_id}/view" if file_id else ""


def build_drive_folder_link(folder_id: str | None) -> str:
    return f"https://drive.google.com/drive/folders/{folder_id}" if folder_id else ""


def audit_folder_link(audit: WeekAudit) -> str:
    if not audit.folder_id:
        return ""
    try:
        folder = GoogleDriveClient().get_file(audit.folder_id)
        return folder.get("webViewLink") or build_drive_folder_link(audit.folder_id)
    except Exception:
        return build_drive_folder_link(audit.folder_id)


def format_week_audit_with_links(audit: WeekAudit, program: str = "TDS") -> str:
    base_report = format_week_audit(audit, program)
    folder_link = audit_folder_link(audit)
    if not folder_link:
        return base_report
    return base_report.replace(
        f"Thư mục tuần: {audit.folder_name or 'Chưa có'}",
        f"Thư mục tuần: {audit.folder_name or 'Chưa có'}\nLink thư mục: {folder_link}",
        1,
    )


def find_uploaded_file_link(audit: WeekAudit, filename: str) -> str:
    for file in audit.existing_files:
        if file.get("name") == filename:
            return file.get("webViewLink", "") or build_drive_file_link(file.get("id"))
    return ""


def format_generation_result(command: BotCommand, generated_files: GeneratedLessonFiles) -> str:
    audit = audit_week(command.program, command.grade, command.week)
    folder_link = audit_folder_link(audit)
    docx_link = find_uploaded_file_link(audit, generated_files.docx_path.name)
    pdf_link = find_uploaded_file_link(audit, generated_files.pdf_path.name) if generated_files.pdf_path else ""

    lines = [
        f"Đã soạn và upload giáo án {command.program.upper()} G{command.grade} tuần {command.week:02d}.",
        f"File DOCX local: {generated_files.docx_path.name}",
    ]
    if generated_files.pdf_path:
        lines.append(f"File PDF local: {generated_files.pdf_path.name}")
    lines.append(f"Thư mục chứa file: {audit.folder_name or 'Chưa xác định'}")
    if folder_link:
        lines.append(f"Link thư mục: {folder_link}")
    if docx_link:
        lines.append(f"Link DOCX: {docx_link}")
    else:
        lines.append("Link DOCX: chưa tìm thấy trong lần kiểm tra lại Drive; hãy mở thư mục tuần ở trên để kiểm tra file vừa upload.")
    if pdf_link:
        lines.append(f"Link PDF: {pdf_link}")
    elif generated_files.pdf_path:
        lines.append("Link PDF: chưa tìm thấy trong lần kiểm tra lại Drive; hãy mở thư mục tuần ở trên để kiểm tra file PDF vừa upload.")
    lines.extend([
        "",
        "Kiểm tra nhanh sau khi upload:",
        f"- Số tiết/nội dung theo PPCT: {audit.expected_lessons}",
        f"- Số file giáo án hiện có trong thư mục: {len(audit.existing_files)}",
        f"- Ước lượng còn thiếu: {audit.missing_count}",
    ])
    return "\n".join(lines)


def build_clarification_message(parsed: ParsedRequest) -> str:
    lines = ["Mình cần xác nhận thêm để xử lý đúng yêu cầu."]
    known_parts: list[str] = []
    if parsed.action:
        known_parts.append(f"việc cần làm: {ACTION_LABELS[parsed.action]}")
    if parsed.program:
        known_parts.append(f"hệ: {PROGRAM_LABELS[parsed.program]}")
    if parsed.grade:
        known_parts.append(f"lớp: {parsed.grade}")
    if parsed.week:
        known_parts.append(f"tuần: {parsed.week}")
    if known_parts:
        lines.append("Thông tin đã hiểu: " + "; ".join(known_parts) + ".")

    missing = parsed.missing_fields()
    if "action" in missing:
        lines.append("Bạn muốn bot làm việc nào? Gợi ý: 'kiểm tra' hoặc 'soạn'.")
    if "program" in missing:
        lines.append("Bạn muốn dùng hệ nào? Gợi ý: 'TDS' hoặc 'MOET'.")
    if "grade" in missing:
        lines.append("Bạn muốn lớp mấy? Gợi ý: 'lớp 10', 'lớp 11' hoặc 'lớp 12'.")
    if "week" in missing:
        lines.append("Bạn muốn tuần mấy? Gợi ý: 'tuần 1', 'tuần 2', ...")

    lines.append("")
    lines.append("Bạn có thể trả lời ngắn, ví dụ: 'MOET lớp 12 tuần 1' hoặc 'soạn TDS lớp 10 tuần 22'.")
    return "\n".join(lines)


def handle_command(command: BotCommand) -> str:
    if command.action == "help":
        return HELP_TEXT

    if command.action == "audit":
        audit = audit_week(command.program, command.grade, command.week)
        return format_week_audit_with_links(audit, command.program.upper())

    if command.action == "generate":
        if command.program == "tds":
            generated_files = generate_tds_docx(command.grade, command.week, "dgs", upload=True, notify=False)
        elif command.program == "moet":
            generated_files = generate_moet_docx(command.grade, command.week, upload=True, notify=False)
        else:
            raise ValueError(f"Unsupported program: {command.program}")
        return format_generation_result(command, generated_files)

    return HELP_TEXT


class TelegramPollingBot:
    def __init__(self) -> None:
        settings = load_settings()
        require_values(settings, ["telegram_bot_token"])
        self.owner_chat_id = str(settings.telegram_chat_id)
        self.api_base = f"https://api.telegram.org/bot{settings.telegram_bot_token}"
        self.pending_requests: dict[str, ParsedRequest] = {}

    def get_updates(self, offset: int | None) -> list[dict]:
        params: dict[str, int] = {"timeout": 25}
        if offset is not None:
            params["offset"] = offset
        response = requests.get(f"{self.api_base}/getUpdates", params=params, timeout=35)
        response.raise_for_status()
        payload = response.json()
        if not payload.get("ok"):
            raise RuntimeError(f"Telegram getUpdates failed: {payload}")
        return payload.get("result", [])

    def send_message(self, chat_id: str, text: str) -> None:
        for chunk in split_message(text):
            response = requests.post(
                f"{self.api_base}/sendMessage",
                json={"chat_id": chat_id, "text": chunk},
                timeout=20,
            )
            response.raise_for_status()

    def resolve_command(self, chat_id: str, text: str) -> BotCommand | None:
        parsed = parse_request(text)
        if isinstance(parsed, BotCommand):
            self.pending_requests.pop(chat_id, None)
            return parsed

        pending = self.pending_requests.get(chat_id)
        merged = pending.merge(parsed) if pending else parsed
        command = merged.to_command()
        if command:
            self.pending_requests.pop(chat_id, None)
            return command

        self.pending_requests[chat_id] = merged
        self.send_message(chat_id, build_clarification_message(merged))
        return None

    def handle_message(self, message: dict) -> None:
        chat = message.get("chat") or {}
        chat_id = str(chat.get("id", ""))
        if not chat_id:
            return

        text = (message.get("text") or "").strip()
        if not text:
            return

        command = self.resolve_command(chat_id, text)
        if not command:
            return

        try:
            if command.action != "help":
                action_label = ACTION_LABELS.get(command.action, command.action)
                self.send_message(
                    chat_id,
                    f"Đã nhận lệnh {action_label} {command.program.upper()} G{command.grade} tuần {command.week:02d}. Tôi đang xử lý...",
                )
            result = handle_command(command)
        except Exception as exc:
            result = f"Có lỗi khi xử lý lệnh: {exc}"
        self.send_message(chat_id, result)

    def run(self) -> None:
        if self.owner_chat_id:
            self.send_message(self.owner_chat_id, "Bot giáo án đã bật chế độ nhận lệnh trực tiếp trên Telegram. Gõ /help để xem mẫu lệnh.")
        offset: int | None = None
        while True:
            try:
                updates = self.get_updates(offset)
                for update in updates:
                    offset = update["update_id"] + 1
                    message = update.get("message") or update.get("edited_message")
                    if message:
                        self.handle_message(message)
            except Exception as exc:
                print(f"Telegram polling error: {exc}")
                time.sleep(5)


def main() -> None:
    parser = argparse.ArgumentParser(description="Interactive Telegram listener for Lesson Bot")
    parser.add_argument("--run", action="store_true", help="Run Telegram polling listener")
    args = parser.parse_args()

    if args.run:
        TelegramPollingBot().run()
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
