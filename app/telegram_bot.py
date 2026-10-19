from __future__ import annotations

import argparse
import hashlib
import re
import time
import unicodedata
from dataclasses import dataclass, replace

import requests

from app.config import load_settings, require_values
from app.assistant_client import AssistantClient, AssistantClientError, AssistantIdentity, redact_assistant_text
from app.drive_audit import WeekAudit, audit_week, format_week_audit
from app.drive_client import GoogleDriveClient
from app.lesson_generator import (
    ExistingOutputConflict,
    GeneratedLessonBatch,
    find_existing_output_conflicts,
    generate_moet_docx,
    generate_tds_docx,
)
from app.telegram_notify import build_notifier


@dataclass(frozen=True)
class BotCommand:
    action: str
    program: str
    grade: int
    week: int
    upload: bool = False


@dataclass(frozen=True)
class PendingOverwriteConfirmation:
    command: BotCommand
    conflicts: list[ExistingOutputConflict]


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

Trợ lý SmartPlan (chat riêng đã liên kết):
- /link CODE — liên kết tài khoản
- /unlink — mở SmartPlan để hủy liên kết
- /confirm ID HASH — xác nhận một yêu cầu
- /cancel ID HASH — hủy một yêu cầu
- /status JOB_ID — xem trạng thái công việc
"""


ACTION_LABELS = {"audit": "kiểm tra Drive", "generate": "soạn và upload giáo án"}
PROGRAM_LABELS = {"tds": "TDS", "moet": "MOET"}


@dataclass(frozen=True)
class AssistantCommand:
    action: str
    args: tuple[str, ...] = ()


def parse_assistant_command(message: str) -> AssistantCommand | None:
    parts = message.strip().split()
    if not parts or not parts[0].startswith("/"):
        return None
    action = parts[0].lower().split("@", 1)[0]
    if action in {"/link", "/unlink", "/confirm", "/cancel", "/status"}:
        return AssistantCommand(action[1:], tuple(parts[1:]))
    return None


def _safe_assistant_result(response: dict) -> str:
    result = response.get("result")
    if isinstance(result, dict):
        if result.get("linked") is True:
            return "Đã liên kết tài khoản Telegram với SmartPlan."
        for key in ("answer", "text"):
            value = result.get(key)
            if isinstance(value, str) and value.strip():
                return redact_assistant_text(value)
        if isinstance(result.get("confirmationId"), str) and isinstance(result.get("requestHash"), str):
            preview = redact_assistant_text(result.get("preview")) or "Yêu cầu cần xác nhận."
            return f"{preview}\nDùng /confirm {result['confirmationId']} {result['requestHash']} để xác nhận, hoặc /cancel {result['confirmationId']} {result['requestHash']} để hủy."
        if isinstance(result.get("status"), str):
            progress = result.get("progress")
            suffix = f" ({progress}%)" if isinstance(progress, (int, float)) else ""
            return f"Trạng thái công việc: {result['status']}{suffix}."
    return "Đã xử lý yêu cầu trợ lý."


def _safe_assistant_error(error: AssistantClientError) -> str:
    if error.code in {"ASSISTANT_NOT_CONFIGURED", "ASSISTANT_UNAVAILABLE"}:
        return "Trợ lý hiện chưa sẵn sàng."
    if error.code in {"UNAUTHORIZED", "INVALID_REQUEST", "CHANNEL_NOT_LINKED", "NOT_LINKED"}:
        return "Không thể xác thực yêu cầu trợ lý. Hãy liên kết lại trong SmartPlan."
    if error.code == "NOT_IMPLEMENTED":
        return "Tính năng trợ lý này chưa được bật."
    return "Không xử lý được yêu cầu trợ lý."


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


def format_conflict_confirmation(command: BotCommand, conflicts: list[ExistingOutputConflict]) -> str:
    lines = [
        f"Phát hiện giáo án cũ đã có sẵn trong Drive cho {command.program.upper()} G{command.grade} tuần {command.week:02d}.",
        "Bot chưa xóa hoặc thay thế file nào.",
        "",
        "Các file đang trùng tên với giáo án chuẩn sắp tạo:",
    ]
    for index, conflict in enumerate(conflicts, start=1):
        link = conflict.web_view_link or build_drive_file_link(conflict.file_id)
        suffix = f" | {link}" if link else ""
        lines.append(f"{index}. {conflict.filename}{suffix}")

    lines.extend([
        "",
        "Nếu muốn thay các file cũ bằng bản mới, hãy trả lời đúng một trong các câu sau:",
        "- xác nhận ghi đè",
        "- đồng ý ghi đè",
        "- thay file cũ",
        "",
        "Nếu không muốn thay, hãy trả lời: hủy",
    ])
    return "\n".join(lines)


def is_overwrite_confirmation(text: str) -> bool:
    normalized = normalize_text(text)
    confirmation_phrases = [
        "xác nhận ghi đè",
        "xac nhan ghi de",
        "đồng ý ghi đè",
        "dong y ghi de",
        "ghi đè",
        "ghi de",
        "thay file cũ",
        "thay file cu",
        "replace",
        "overwrite",
    ]
    return any(phrase in normalized for phrase in confirmation_phrases)


def is_cancel_confirmation(text: str) -> bool:
    normalized = normalize_text(text)
    cancel_phrases = ["hủy", "huy", "không", "khong", "không ghi đè", "khong ghi de", "cancel", "bỏ qua", "bo qua"]
    return any(phrase in normalized for phrase in cancel_phrases)


def format_generation_result(command: BotCommand, generated_files: GeneratedLessonBatch) -> str:
    audit = audit_week(command.program, command.grade, command.week)
    folder_link = audit_folder_link(audit)

    lines = [
        f"Đã soạn và upload giáo án {command.program.upper()} G{command.grade} tuần {command.week:02d}.",
        f"Đã tạo {len(generated_files.items)} bài riêng biệt theo PPCT; mỗi bài có 1 file DOCX và 1 file PDF.",
        f"Thư mục chứa file: {audit.folder_name or 'Chưa xác định'}",
    ]
    if folder_link:
        lines.append(f"Link thư mục: {folder_link}")

    for index, item in enumerate(generated_files.items, start=1):
        docx_link = find_uploaded_file_link(audit, item.docx_path.name)
        pdf_link = find_uploaded_file_link(audit, item.pdf_path.name) if item.pdf_path else ""
        lines.extend([
            "",
            f"Bài {index}: {item.lesson_title or item.docx_path.stem}",
            f"- DOCX: {item.docx_path.name}",
        ])
        if docx_link:
            lines.append(f"  Link DOCX: {docx_link}")
        elif item.uploaded_links and item.uploaded_links.get("docx"):
            lines.append(f"  Link DOCX: {item.uploaded_links['docx']}")
        else:
            lines.append("  Link DOCX: chưa tìm thấy khi kiểm tra lại Drive; hãy mở thư mục tuần ở trên để kiểm tra file vừa upload.")

        if item.pdf_path:
            lines.append(f"- PDF: {item.pdf_path.name}")
            if pdf_link:
                lines.append(f"  Link PDF: {pdf_link}")
            elif item.uploaded_links and item.uploaded_links.get("pdf"):
                lines.append(f"  Link PDF: {item.uploaded_links['pdf']}")
            else:
                lines.append("  Link PDF: chưa tìm thấy khi kiểm tra lại Drive; hãy mở thư mục tuần ở trên để kiểm tra file PDF vừa upload.")

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


def generate_and_upload(command: BotCommand, replace_existing: bool = False) -> GeneratedLessonBatch:
    if command.program == "tds":
        return generate_tds_docx(command.grade, command.week, "dgs", upload=True, notify=False, replace_existing=replace_existing)
    if command.program == "moet":
        return generate_moet_docx(command.grade, command.week, upload=True, notify=False, replace_existing=replace_existing)
    raise ValueError(f"Unsupported program: {command.program}")


def handle_command(command: BotCommand) -> str:
    if command.action == "help":
        return HELP_TEXT

    if command.action == "audit":
        audit = audit_week(command.program, command.grade, command.week)
        return format_week_audit_with_links(audit, command.program.upper())

    if command.action == "generate":
        generated_files = generate_and_upload(command, replace_existing=False)
        return format_generation_result(command, generated_files)

    return HELP_TEXT


class TelegramPollingBot:
    def __init__(self) -> None:
        settings = load_settings()
        require_values(settings, ["telegram_bot_token"])
        self.owner_chat_id = str(settings.telegram_chat_id)
        self.api_base = f"https://api.telegram.org/bot{settings.telegram_bot_token}"
        self.pending_requests: dict[str, ParsedRequest] = {}
        self.pending_overwrites: dict[str, PendingOverwriteConfirmation] = {}
        self.seen_assistant_messages: set[tuple[str, str]] = set()
        self.assistant_account_id = str(getattr(settings, "telegram_account_id", "") or "").strip()
        self.assistant_client = (
            AssistantClient(settings)
            if self.assistant_account_id and settings.smartplan_assistant_url and settings.smartplan_assistant_secret
            else None
        )

    def get_updates(self, offset: int | None) -> list[dict]:
        params: dict[str, int] = {"timeout": 25}
        if offset is not None:
            params["offset"] = offset
        try:
            response = requests.get(f"{self.api_base}/getUpdates", params=params, timeout=35)
            response.raise_for_status()
        except requests.RequestException as exc:
            raise RuntimeError("Telegram transport unavailable") from exc
        payload = response.json()
        if not payload.get("ok"):
            raise RuntimeError("Telegram getUpdates failed")
        return payload.get("result", [])

    def send_message(self, chat_id: str, text: str) -> None:
        for chunk in split_message(text):
            try:
                response = requests.post(
                    f"{self.api_base}/sendMessage",
                    json={"chat_id": chat_id, "text": chunk},
                    timeout=20,
                )
                response.raise_for_status()
            except requests.RequestException as exc:
                raise RuntimeError("Telegram transport unavailable") from exc

    def resolve_overwrite_confirmation(self, chat_id: str, text: str) -> str | None:
        pending = self.pending_overwrites.get(chat_id)
        if not pending:
            return None

        if is_cancel_confirmation(text):
            self.pending_overwrites.pop(chat_id, None)
            return "Đã hủy yêu cầu ghi đè. Bot không xóa, không thay thế và không upload file mới."

        if not is_overwrite_confirmation(text):
            self.send_message(
                chat_id,
                "Mình đang chờ xác nhận ghi đè giáo án cũ. Hãy trả lời 'xác nhận ghi đè' để thay file cũ, hoặc 'hủy' để dừng.",
            )
            return ""

        self.pending_overwrites.pop(chat_id, None)
        self.send_message(
            chat_id,
            f"Đã nhận xác nhận ghi đè {pending.command.program.upper()} G{pending.command.grade} tuần {pending.command.week:02d}. Tôi sẽ đưa file cũ vào thùng rác rồi upload bản mới...",
        )
        generated_files = generate_and_upload(pending.command, replace_existing=True)
        return format_generation_result(pending.command, generated_files)

    def resolve_command(self, chat_id: str, text: str) -> BotCommand | None:
        self.pending_overwrites.pop(chat_id, None)
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

    def _assistant_identity(self, chat_id: str, message: dict) -> AssistantIdentity | None:
        chat = message.get("chat") or {}
        sender = message.get("from") or {}
        if chat.get("type") != "private" or not chat_id.isdigit() or not str(sender.get("id", "")).isdigit() or str(sender.get("id")) != chat_id:
            return None
        if not self.assistant_client or not self.assistant_account_id:
            return None
        return AssistantIdentity(
            channel="telegram",
            channel_account_id=self.assistant_account_id,
            channel_user_id=chat_id,
            channel_chat_id=chat_id,
        )

    @staticmethod
    def _assistant_request_id(message: dict, chat_id: str, text: str, action: str) -> str:
        message_id = str((message.get("message_id") or "")).strip()
        suffix = message_id if message_id.isdigit() else hashlib.sha256(text.encode("utf-8")).hexdigest()[:16]
        return f"telegram-{action}-{chat_id}-{suffix}"[:128]

    def _handle_assistant_message(self, message: dict, chat_id: str, text: str) -> bool:
        if not hasattr(self, "seen_assistant_messages"):
            self.seen_assistant_messages = set()
        message_id = str(message.get("message_id", "")).strip()
        command = parse_assistant_command(text)
        has_legacy_state = chat_id in self.pending_requests or chat_id in self.pending_overwrites
        if has_legacy_state and command and command.action == "cancel":
            return False
        is_natural_language = command is None and not has_legacy_state and not parse_request(text).action and not parse_request(text).program and not parse_request(text).grade and not parse_request(text).week
        if command is None and not is_natural_language:
            return False
        if command and command.action == "help":
            self.send_message(chat_id, HELP_TEXT)
            return True
        identity = self._assistant_identity(chat_id, message)
        if identity is None:
            if command is not None or is_natural_language:
                self.send_message(chat_id, "Trợ lý chỉ nhận lệnh trong chat riêng đã liên kết và đã cấu hình.")
                return True
            return False

        if message_id:
            key = (chat_id, message_id)
            if key in self.seen_assistant_messages:
                return True
            self.seen_assistant_messages.add(key)
            if len(self.seen_assistant_messages) > 10000:
                self.seen_assistant_messages = set(list(self.seen_assistant_messages)[-5000:])

        request_id = self._assistant_request_id(message, chat_id, text, command.action if command else "chat")
        try:
            if command and command.action == "unlink":
                self.send_message(chat_id, "Mở SmartPlan để hủy liên kết trợ lý.")
                return True
            if command and command.action == "link":
                if len(command.args) != 1:
                    self.send_message(chat_id, "Cú pháp: /link CODE")
                    return True
                result = self.assistant_client.send_request("consumeLinkCode", identity, {"code": command.args[0]}, request_id=request_id)
            elif command and command.action in {"confirm", "cancel"}:
                if len(command.args) != 2:
                    self.send_message(chat_id, f"Cú pháp: /{command.action} ID HASH")
                    return True
                method = self.assistant_client.confirm if command.action == "confirm" else self.assistant_client.cancel
                result = method(identity, command.args[0], command.args[1], request_id=request_id)
            elif command and command.action == "status":
                if len(command.args) != 1:
                    self.send_message(chat_id, "Cú pháp: /status JOB_ID")
                    return True
                result = self.assistant_client.get_job_status(identity, command.args[0], request_id=request_id)
            else:
                result = self.assistant_client.send_request("chat", identity, {"text": text}, request_id=request_id)
            self.send_message(chat_id, _safe_assistant_result(result))
        except AssistantClientError as exc:
            self.send_message(chat_id, _safe_assistant_error(exc))
        return True

    def handle_message(self, message: dict) -> None:
        chat = message.get("chat") or {}
        chat_id = str(chat.get("id", ""))
        if not chat_id:
            return

        text = (message.get("text") or "").strip()
        if not text:
            return

        try:
            if self._handle_assistant_message(message, chat_id, text):
                return

            overwrite_result = self.resolve_overwrite_confirmation(chat_id, text)
            if overwrite_result is not None:
                if overwrite_result:
                    self.send_message(chat_id, overwrite_result)
                return

            command = self.resolve_command(chat_id, text)
            if not command:
                return

            if command.action == "generate":
                conflicts = find_existing_output_conflicts(command.program, command.grade, command.week)
                if conflicts:
                    self.pending_overwrites[chat_id] = PendingOverwriteConfirmation(command=command, conflicts=conflicts)
                    self.send_message(chat_id, format_conflict_confirmation(command, conflicts))
                    return

            if command.action != "help":
                action_label = ACTION_LABELS.get(command.action, command.action)
                self.send_message(
                    chat_id,
                    f"Đã nhận lệnh {action_label} {command.program.upper()} G{command.grade} tuần {command.week:02d}. Tôi đang xử lý...",
                )
            result = handle_command(command)
        except Exception:
            result = "Có lỗi khi xử lý lệnh. Vui lòng thử lại."
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
                    message = update.get("message")
                    if message:
                        self.handle_message(message)
            except Exception as exc:
                print(f"Telegram polling error: {type(exc).__name__}")
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
