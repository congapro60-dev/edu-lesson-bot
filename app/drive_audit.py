from __future__ import annotations

import argparse
import re
import unicodedata
from dataclasses import dataclass

from app.config import load_settings
from app.drive_client import GoogleDriveClient
from app.moet_parser import extract_moet_week
from app.ppct_parser import TDS_EXCEL_PATH, extract_tds_week
from app.telegram_notify import build_notifier


PROGRAM_LABELS = {"tds": "TDS", "moet": "MOET"}


DOCX_MIME_TYPES = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/vnd.google-apps.document",
}


@dataclass(frozen=True)
class WeekAudit:
    grade: int
    week: int
    week_label: str
    folder_id: str | None
    folder_name: str | None
    expected_lessons: int
    existing_files: list[dict[str, str]]

    @property
    def missing_count(self) -> int:
        return max(self.expected_lessons - len(self.existing_files), 0)


def tds_grade_folder_id(grade: int) -> str:
    settings = load_settings()
    if grade == 10:
        return settings.tds_g10_folder_id
    if grade == 11:
        return settings.tds_g11_folder_id
    if grade == 12:
        return settings.tds_g12_folder_id
    raise ValueError(f"Unsupported TDS grade: {grade}")


def moet_grade_folder_id(grade: int) -> str:
    settings = load_settings()
    if grade == 10:
        return settings.moet_g10_folder_id
    if grade == 11:
        return settings.moet_g11_folder_id
    if grade == 12:
        return settings.moet_g12_folder_id
    raise ValueError(f"Unsupported Moet grade: {grade}")


def normalize_week_folder_name(folder_name: str) -> str:
    normalized = unicodedata.normalize("NFC", folder_name).lower().strip()
    normalized = normalized.replace("\u0300", "")
    return re.sub(r"\s+", " ", normalized)


def extract_week_number(folder_name: str) -> int | None:
    normalized = normalize_week_folder_name(folder_name)
    match = re.search(r"(?:tuần|tuan)\s*0*(\d{1,2})(?:\b|$)", normalized)
    if not match:
        return None
    return int(match.group(1))


def find_week_folder(client: GoogleDriveClient, grade_folder_id: str, week: int) -> dict[str, str] | None:
    candidates = client.list_files(
        query=(
            f"'{grade_folder_id}' in parents and "
            "mimeType = 'application/vnd.google-apps.folder' and "
            "trashed = false"
        ),
        page_size=100,
    )
    for folder in candidates:
        if extract_week_number(folder.get("name", "")) == week:
            return folder
    return None


def list_docx_files(client: GoogleDriveClient, folder_id: str) -> list[dict[str, str]]:
    files = client.list_files(query=f"'{folder_id}' in parents and trashed = false", page_size=100)
    return [
        file
        for file in files
        if file.get("mimeType") in DOCX_MIME_TYPES or file.get("name", "").lower().endswith(".docx")
    ]


def audit_tds_week(grade: int, week: int) -> WeekAudit:
    client = GoogleDriveClient()
    grade_folder_id = tds_grade_folder_id(grade)
    expected_plan = extract_tds_week(TDS_EXCEL_PATH, grade, week, "dgs")
    week_folder = find_week_folder(client, grade_folder_id, week)

    existing_files: list[dict[str, str]] = []
    if week_folder:
        existing_files = list_docx_files(client, week_folder["id"])

    return WeekAudit(
        grade=grade,
        week=week,
        week_label=expected_plan.week_label,
        folder_id=week_folder.get("id") if week_folder else None,
        folder_name=week_folder.get("name") if week_folder else None,
        expected_lessons=len(expected_plan.lessons),
        existing_files=existing_files,
    )


def audit_moet_week(grade: int, week: int) -> WeekAudit:
    client = GoogleDriveClient()
    grade_folder_id = moet_grade_folder_id(grade)
    expected_plan = extract_moet_week(grade, week)
    week_folder = find_week_folder(client, grade_folder_id, week)

    existing_files: list[dict[str, str]] = []
    if week_folder:
        existing_files = list_docx_files(client, week_folder["id"])

    return WeekAudit(
        grade=grade,
        week=week,
        week_label=f"Tuần {week:02d}",
        folder_id=week_folder.get("id") if week_folder else None,
        folder_name=week_folder.get("name") if week_folder else None,
        expected_lessons=len(expected_plan.lessons),
        existing_files=existing_files,
    )


def format_week_audit(audit: WeekAudit, program: str = "TDS") -> str:
    lines = [
        f"{program} G{audit.grade} - Tuần {audit.week:02d}: {audit.week_label}",
        f"Thư mục tuần: {audit.folder_name or 'Chưa có'}",
        f"Số tiết/nội dung theo PPCT: {audit.expected_lessons}",
        f"Số file giáo án hiện có: {len(audit.existing_files)}",
        f"Ước lượng còn thiếu: {audit.missing_count}",
    ]
    if audit.existing_files:
        lines.append("File hiện có:")
        for file in audit.existing_files:
            lines.append(f"- {file.get('name')} | {file.get('webViewLink', '')}")
    if audit.missing_count:
        lines.append("Đề xuất: cần soạn/bổ sung giáo án cho các nội dung còn thiếu theo PPCT tuần này.")
    else:
        lines.append("Đề xuất: tuần này có vẻ đã đủ số lượng file so với PPCT.")
    return "\n".join(lines)


def send_report_if_requested(report_text: str, notify: bool) -> None:
    if not notify:
        return
    chunks = [report_text[index : index + 3500] for index in range(0, len(report_text), 3500)]
    notifier = build_notifier()
    for chunk in chunks:
        notifier.send_message(chunk)
    print("Telegram audit notification sent.")


def audit_tds_range(start_week: int, end_week: int, notify: bool = False) -> str:
    reports: list[str] = []
    for grade in [10, 11, 12]:
        for week in range(start_week, end_week + 1):
            reports.append(format_week_audit(audit_tds_week(grade, week), "TDS"))

    report_text = "\n\n".join(reports)
    print(report_text)

    send_report_if_requested(report_text, notify)

    return report_text


def audit_week(program: str, grade: int, week: int) -> WeekAudit:
    normalized = program.lower().strip()
    if normalized == "tds":
        return audit_tds_week(grade, week)
    if normalized == "moet":
        return audit_moet_week(grade, week)
    raise ValueError(f"Unsupported program: {program}")


def audit_all_range(programs: list[str], start_week: int, end_week: int, notify: bool = False) -> str:
    reports: list[str] = []
    for program in programs:
        label = PROGRAM_LABELS[program]
        for grade in [10, 11, 12]:
            for week in range(start_week, end_week + 1):
                reports.append(format_week_audit(audit_week(program, grade, week), label))

    report_text = "\n\n".join(reports)
    print(report_text)
    send_report_if_requested(report_text, notify)
    return report_text


def collect_missing_weeks(
    programs: list[str],
    start_week: int,
    end_week: int,
    generate_partial_weeks: bool = False,
) -> list[tuple[str, int, int, WeekAudit]]:
    missing: list[tuple[str, int, int, WeekAudit]] = []
    for program in programs:
        for grade in [10, 11, 12]:
            for week in range(start_week, end_week + 1):
                audit = audit_week(program, grade, week)
                if audit.missing_count <= 0:
                    continue
                if audit.existing_files and not generate_partial_weeks:
                    continue
                missing.append((program, grade, week, audit))
    return missing


def audit_moet_range(start_week: int, end_week: int, notify: bool = False) -> str:
    reports: list[str] = []
    for grade in [10, 11, 12]:
        for week in range(start_week, end_week + 1):
            reports.append(format_week_audit(audit_moet_week(grade, week), "MOET"))

    report_text = "\n\n".join(reports)
    print(report_text)

    send_report_if_requested(report_text, notify)

    return report_text


def main() -> None:
    parser = argparse.ArgumentParser(description="Audit existing lesson plans in Google Drive")
    parser.add_argument("--tds", action="store_true", help="Audit TDS output folders")
    parser.add_argument("--moet", action="store_true", help="Audit Moet output folders")
    parser.add_argument("--start-week", type=int, default=1)
    parser.add_argument("--end-week", type=int, default=1)
    parser.add_argument("--notify", action="store_true", help="Send audit report to Telegram")
    args = parser.parse_args()

    if args.tds:
        audit_tds_range(args.start_week, args.end_week, args.notify)
    elif args.moet:
        audit_moet_range(args.start_week, args.end_week, args.notify)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
