from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd
from pypdf import PdfReader

from app.config import BASE_DIR, load_settings, require_values
from app.drive_client import GoogleDriveClient


DOWNLOAD_DIR = BASE_DIR / "outputs" / "ppct"
TDS_EXCEL_PATH = DOWNLOAD_DIR / "PPCT_TDS_25_26.xlsx"


PPCT_REQUIRED_SETTINGS = [
    "ppct_tds_excel_file_id",
    "ppct_moet_g10_pdf_file_id",
    "ppct_moet_g11_pdf_file_id",
    "ppct_moet_g12_pdf_file_id",
]


@dataclass(frozen=True)
class LessonItem:
    subject: str
    period: str
    content: str


@dataclass(frozen=True)
class TDSWeekPlan:
    grade: int
    week: int
    week_label: str
    month: str
    track: str
    notes: str
    lessons: list[LessonItem]


def clean_cell(value: Any) -> str:
    if pd.isna(value):
        return ""
    return "\n".join(line.strip() for line in str(value).splitlines() if line.strip())


def download_ppct_files() -> dict[str, Path]:
    settings = load_settings()
    require_values(settings, PPCT_REQUIRED_SETTINGS)
    client = GoogleDriveClient()

    files = {
        "tds_excel": (settings.ppct_tds_excel_file_id, DOWNLOAD_DIR / "PPCT_TDS_25_26.xlsx"),
        "moet_g10_pdf": (settings.ppct_moet_g10_pdf_file_id, DOWNLOAD_DIR / "PPCT_MOET_G10.pdf"),
        "moet_g11_pdf": (settings.ppct_moet_g11_pdf_file_id, DOWNLOAD_DIR / "PPCT_MOET_G11.pdf"),
        "moet_g12_pdf": (settings.ppct_moet_g12_pdf_file_id, DOWNLOAD_DIR / "PPCT_MOET_G12.pdf"),
    }

    downloaded: dict[str, Path] = {}
    for key, (file_id, output_path) in files.items():
        print(f"Downloading {key} -> {output_path}")
        downloaded[key] = client.download_file(file_id=file_id, output_path=output_path)

    return downloaded


def inspect_tds_excel(path: Path, sheet_names: list[str] | None = None, rows: int = 30) -> None:
    excel = pd.ExcelFile(path)
    selected_sheets = sheet_names or excel.sheet_names[:5]
    print(f"TDS Excel sheets: {excel.sheet_names}")
    for sheet_name in selected_sheets:
        frame = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=rows)
        print(f"\nSheet: {sheet_name}")
        print(f"Shape preview: {frame.shape[0]} rows x {frame.shape[1]} columns")
        print(frame.to_string(index=True, header=False))


def tds_columns_for_grade(grade: int, track: str = "dgs") -> tuple[int, int, int, int | None]:
    normalized_track = track.lower().strip()
    if grade == 10:
        return (6, 7, 8, 5) if normalized_track == "dgs" else (3, 4, 5, 5)
    if grade == 11:
        return (2, 5, 7, 4) if normalized_track == "dgs" else (2, 3, 7, 4)
    if grade == 12:
        return (2, 6, 7, 5) if normalized_track == "dgs" else (2, 3, 4, 5)
    raise ValueError(f"Unsupported TDS grade: {grade}")


def extract_tds_week(path: Path, grade: int, week: int, track: str = "dgs") -> TDSWeekPlan:
    sheet_name = f"G{grade}"
    frame = pd.read_excel(path, sheet_name=sheet_name, header=None)
    subject_col, period_col, content_col, notes_col = tds_columns_for_grade(grade, track)

    active_month = ""
    active_week = 0
    active_week_label = ""
    active_subject = ""
    active_notes = ""
    target_month = ""
    target_week_label = ""
    target_notes = ""
    lessons: list[LessonItem] = []

    for _, row in frame.iterrows():
        month_value = clean_cell(row.iloc[0]) if len(row) > 0 else ""
        week_value = clean_cell(row.iloc[1]) if len(row) > 1 else ""
        subject_value = clean_cell(row.iloc[subject_col]) if len(row) > subject_col else ""
        period_value = clean_cell(row.iloc[period_col]) if len(row) > period_col else ""
        content_value = clean_cell(row.iloc[content_col]) if len(row) > content_col else ""
        notes_value = clean_cell(row.iloc[notes_col]) if notes_col is not None and len(row) > notes_col else ""

        if month_value:
            active_month = month_value
        if notes_value:
            active_notes = notes_value
        if subject_value:
            active_subject = subject_value
        if week_value:
            first_token = week_value.split()[0]
            if first_token.isdigit():
                active_week = int(first_token)
                active_week_label = week_value

        if active_week == week:
            target_month = active_month
            target_week_label = active_week_label
            target_notes = active_notes
            if period_value or content_value:
                lessons.append(
                    LessonItem(
                        subject=active_subject,
                        period=period_value,
                        content=content_value,
                    )
                )
        elif active_week > week and lessons:
            break

    if not lessons:
        raise RuntimeError(f"No TDS lessons found for grade {grade}, week {week}, track {track}")

    return TDSWeekPlan(
        grade=grade,
        week=week,
        week_label=target_week_label,
        month=target_month,
        track=track,
        notes=target_notes,
        lessons=lessons,
    )


def print_tds_week_plan(plan: TDSWeekPlan) -> None:
    print(f"TDS G{plan.grade} - Week {plan.week}: {plan.week_label}")
    print(f"Month: {plan.month}")
    print(f"Track: {plan.track}")
    if plan.notes:
        print(f"Notes: {plan.notes}")
    print("Lessons:")
    for index, lesson in enumerate(plan.lessons, start=1):
        print(f"\n{index}. {lesson.subject}")
        print(f"   Period: {lesson.period}")
        print(f"   Content: {lesson.content}")


def inspect_pdf(path: Path, label: str) -> None:
    reader = PdfReader(str(path))
    print(f"\n{label}: {len(reader.pages)} pages")
    if reader.pages:
        text = reader.pages[0].extract_text() or ""
        preview = "\n".join(text.splitlines()[:20])
        print(f"First page preview:\n{preview}")


def test_download_and_inspect() -> None:
    downloaded = download_ppct_files()
    print("\nDownloaded PPCT files:")
    for key, path in downloaded.items():
        print(f"- {key}: {path} ({path.stat().st_size} bytes)")

    inspect_tds_excel(downloaded["tds_excel"])
    inspect_pdf(downloaded["moet_g10_pdf"], "Moet G10 PDF")
    inspect_pdf(downloaded["moet_g11_pdf"], "Moet G11 PDF")
    inspect_pdf(downloaded["moet_g12_pdf"], "Moet G12 PDF")


def main() -> None:
    parser = argparse.ArgumentParser(description="PPCT downloader and parser")
    parser.add_argument("--download", action="store_true", help="Download PPCT source files")
    parser.add_argument("--inspect", action="store_true", help="Download and inspect PPCT source files")
    parser.add_argument(
        "--inspect-tds-grades",
        action="store_true",
        help="Inspect TDS Excel sheets G10, G11, and G12 without headers",
    )
    parser.add_argument(
        "--rows",
        type=int,
        default=40,
        help="Number of rows to print when inspecting Excel sheets",
    )
    parser.add_argument("--extract-tds-week", action="store_true", help="Extract one TDS weekly plan")
    parser.add_argument("--grade", type=int, choices=[10, 11, 12], default=10, help="Grade for TDS extraction")
    parser.add_argument("--week", type=int, default=1, help="Week number for TDS extraction")
    parser.add_argument("--track", default="dgs", choices=["dgs", "discover"], help="TDS track to extract")
    args = parser.parse_args()

    if args.inspect:
        test_download_and_inspect()
    elif args.inspect_tds_grades:
        inspect_tds_excel(TDS_EXCEL_PATH, ["G10", "G11", "G12"], args.rows)
    elif args.download:
        downloaded = download_ppct_files()
        for key, path in downloaded.items():
            print(f"{key}: {path}")
    elif args.extract_tds_week:
        plan = extract_tds_week(TDS_EXCEL_PATH, args.grade, args.week, args.track)
        print_tds_week_plan(plan)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
