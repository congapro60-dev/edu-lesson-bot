from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path

from pypdf import PdfReader

from app.config import BASE_DIR
from app.ppct_parser import LessonItem


MOET_PDF_PATHS = {
    10: BASE_DIR / "outputs" / "ppct" / "PPCT_MOET_G10.pdf",
    11: BASE_DIR / "outputs" / "ppct" / "PPCT_MOET_G11.pdf",
    12: BASE_DIR / "outputs" / "ppct" / "PPCT_MOET_G12.pdf",
}


@dataclass(frozen=True)
class MoetWeekPlan:
    grade: int
    week: int
    lessons: list[LessonItem]


def clean_text(value: str) -> str:
    return "\n".join(line.strip() for line in value.splitlines() if line.strip())


def extract_pdf_text(path: Path) -> str:
    reader = PdfReader(str(path))
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def find_distribution_section(text: str) -> str:
    markers = [
        "Tuần Tiết Bài học Yêu cầu cần đạt",
        "Tuần\nTiết\nBài học",
        "Phân phối chương trình",
        "PHÂN PHỐI CHƯƠNG TRÌNH",
    ]
    starts = [text.find(marker) for marker in markers if text.find(marker) >= 0]
    if not starts:
        return text
    return text[min(starts):]


def parse_numbered_lessons(text: str) -> list[LessonItem]:
    section = find_distribution_section(text)
    lesson_keywords = [
        "Mệnh đề",
        "Tập hợp",
        "Giá trị lượng giác",
        "Hệ thức lượng",
        "Bất phương trình",
        "Hàm số",
        "Vectơ",
        "Vector",
        "Thống kê",
        "Xác suất",
        "Đạo hàm",
        "Nguyên hàm",
        "Tích phân",
        "Phương trình",
        "Đường thẳng",
        "Mặt phẳng",
        "Luyện tập",
        "Ôn tập",
        "Kiểm tra",
        "Hoạt động",
    ]
    keyword_regex = "|".join(lesson_keywords)
    normalized_section = re.sub(rf"\s+(\d{{1,3}})\s+({keyword_regex})", r"\n\1 \2", section)
    lines = [line.strip() for line in normalized_section.splitlines() if line.strip()]
    lessons: list[LessonItem] = []
    current_period = ""
    current_content: list[str] = []

    period_pattern = re.compile(r"^(\d{1,3})\s+(.+)$", re.IGNORECASE)
    lesson_title_keywords = (
        "mệnh đề",
        "tập hợp",
        "giá trị lượng giác",
        "hệ thức lượng",
        "bất phương trình",
        "hàm số",
        "vectơ",
        "vector",
        "thống kê",
        "xác suất",
        "đạo hàm",
        "nguyên hàm",
        "tích phân",
        "phương trình",
        "đường thẳng",
        "mặt phẳng",
        "ôn tập",
        "kiểm tra",
        "hoạt động",
    )

    for line in lines:
        normalized = line.lower()
        if any(skip in normalized for skip in ["số lớp", "tình hình đội ngũ", "thiết bị dạy học", "phòng học", "phân phối chương trình", "tuần tiết"]):
            continue
        match = period_pattern.match(line)
        if match:
            number = int(match.group(1))
            tail = match.group(2).strip()
            tail_lower = tail.lower()
            if 1 <= number <= 140 and len(tail) >= 3 and any(keyword in tail_lower for keyword in lesson_title_keywords):
                if current_period and current_content:
                    lessons.append(
                        LessonItem(subject="Toán", period=current_period, content=clean_text("\n".join(current_content)))
                    )
                current_period = str(number)
                current_content = [tail]
                continue

        if current_period and len(line) > 3:
            current_content.append(line)

    if current_period and current_content:
        lessons.append(LessonItem(subject="Toán", period=current_period, content=clean_text("\n".join(current_content))))

    expanded_lessons: list[LessonItem] = []
    embedded_pattern = re.compile(rf"\b(\d{{1,3}})\s+({keyword_regex}[^\n]*)")
    for lesson in lessons:
        content = lesson.content
        embedded_matches = list(embedded_pattern.finditer(content))
        expanded_lessons.append(lesson)
        for match in embedded_matches:
            embedded_period = int(match.group(1))
            if embedded_period != int(lesson.period) and 1 <= embedded_period <= 140:
                expanded_lessons.append(
                    LessonItem(
                        subject="Toán",
                        period=str(embedded_period),
                        content=clean_text(match.group(2)),
                    )
                )

    lessons = sorted(expanded_lessons, key=lambda item: int(item.period))

    deduped: list[LessonItem] = []
    seen: set[tuple[str, str]] = set()
    for lesson in lessons:
        key = (lesson.period, lesson.content[:80])
        if key not in seen:
            seen.add(key)
            deduped.append(lesson)
    return deduped


def extract_moet_all_lessons(grade: int) -> list[LessonItem]:
    path = MOET_PDF_PATHS[grade]
    if not path.exists():
        from app.ppct_parser import download_ppct_files

        download_ppct_files()
    if not path.exists():
        raise FileNotFoundError(f"PPCT MOET file is missing after download attempt: {path}")

    text = extract_pdf_text(path)
    lessons = parse_numbered_lessons(text)
    if not lessons:
        raise RuntimeError(f"No Moet lessons parsed for grade {grade} from {path}")
    return lessons


def extract_moet_week(grade: int, week: int, periods_per_week: int = 5) -> MoetWeekPlan:
    lessons = extract_moet_all_lessons(grade)
    start = (week - 1) * periods_per_week
    end = start + periods_per_week
    week_lessons = lessons[start:end]
    if not week_lessons:
        raise RuntimeError(f"No Moet lessons found for grade {grade}, week {week}")
    return MoetWeekPlan(grade=grade, week=week, lessons=week_lessons)


def print_moet_week(plan: MoetWeekPlan) -> None:
    print(f"Moet G{plan.grade} - Week {plan.week:02d}")
    print(f"Lessons: {len(plan.lessons)}")
    for index, lesson in enumerate(plan.lessons, start=1):
        print(f"\n{index}. Period {lesson.period}")
        print(lesson.content)


def inspect_moet(grade: int, limit: int = 20) -> None:
    lessons = extract_moet_all_lessons(grade)
    print(f"Moet G{grade}: parsed {len(lessons)} numbered lessons")
    for lesson in lessons[:limit]:
        print(f"\nPeriod {lesson.period}: {lesson.content[:500]}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Parse Moet PPCT PDF files")
    parser.add_argument("--inspect", action="store_true", help="Inspect parsed Moet lessons")
    parser.add_argument("--week", type=int, help="Extract one Moet week")
    parser.add_argument("--grade", type=int, choices=[10, 11, 12], default=10)
    parser.add_argument("--limit", type=int, default=20)
    args = parser.parse_args()

    if args.inspect:
        inspect_moet(args.grade, args.limit)
    elif args.week:
        print_moet_week(extract_moet_week(args.grade, args.week))
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
