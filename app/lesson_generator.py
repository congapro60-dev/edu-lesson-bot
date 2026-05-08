from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path

from app.drive_audit import PROGRAM_LABELS, audit_all_range, collect_missing_weeks, find_week_folder

from anthropic import Anthropic
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from app.config import BASE_DIR, load_settings, require_values
from app.drive_client import GoogleDriveClient
from app.math_docx import add_math_aware_paragraph
from app.moet_parser import MoetWeekPlan, extract_moet_week
from app.pdf_renderer import render_pdf
from app.ppct_parser import TDS_EXCEL_PATH, TDSWeekPlan, extract_tds_week
from app.telegram_notify import build_notifier
from app.web_lesson_client import render_docx_with_web, web_render_enabled


GENERATED_DIR = BASE_DIR / "outputs" / "generated"
TEMPLATE_DIR = BASE_DIR / "outputs" / "templates"
TDS_TEMPLATE_FILE_ID = "152TEJ_7P0RrqQbtk14TlQvPxJ8peNL2g"
TDS_TEMPLATE_DOCX_PATH = TEMPLATE_DIR / "mau_giao_an_tds.docx"
DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
PDF_MIME_TYPE = "application/pdf"


@dataclass(frozen=True)
class GeneratedLessonFiles:
    docx_path: Path
    pdf_path: Path | None = None


LESSON_SYSTEM_PROMPT = """Bạn là chuyên gia thiết kế kế hoạch dạy học môn Toán THPT theo chuẩn Ban Toán TDS.
Nhiệm vụ của bạn là tạo giáo án dùng được ngay, chi tiết theo hoạt động lớp học, không viết chung chung.

Nguyên tắc bắt buộc:
1. Bám tuyệt đối PPCT được cung cấp; không bịa bài học, tài liệu, đường dẫn.
2. Giáo án phải có đúng 4 bước TDS:
   - KHỞI ĐỘNG/TRẢI NGHIỆM (tối đa 5 phút)
   - HÌNH THÀNH KIẾN THỨC
   - RÈN LUYỆN CỦNG CỐ
   - SƠ KẾT + BTVN
3. Mỗi hoạt động phải có bảng Markdown đúng 3 cột:
   | Thời gian thực | Giáo viên và Học sinh | Nội dung |
4. Thời gian phải ghi theo giờ thực, ví dụ 8h00-8h05, không chỉ ghi "5 phút".
5. Hoạt động của GV phải là câu hỏi dẫn dắt, scaffolding, phản biện; tránh thuyết trình áp đặt.
6. Phải có dự kiến câu trả lời cụ thể của HS và phương án hỗ trợ.
7. Phải phân hóa rõ 3 đối tượng: HS yếu / HS đại trà / HS giỏi.
8. Phải gắn 5 năng lực cốt lõi Toán học vào hoạt động cụ thể: tư duy và lập luận; mô hình hóa; giải quyết vấn đề; giao tiếp toán học; sử dụng công cụ/phương tiện.
9. Phần RÈN LUYỆN CỦNG CỐ phải có tối thiểu 3 ví dụ/bài tập từ dễ đến khó, gồm cơ bản, nâng cao, thách thức; đa dạng câu hỏi xuôi/ngược, có/không.
10. Công thức Toán bắt buộc dùng LaTeX chuẩn:
    - Inline: $x^2 + 2x + 1 = 0$
    - Display: $$\\Delta=b^2-4ac$$
    - Dùng \\frac{a}{b}, \\sqrt{x}, x^{2}, a_{n}, \\sin x, \\cos x, \\vec{u}.
11. Không dùng ký tự | trong công thức Toán vì sẽ làm vỡ bảng Markdown; dùng \\mid nếu cần.
12. Trả về nội dung Markdown thuần của giáo án, không bọc trong ```markdown, không giải thích ngoài giáo án."""


def tds_grade_output_folder_id(grade: int) -> str:
    settings = load_settings()
    if grade == 10:
        return settings.tds_g10_folder_id
    if grade == 11:
        return settings.tds_g11_folder_id
    if grade == 12:
        return settings.tds_g12_folder_id
    raise ValueError(f"Unsupported grade: {grade}")


def moet_grade_output_folder_id(grade: int) -> str:
    settings = load_settings()
    if grade == 10:
        return settings.moet_g10_folder_id
    if grade == 11:
        return settings.moet_g11_folder_id
    if grade == 12:
        return settings.moet_g12_folder_id
    raise ValueError(f"Unsupported grade: {grade}")


def build_lesson_prompt(plan: TDSWeekPlan | MoetWeekPlan, program: str = "TDS") -> str:
    lessons_text = "\n".join(
        f"- Môn/chủ đề: {lesson.subject}\n  Tiết: {lesson.period}\n  Nội dung: {lesson.content}"
        for lesson in plan.lessons
    )
    week_label = getattr(plan, "week_label", f"Tuần {plan.week:02d}")
    month = getattr(plan, "month", "")
    track = getattr(plan, "track", program)
    notes = getattr(plan, "notes", "Không có")
    return f"""Hãy soạn giáo án tuần môn Toán {program} theo chuẩn TDS, dùng trực tiếp trong lớp học.

THÔNG TIN PPCT:
- Khối: {plan.grade}
- Tuần: {week_label}
- Tháng: {month or "Không có"}
- Hệ/chương trình: {track}
- Ghi chú tuần: {notes or "Không có"}
- Số tiết trong tuần: {len(plan.lessons)}

NỘI DUNG PPCT CẦN SOẠN:
{lessons_text}

YÊU CẦU ĐẦU RA BẮT BUỘC:

# KẾ HOẠCH DẠY HỌC MÔN TOÁN {program} - KHỐI {plan.grade} - {week_label}

## I. THÔNG TIN CHUNG
- Môn học: Toán
- Khối: {plan.grade}
- Tuần: {week_label}
- Thời lượng: {len(plan.lessons)} tiết
- Nội dung PPCT: tóm tắt đúng các tiết được cung cấp.

## II. MỤC TIÊU HỌC TẬP PHÂN HÓA
Viết rõ 3 nhóm:
- HS yếu: mục tiêu tối thiểu, thao tác cơ bản, hỗ trợ cần có.
- HS đại trà: mục tiêu chuẩn cần đạt, dạng bài trọng tâm.
- HS giỏi: mục tiêu mở rộng, phản biện, khái quát hóa hoặc vận dụng mới.

## III. NĂNG LỰC TOÁN HỌC GẮN VỚI HOẠT ĐỘNG
Liệt kê đủ 5 năng lực, mỗi năng lực phải nói rõ xuất hiện ở hoạt động nào:
1. Tư duy và lập luận toán học
2. Mô hình hóa toán học
3. Giải quyết vấn đề toán học
4. Giao tiếp toán học
5. Sử dụng công cụ và phương tiện học toán

## IV. CHUẨN BỊ
- Giáo viên
- Học sinh
- Công cụ/phương tiện nếu phù hợp

## V. TIẾN TRÌNH DẠY HỌC
Với mỗi tiết trong PPCT, viết đủ 4 bước sau. Mỗi bước PHẢI có bảng Markdown 3 cột chính xác:
| Thời gian thực | Giáo viên và Học sinh | Nội dung |
|---|---|---|

Quy ước thời gian:
- Giả định tiết học bắt đầu lúc 8h00 nếu PPCT không nêu giờ.
- Ghi giờ thực theo khoảng, ví dụ 8h00-8h05, 8h05-8h20.
- KHỞI ĐỘNG/TRẢI NGHIỆM tối đa 5 phút.

BƯỚC 1. KHỞI ĐỘNG/TRẢI NGHIỆM
- Tối đa 5 phút.
- Kết nối kiến thức cũ, khơi gợi tò mò.
- Ưu tiên trải nghiệm/tình huống/câu hỏi dẫn thẳng vào kiến thức mới.

BƯỚC 2. HÌNH THÀNH KIẾN THỨC
- Không giảng áp đặt.
- Dùng chuỗi câu hỏi định hướng tư duy bậc cao: phân tích, so sánh, tổng hợp, phản biện.
- Có dự kiến câu trả lời của HS.
- Có ví dụ minh họa đơn giản, tránh phép tính cồng kềnh.
- Công thức phải viết bằng LaTeX: $x^2 + 2x + 1 = 0$, $$\\Delta=b^2-4ac$$.

BƯỚC 3. RÈN LUYỆN CỦNG CỐ
- Tối thiểu 3 ví dụ/bài tập.
- Bài 1 cơ bản cho HS yếu.
- Bài 2 chuẩn cho HS đại trà.
- Bài 3 nâng cao/thách thức cho HS giỏi.
- Đa dạng cách hỏi: câu hỏi xuôi, câu hỏi ngược, câu hỏi có/không kèm giải thích.

BƯỚC 4. SƠ KẾT + BTVN
- HS tự đánh giá mức đạt mục tiêu.
- GV chốt lỗi thường gặp.
- BTVN rõ ràng, phân hóa 3 mức: cơ bản / chuẩn / nâng cao.

## VI. KIỂM TRA NHANH TIÊU CHÍ
Cuối giáo án thêm checklist ngắn xác nhận:
- Đủ 4 bước TDS.
- Có giờ thực.
- Có phân hóa HS yếu/đại trà/giỏi.
- Có 5 năng lực Toán gắn hoạt động.
- Có tối thiểu 3 ví dụ rèn luyện.
- Công thức dùng LaTeX.
"""


def generate_fallback_lesson_text(plan: TDSWeekPlan | MoetWeekPlan, reason: str, program: str = "TDS") -> str:
    week_label = getattr(plan, "week_label", f"Tuần {plan.week}")
    track = getattr(plan, "track", program)
    lesson_sections = []
    for lesson in plan.lessons:
        lesson_sections.append(
            f"""## Tiết {lesson.period}: {lesson.content}

### BƯỚC 1. KHỞI ĐỘNG/TRẢI NGHIỆM
| Thời gian thực | Giáo viên và Học sinh | Nội dung |
|---|---|---|
| 8h00-8h05 | GV nêu tình huống gắn với kiến thức cũ và hỏi: "Điều gì trong tình huống này khiến ta cần kiến thức mới?"<br/>HS dự đoán, nêu hiểu biết ban đầu. | Câu hỏi khởi động dẫn vào nội dung: {lesson.content}. |

### BƯỚC 2. HÌNH THÀNH KIẾN THỨC
| Thời gian thực | Giáo viên và Học sinh | Nội dung |
|---|---|---|
| 8h05-8h20 | GV đặt câu hỏi phân tích, so sánh và phản biện để HS tự rút ra kiến thức.<br/>HS yếu trả lời câu hỏi nhận biết; HS đại trà giải thích quy trình; HS giỏi nêu phản ví dụ hoặc mở rộng. | Kiến thức trọng tâm của tiết {lesson.period}.<br/>Ví dụ công thức minh họa nếu cần: $x^2 + 2x + 1 = 0$. |

### BƯỚC 3. RÈN LUYỆN CỦNG CỐ
| Thời gian thực | Giáo viên và Học sinh | Nội dung |
|---|---|---|
| 8h20-8h28 | GV giao Bài 1 cơ bản, hỗ trợ HS yếu bằng gợi ý từng bước.<br/>HS trình bày thao tác nền tảng. | Bài 1 cơ bản: nhận diện và áp dụng trực tiếp kiến thức. |
| 8h28-8h35 | GV giao Bài 2 chuẩn, yêu cầu HS giải thích vì sao chọn cách làm.<br/>HS đại trà trình bày lời giải và kiểm tra kết quả. | Bài 2 chuẩn: bài tập thông hiểu/vận dụng trực tiếp. |
| 8h35-8h40 | GV giao Bài 3 nâng cao dạng ngược hoặc có/không kèm giải thích.<br/>HS giỏi phản biện và khái quát hóa. | Bài 3 nâng cao: câu hỏi ngược/có-không, yêu cầu lập luận. |

### BƯỚC 4. SƠ KẾT + BTVN
| Thời gian thực | Giáo viên và Học sinh | Nội dung |
|---|---|---|
| 8h40-8h45 | GV yêu cầu HS tự đánh giá: "Em đã đạt mục tiêu nào? Còn vướng ở đâu?"<br/>HS ghi phiếu tự đánh giá và nhận BTVN phân hóa. | BTVN: cơ bản cho HS yếu; chuẩn cho HS đại trà; nâng cao cho HS giỏi. |"""
        )

    return f"""# KẾ HOẠCH DẠY HỌC MÔN TOÁN {program} - KHỐI {plan.grade} - {week_label}

> Bản fallback do chưa gọi được API sinh nội dung chi tiết. Lý do: {reason}

## I. THÔNG TIN CHUNG
- Môn học: Toán
- Khối: {plan.grade}
- Tuần: {week_label}
- Hệ/chương trình: {track}
- Thời lượng: {len(plan.lessons)} tiết

## II. MỤC TIÊU HỌC TẬP PHÂN HÓA
- HS yếu: nhận diện được dạng toán và thực hiện thao tác cơ bản với gợi ý.
- HS đại trà: giải được bài tập chuẩn, giải thích được quy trình và kiểm tra kết quả.
- HS giỏi: xử lí câu hỏi mở rộng, phản biện cách làm và khái quát hóa.

## III. NĂNG LỰC TOÁN HỌC GẮN VỚI HOẠT ĐỘNG
- Tư duy và lập luận toán học: dùng ở hình thành kiến thức và rèn luyện.
- Mô hình hóa toán học: dùng khi chuyển tình huống thành biểu thức/hình vẽ.
- Giải quyết vấn đề toán học: dùng trong chuỗi bài tập từ cơ bản đến nâng cao.
- Giao tiếp toán học: dùng khi HS trình bày và phản biện lời giải.
- Sử dụng công cụ và phương tiện học toán: dùng máy tính, thước, hình vẽ hoặc phần mềm khi phù hợp.

## IV. CHUẨN BỊ
- Giáo viên: kế hoạch dạy học, câu hỏi định hướng, phiếu bài tập phân hóa.
- Học sinh: vở ghi, dụng cụ học tập, kiến thức nền liên quan.

## V. TIẾN TRÌNH DẠY HỌC
{chr(10).join(lesson_sections)}

## VI. KIỂM TRA NHANH TIÊU CHÍ
- Đủ 4 bước TDS.
- Có giờ thực.
- Có phân hóa HS yếu/đại trà/giỏi.
- Có 5 năng lực Toán gắn hoạt động.
- Có tối thiểu 3 ví dụ rèn luyện.
- Công thức dùng LaTeX."""


def generate_lesson_text(plan: TDSWeekPlan | MoetWeekPlan, program: str = "TDS") -> str:
    settings = load_settings()
    require_values(settings, ["anthropic_api_key"])
    client = Anthropic(api_key=settings.anthropic_api_key)
    try:
        response = client.messages.create(
            model=settings.claude_model,
            max_tokens=10000,
            system=LESSON_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": build_lesson_prompt(plan, program)}],
        )
    except Exception as exc:
        return generate_fallback_lesson_text(plan, str(exc), program)

    return "\n".join(
        block.text for block in response.content if getattr(block, "type", "") == "text"
    )


def clean_markdown_inline(text: str) -> str:
    cleaned = text.replace("<br/>", "\n").replace("<br />", "\n").replace("<br>", "\n")
    cleaned = re.sub(r"(?<!\\)\*\*(.+?)(?<!\\)\*\*", r"\1", cleaned)
    cleaned = re.sub(r"(?<!\\)__(.+?)(?<!\\)__", r"\1", cleaned)
    cleaned = cleaned.replace("`", "")
    return cleaned.strip()


def set_run_font(paragraph, size: int = 12, bold: bool = False) -> None:
    for run in paragraph.runs:
        run.font.name = "Times New Roman"
        run.font.size = Pt(size)
        run.bold = bool(run.bold or bold)
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")


def add_multiline_paragraph(document: Document, text: str) -> None:
    render_markdown_to_docx(document, text)


def set_cell_text(cell, text: str, *, bold: bool = False, font_size: int = 12, alignment: WD_ALIGN_PARAGRAPH | None = None) -> None:
    cell.text = ""
    lines = clean_markdown_inline(text).splitlines() or [""]
    first_paragraph = cell.paragraphs[0]
    add_math_aware_paragraph(first_paragraph, lines[0])
    for line in lines[1:]:
        paragraph = cell.add_paragraph()
        add_math_aware_paragraph(paragraph, line)
    for paragraph in cell.paragraphs:
        if alignment is not None:
            paragraph.alignment = alignment
        set_run_font(paragraph, font_size, bold)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def set_table_borders(table) -> None:
    table_element = table._tbl
    table_properties = table_element.tblPr
    borders = table_properties.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        table_properties.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "8")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "000000")


def apply_table_grid_style(table) -> None:
    try:
        table.style = "Table Grid"
    except KeyError:
        pass


def set_cell_width(cell, width_twips: int) -> None:
    cell_properties = cell._tc.get_or_add_tcPr()
    width = cell_properties.first_child_found_in("w:tcW")
    if width is None:
        width = OxmlElement("w:tcW")
        cell_properties.append(width)
    width.set(qn("w:w"), str(width_twips))
    width.set(qn("w:type"), "dxa")


def shade_cell(cell, fill: str = "E2E8F0") -> None:
    cell_properties = cell._tc.get_or_add_tcPr()
    shading = cell_properties.first_child_found_in("w:shd")
    if shading is None:
        shading = OxmlElement("w:shd")
        cell_properties.append(shading)
    shading.set(qn("w:fill"), fill)


def set_activity_table_widths(table, col_count: int) -> None:
    if col_count != 3:
        return
    widths = [1700, 4200, 4100]
    for row in table.rows:
        for cell, width in zip(row.cells, widths, strict=False):
            set_cell_width(cell, width)


def add_table_row(table, values: list[str], *, bold: bool = False) -> None:
    row = table.add_row()
    for index, value in enumerate(values):
        set_cell_text(row.cells[index], value, bold=bold, alignment=WD_ALIGN_PARAGRAPH.CENTER if bold else None)


def lesson_title(plan: TDSWeekPlan | MoetWeekPlan) -> str:
    if not plan.lessons:
        return "Kế hoạch dạy học theo PPCT"
    first = plan.lessons[0].content.strip()
    if len(plan.lessons) == 1:
        return first
    return f"{first} và các nội dung tuần {plan.week}"


def ppct_summary(plan: TDSWeekPlan | MoetWeekPlan) -> str:
    return "\n".join(f"Tiết {lesson.period}: {lesson.content}" for lesson in plan.lessons)


def apply_base_style(document: Document) -> None:
    styles = document.styles
    styles["Normal"].font.name = "Times New Roman"
    styles["Normal"].font.size = Pt(12)
    styles["Normal"]._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    for style_name in ("Heading 1", "Heading 2", "Heading 3"):
        if style_name in styles:
            styles[style_name].font.name = "Times New Roman"
            styles[style_name]._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
            styles[style_name].font.bold = True
            styles[style_name].font.size = Pt(13 if style_name != "Heading 1" else 14)


def add_section_table(document: Document, title: str, instruction: str, body: str) -> None:
    table = document.add_table(rows=2, cols=1)
    apply_table_grid_style(table)
    set_table_borders(table)
    set_cell_text(table.rows[0].cells[0], f"{title}\n{instruction}".strip(), bold=True, font_size=12)
    shade_cell(table.rows[0].cells[0])
    set_cell_text(table.rows[1].cells[0], body, font_size=12)


def add_activity_table(document: Document, title: str, instruction: str, rows: list[list[str]]) -> None:
    header = document.add_table(rows=1, cols=1)
    apply_table_grid_style(header)
    set_table_borders(header)
    set_cell_text(header.rows[0].cells[0], f"{title}\n{instruction}".strip(), bold=True, font_size=12)
    shade_cell(header.rows[0].cells[0], "D9EAD3")

    table = document.add_table(rows=1, cols=3)
    apply_table_grid_style(table)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table)
    for index, value in enumerate(["Thời gian thực", "Giáo viên và Học sinh", "Nội dung"]):
        set_cell_text(table.rows[0].cells[index], value, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)
        shade_cell(table.rows[0].cells[index])
    for row in rows:
        add_table_row(table, row)
    set_activity_table_widths(table, 3)


def is_markdown_table_separator(line: str) -> bool:
    cells = split_markdown_row(line)
    return bool(cells) and all(re.fullmatch(r":?-{3,}:?", cell.strip()) for cell in cells)


def split_markdown_row(line: str) -> list[str]:
    stripped = line.strip()
    if not stripped.startswith("|") or "|" not in stripped[1:]:
        return []
    if stripped.endswith("|"):
        stripped = stripped[:-1]
    stripped = stripped[1:]
    return [cell.strip() for cell in re.split(r"(?<!\\)\|", stripped)]


def add_markdown_table(document: Document, table_rows: list[list[str]]) -> None:
    rows = [row for row in table_rows if not (len(row) > 0 and all(re.fullmatch(r":?-{3,}:?", cell.strip()) for cell in row))]
    if not rows:
        return
    col_count = max(len(row) for row in rows)
    if col_count <= 0:
        return

    table = document.add_table(rows=1, cols=col_count)
    apply_table_grid_style(table)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table)

    for index in range(col_count):
        value = rows[0][index] if index < len(rows[0]) else ""
        set_cell_text(table.rows[0].cells[index], value, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)
        shade_cell(table.rows[0].cells[index])

    for source_row in rows[1:]:
        row = table.add_row()
        for index in range(col_count):
            value = source_row[index] if index < len(source_row) else ""
            set_cell_text(row.cells[index], value)

    set_activity_table_widths(table, col_count)


def add_markdown_paragraph(document: Document, line: str) -> None:
    stripped = line.strip()
    if not stripped:
        return
    if stripped.startswith("#"):
        level = min(max(len(stripped) - len(stripped.lstrip("#")), 1), 3)
        paragraph = document.add_heading("", level=level)
        add_math_aware_paragraph(paragraph, clean_markdown_inline(stripped.lstrip("#").strip()))
        set_run_font(paragraph, 13 if level > 1 else 14, True)
        return
    if stripped.startswith("-"):
        paragraph = document.add_paragraph()
        add_math_aware_paragraph(paragraph, clean_markdown_inline(f"• {stripped[1:].strip()}"))
        set_run_font(paragraph, 12)
        return
    numbered_match = re.match(r"^(\d+\.)\s+", stripped)
    if numbered_match:
        paragraph = document.add_paragraph()
        body = re.sub(r"^\d+\.\s+", "", stripped)
        add_math_aware_paragraph(paragraph, clean_markdown_inline(f"{numbered_match.group(1)} {body}"))
        set_run_font(paragraph, 12)
        return
    if stripped.startswith(">"):
        paragraph = document.add_paragraph()
        add_math_aware_paragraph(paragraph, clean_markdown_inline(stripped.lstrip(">").strip()))
        set_run_font(paragraph, 12, False)
        return

    paragraph = document.add_paragraph()
    add_math_aware_paragraph(paragraph, clean_markdown_inline(stripped))
    set_run_font(paragraph, 12)


def render_markdown_to_docx(document: Document, markdown: str) -> None:
    lines = markdown.splitlines()
    position = 0
    while position < len(lines):
        line = lines[position]
        if split_markdown_row(line):
            table_rows: list[list[str]] = []
            while position < len(lines) and split_markdown_row(lines[position]):
                row = split_markdown_row(lines[position])
                table_rows.append(row)
                position += 1
            add_markdown_table(document, table_rows)
            continue
        add_markdown_paragraph(document, line)
        position += 1


def download_tds_template_if_missing() -> Path | None:
    if TDS_TEMPLATE_DOCX_PATH.exists():
        return TDS_TEMPLATE_DOCX_PATH
    try:
        TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
        client = GoogleDriveClient()
        metadata = client.get_file(TDS_TEMPLATE_FILE_ID)
        mime_type = metadata.get("mimeType", "")
        if mime_type == "application/vnd.google-apps.document":
            return client.export_google_doc(TDS_TEMPLATE_FILE_ID, TDS_TEMPLATE_DOCX_PATH)
        return client.download_file(TDS_TEMPLATE_FILE_ID, TDS_TEMPLATE_DOCX_PATH)
    except Exception as exc:
        print(f"Could not download TDS template, falling back to blank DOCX: {exc}")
        return None


def clear_document_body(document: Document) -> None:
    body = document._body._element
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)


def create_document_from_template() -> Document:
    template_path = download_tds_template_if_missing()
    if template_path and template_path.exists():
        try:
            document = Document(str(template_path))
            clear_document_body(document)
            return document
        except Exception as exc:
            print(f"Could not open TDS template, falling back to blank DOCX: {exc}")
    return Document()


def build_docx(plan: TDSWeekPlan | MoetWeekPlan, lesson_text: str, program: str = "TDS") -> Path:
    GENERATED_DIR.mkdir(parents=True, exist_ok=True)
    output_path = GENERATED_DIR / f"{program}_G{plan.grade}_Tuan_{plan.week:02d}_draft.docx"

    document = create_document_from_template()
    apply_base_style(document)

    notes = getattr(plan, "notes", "")
    title = lesson_title(plan)
    heading = document.add_heading("KẾ HOẠCH DẠY HỌC", level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    info_table = document.add_table(rows=2, cols=6)
    apply_table_grid_style(info_table)
    set_table_borders(info_table)
    rows = [
        ["Lớp", str(plan.grade), "Tên bài học", title, "Môn học", "Toán"],
        ["Giáo viên", "TDS THT", "Tuần học", str(plan.week), "Năm học", "2025 – 2026"],
    ]
    for row, values in zip(info_table.rows, rows, strict=True):
        for index, value in enumerate(values):
            is_label = index % 2 == 0
            set_cell_text(row.cells[index], value, bold=is_label, alignment=WD_ALIGN_PARAGRAPH.CENTER if is_label else None)
            if is_label:
                shade_cell(row.cells[index])

    materials = "\n".join([
        f"Nội dung PPCT {program}:",
        ppct_summary(plan),
    ])
    if notes:
        materials += f"\nGhi chú PPCT: {notes}"
    add_section_table(document, "PPCT TUẦN", "Dữ liệu nguồn dùng để soạn giáo án", materials)

    render_markdown_to_docx(document, lesson_text)

    document.save(output_path)
    return output_path


def get_existing_or_create_week_folder(client: GoogleDriveClient, parent_id: str, week: int) -> dict[str, str]:
    existing_folder = find_week_folder(client, parent_id, week)
    if existing_folder:
        return existing_folder
    return client.get_or_create_child_folder(parent_id, f"Tuần {week}")


def upload_generated_files(files: GeneratedLessonFiles, parent_id: str, week: int) -> dict[str, str]:
    client = GoogleDriveClient()
    week_folder = get_existing_or_create_week_folder(client, parent_id, week)
    links: dict[str, str] = {}

    uploaded_docx = client.upload_file(files.docx_path, week_folder["id"], DOCX_MIME_TYPE)
    links["docx"] = uploaded_docx.get("webViewLink", "")
    print(f"Uploaded DOCX: {uploaded_docx.get('name')} | {uploaded_docx.get('id')} | {links['docx']}")

    if files.pdf_path:
        uploaded_pdf = client.upload_file(files.pdf_path, week_folder["id"], PDF_MIME_TYPE)
        links["pdf"] = uploaded_pdf.get("webViewLink", "")
        print(f"Uploaded PDF: {uploaded_pdf.get('name')} | {uploaded_pdf.get('id')} | {links['pdf']}")

    return links


def generate_tds_docx(grade: int, week: int, track: str, upload: bool = False, notify: bool = False) -> GeneratedLessonFiles:
    plan = extract_tds_week(TDS_EXCEL_PATH, grade, week, track)
    lesson_text = generate_lesson_text(plan, "TDS")
    title = lesson_title(plan)
    if web_render_enabled():
        docx_path = render_docx_with_web(
            title=title,
            content=lesson_text,
            output_dir=GENERATED_DIR,
            filename=f"TDS_G{plan.grade}_Tuan_{plan.week:02d}_web.docx",
        )
    else:
        docx_path = build_docx(plan, lesson_text, "TDS")
    pdf_path = render_pdf(plan, lesson_text, "TDS")
    files = GeneratedLessonFiles(docx_path=docx_path, pdf_path=pdf_path)
    print(f"Generated DOCX: {docx_path}")
    print(f"Generated PDF: {pdf_path}")

    uploaded_links: dict[str, str] = {}
    if upload:
        uploaded_links = upload_generated_files(files, tds_grade_output_folder_id(grade), week)

    if notify:
        message = f"Đã tạo giáo án nháp TDS G{grade} tuần {week:02d}: {docx_path.name} và {pdf_path.name}"
        for label, link in uploaded_links.items():
            if link:
                message += f"\n{label.upper()}: {link}"
        build_notifier().send_message(message)
        print("Telegram notification sent.")

    return files


def generate_moet_docx(grade: int, week: int, upload: bool = False, notify: bool = False) -> GeneratedLessonFiles:
    plan = extract_moet_week(grade, week)
    lesson_text = generate_lesson_text(plan, "MOET")
    title = lesson_title(plan)
    if web_render_enabled():
        docx_path = render_docx_with_web(
            title=title,
            content=lesson_text,
            output_dir=GENERATED_DIR,
            filename=f"MOET_G{plan.grade}_Tuan_{plan.week:02d}_web.docx",
        )
    else:
        docx_path = build_docx(plan, lesson_text, "MOET")
    pdf_path = render_pdf(plan, lesson_text, "MOET")
    files = GeneratedLessonFiles(docx_path=docx_path, pdf_path=pdf_path)
    print(f"Generated DOCX: {docx_path}")
    print(f"Generated PDF: {pdf_path}")

    uploaded_links: dict[str, str] = {}
    if upload:
        uploaded_links = upload_generated_files(files, moet_grade_output_folder_id(grade), week)

    if notify:
        message = f"Đã tạo giáo án nháp MOET G{grade} tuần {week:02d}: {docx_path.name} và {pdf_path.name}"
        for label, link in uploaded_links.items():
            if link:
                message += f"\n{label.upper()}: {link}"
        build_notifier().send_message(message)
        print("Telegram notification sent.")

    return files


def selected_programs(include_tds: bool, include_moet: bool) -> list[str]:
    programs: list[str] = []
    if include_tds:
        programs.append("tds")
    if include_moet:
        programs.append("moet")
    return programs


def generate_missing_docx(
    programs: list[str],
    start_week: int,
    end_week: int,
    upload: bool = True,
    notify: bool = True,
    generate_partial_weeks: bool = False,
) -> None:
    audit_all_range(programs, start_week, end_week, notify)
    missing = collect_missing_weeks(programs, start_week, end_week, generate_partial_weeks)
    if not missing:
        message = "Không phát hiện tuần trống hoàn toàn cần tạo trong phạm vi đã kiểm tra."
        print(message)
        if notify:
            build_notifier().send_message(message)
        return

    summary_lines = ["Bắt đầu tạo giáo án cho các tuần còn thiếu:"]
    for program, grade, week, audit in missing:
        summary_lines.append(
            f"- {PROGRAM_LABELS[program]} G{grade} tuần {week:02d}: còn thiếu ước lượng {audit.missing_count} file"
        )
    summary = "\n".join(summary_lines)
    print(summary)
    if notify:
        build_notifier().send_message(summary)

    for program, grade, week, _audit in missing:
        if program == "tds":
            generate_tds_docx(grade, week, "dgs", upload, notify)
        elif program == "moet":
            generate_moet_docx(grade, week, upload, notify)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate lesson-plan DOCX files")
    parser.add_argument("--tds", action="store_true", help="Generate one TDS lesson-plan DOCX")
    parser.add_argument("--moet", action="store_true", help="Generate one Moet lesson-plan DOCX")
    parser.add_argument("--missing", action="store_true", help="Audit first, then generate only missing weekly lesson-plan DOCX files")
    parser.add_argument("--start-week", type=int, default=1)
    parser.add_argument("--end-week", type=int, default=1)
    parser.add_argument("--grade", type=int, choices=[10, 11, 12], default=10)
    parser.add_argument("--week", type=int, default=1)
    parser.add_argument("--track", choices=["dgs", "discover"], default="dgs")
    parser.add_argument("--partial-weeks", action="store_true", help="Also regenerate weeks that already have some DOCX files but still look incomplete")
    parser.add_argument("--upload", action="store_true", help="Upload generated DOCX to Google Drive")
    parser.add_argument("--notify", action="store_true", help="Send Telegram notification after generation")
    args = parser.parse_args()

    if args.missing:
        programs = selected_programs(args.tds, args.moet) or ["tds", "moet"]
        generate_missing_docx(programs, args.start_week, args.end_week, args.upload, args.notify, args.partial_weeks)
    elif args.tds:
        generate_tds_docx(args.grade, args.week, args.track, args.upload, args.notify)
    elif args.moet:
        generate_moet_docx(args.grade, args.week, args.upload, args.notify)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
