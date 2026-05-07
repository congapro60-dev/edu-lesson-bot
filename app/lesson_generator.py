from __future__ import annotations

import argparse
from pathlib import Path

from app.drive_audit import PROGRAM_LABELS, audit_all_range, collect_missing_weeks, find_week_folder

from anthropic import Anthropic
from docx import Document
from docx.shared import Pt

from app.config import BASE_DIR, load_settings, require_values
from app.drive_client import GoogleDriveClient
from app.math_docx import add_math_aware_paragraph
from app.moet_parser import MoetWeekPlan, extract_moet_week
from app.ppct_parser import TDS_EXCEL_PATH, TDSWeekPlan, extract_tds_week
from app.telegram_notify import build_notifier


GENERATED_DIR = BASE_DIR / "outputs" / "generated"


LESSON_SYSTEM_PROMPT = """Bạn là chuyên gia xây dựng kế hoạch dạy học môn Toán THPT theo mẫu của Ban Toán TDS.
Hãy viết bằng tiếng Việt, giọng chuyên môn, rõ ràng, có cấu trúc giống giáo án mẫu.
Mỗi giáo án phải bám nội dung PPCT, tuân thủ quy trình: Trải nghiệm - Hình thành kiến thức/kĩ năng - Rèn luyện, phát triển - Sơ kết.
Ưu tiên câu hỏi định hướng, nhiệm vụ học sinh, sản phẩm học tập, phương án đánh giá và phân hóa học sinh.
Không bịa tên tài liệu hoặc đường dẫn. Nếu thiếu dữ liệu, hãy viết phần phù hợp ở mức khung triển khai có thể chỉnh sửa.
Khi có công thức Toán, bắt buộc viết bằng LaTeX chuẩn để hệ thống chuyển thành Word Equation:
- Công thức trong dòng đặt trong \\( ... \\), ví dụ \\(x^2-2x+1=0\\).
- Công thức riêng dòng đặt trong \\[ ... \\], ví dụ \\[\\Delta=b^2-4ac\\].
- Dùng lệnh LaTeX phổ biến như \\frac{a}{b}, \\sqrt{x}, x^{2}, a_{n}, \\sin x, \\cos x, \\vec{u}.
- Không viết công thức bằng ảnh, không dùng ký hiệu Unicode rời rạc nếu có thể viết bằng LaTeX."""


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
    return f"""Hãy soạn một bản nháp kế hoạch dạy học tuần cho môn Toán {program}.

Thông tin:
- Khối: {plan.grade}
- Tuần: {week_label}
- Tháng: {month}
- Hệ: {track}
- Ghi chú tuần: {notes or "Không có"}

Nội dung PPCT cần soạn:
{lessons_text}

Yêu cầu cấu trúc đầu ra theo form mẫu:
1. KẾ HOẠCH DẠY HỌC
2. Thông tin chung
   - Môn học/Hoạt động giáo dục
   - Lớp/Khối
   - Tên bài/chủ đề
   - Thời lượng
   - Thời điểm thực hiện
3. Mục tiêu học tập
   - Mục tiêu trọng tâm
   - Mục tiêu phân hóa cho học sinh cần hỗ trợ và học sinh khá giỏi
4. Năng lực toán học
   - Năng lực tư duy và lập luận toán học
   - Năng lực mô hình hóa toán học nếu phù hợp
   - Năng lực giải quyết vấn đề toán học
   - Năng lực giao tiếp toán học
   - Năng lực sử dụng công cụ và phương tiện học toán nếu phù hợp
5. Thiết bị dạy học và học liệu
6. Tiến trình dạy học theo từng tiết, mỗi tiết gồm bảng có các cột:
   - Thời gian dự kiến
   - Hoạt động của giáo viên
   - Hoạt động của học sinh
   - Sản phẩm học tập
   - Đánh giá/Phương án hỗ trợ
7. Mỗi tiết phải có đủ 4 bước:
   - Trải nghiệm/khởi động ngắn, ưu tiên dẫn tới hình thành kiến thức
   - Hình thành kiến thức/kĩ năng mới bằng câu hỏi định hướng, hạn chế thuyết trình áp đặt
   - Rèn luyện, phát triển với tối thiểu 3 nhiệm vụ/bài tập từ cơ bản đến vận dụng
   - Sơ kết, chốt lỗi thường gặp và giao nhiệm vụ
8. Check list giáo án cuối bài
   - Đủ mục tiêu trọng tâm và mục tiêu phân hóa
   - Đủ 4 bước hoạt động
   - Có câu hỏi định hướng/phân rã
   - Có tối thiểu 3 ý rèn luyện/củng cố
   - Có năng lực toán học phù hợp
"""


def generate_fallback_lesson_text(plan: TDSWeekPlan | MoetWeekPlan, reason: str, program: str = "TDS") -> str:
    week_label = getattr(plan, "week_label", f"Tuần {plan.week}")
    track = getattr(plan, "track", program)
    lesson_sections = []
    for lesson in plan.lessons:
        lesson_sections.append(
            f"""### Tiết {lesson.period}: {lesson.content}

| Thời gian dự kiến | Hoạt động của giáo viên | Hoạt động của học sinh | Sản phẩm học tập | Đánh giá/Phương án hỗ trợ |
|---|---|---|---|---|
| 3-5 phút | Tổ chức hoạt động trải nghiệm/khởi động ngắn gắn với nội dung {lesson.content}; nêu câu hỏi gợi mở để dẫn vào mục tiêu tiết học. | Suy nghĩ cá nhân, trao đổi nhanh và nêu dự đoán/cách hiểu ban đầu. | Câu trả lời ban đầu, dự đoán hoặc ví dụ minh họa của học sinh. | Quan sát mức độ tham gia; đặt câu hỏi phụ cho học sinh cần hỗ trợ. |
| 15-20 phút | Dẫn dắt hình thành kiến thức/kĩ năng bằng câu hỏi định hướng và câu hỏi phân rã; hạn chế thuyết trình áp đặt. | Thảo luận, trả lời câu hỏi, ghi nhận kiến thức/kĩ năng mới. | Kết luận kiến thức/kĩ năng trọng tâm của tiết học. | Kiểm tra nhanh qua câu hỏi làm rõ, câu hỏi liên quan hoặc phản ví dụ. |
| 15-18 phút | Giao tối thiểu 3 nhiệm vụ rèn luyện từ nhận biết, thông hiểu đến vận dụng; chọn ví dụ đặc trưng, tránh tính toán cồng kềnh. | Làm bài cá nhân/nhóm, trình bày và phản biện cách làm. | Lời giải hoặc sản phẩm luyện tập cho các nhiệm vụ trọng tâm. | Hỗ trợ theo nhóm năng lực; giao thêm câu hỏi mở rộng cho học sinh khá giỏi. |
| 3-5 phút | Chốt kiến thức, nêu lỗi thường gặp và giao nhiệm vụ chuẩn bị tiết sau. | Tự đánh giá mức độ đạt mục tiêu và ghi nhiệm vụ học tập. | Phiếu/tóm tắt cuối tiết hoặc nhiệm vụ về nhà. | Thu tín hiệu phản hồi nhanh để điều chỉnh tiết sau. |"""
        )

    return f"""1. Thông tin chung
- Đây là bản nháp được tạo theo PPCT {program}. Lý do chưa gọi được API sinh nội dung chi tiết: {reason}
- Môn học/Hoạt động giáo dục: Toán
- Khối: {plan.grade}
- Tuần: {week_label}
- Hệ: {track}
- Thời lượng: {len(plan.lessons)} tiết

2. Mục tiêu học tập
- Mục tiêu trọng tâm: Học sinh đạt được các yêu cầu cần đạt theo PPCT tuần, nắm được kiến thức/kĩ năng toán học trọng tâm và vận dụng vào bài tập nhận biết, thông hiểu, vận dụng.
- Mục tiêu phân hóa: Học sinh cần hỗ trợ hoàn thành nhiệm vụ cơ bản với câu hỏi gợi ý; học sinh khá giỏi được giao câu hỏi mở rộng, phản ví dụ hoặc bài toán vận dụng thực tế.

3. Năng lực toán học
- Năng lực tư duy và lập luận toán học: phân tích giả thiết, nhận diện quan hệ toán học, lập luận để giải quyết nhiệm vụ.
- Năng lực mô hình hóa toán học: sử dụng khi nội dung bài học có tình huống thực tiễn hoặc biểu diễn vấn đề bằng mô hình toán học.
- Năng lực giải quyết vấn đề toán học: lựa chọn phương pháp, triển khai lời giải và kiểm tra kết quả.
- Năng lực giao tiếp toán học: trình bày, trao đổi, phản biện lời giải trong hoạt động cá nhân/nhóm.
- Năng lực sử dụng công cụ và phương tiện học toán: dùng thước, máy tính, phần mềm hoặc hình vẽ khi phù hợp.

4. Thiết bị dạy học và học liệu
- Giáo viên: kế hoạch dạy học, bảng phụ/slide, phiếu học tập, câu hỏi định hướng, hệ thống bài tập phân hóa.
- Học sinh: sách/vở, dụng cụ học tập, kiến thức nền liên quan và nhiệm vụ chuẩn bị trước giờ học.

5. Tiến trình dạy học
{chr(10).join(lesson_sections)}

6. Lưu ý phân hóa và hỗ trợ học sinh
- Với học sinh cần hỗ trợ: giảm độ phức tạp tính toán, tăng câu hỏi gợi ý từng bước.
- Với học sinh khá giỏi: bổ sung câu hỏi mở rộng, phản ví dụ hoặc bài toán vận dụng thực tế.

7. Check list giáo án
- Có mục tiêu trọng tâm và mục tiêu phân hóa.
- Có đủ quy trình Trải nghiệm - Hình thành - Rèn luyện/phát triển - Sơ kết.
- Có câu hỏi định hướng và câu hỏi phân rã ở hoạt động hình thành kiến thức/kĩ năng.
- Có tối thiểu 3 ý rèn luyện/củng cố cho nội dung trọng tâm.
- Có năng lực toán học phù hợp.
- Có lưu ý phân hóa học sinh."""


def generate_lesson_text(plan: TDSWeekPlan | MoetWeekPlan, program: str = "TDS") -> str:
    settings = load_settings()
    require_values(settings, ["anthropic_api_key"])
    client = Anthropic(api_key=settings.anthropic_api_key)
    try:
        response = client.messages.create(
            model=settings.claude_model,
            max_tokens=6000,
            system=LESSON_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": build_lesson_prompt(plan, program)}],
        )
    except Exception as exc:
        return generate_fallback_lesson_text(plan, str(exc), program)

    return "\n".join(
        block.text for block in response.content if getattr(block, "type", "") == "text"
    )


def add_multiline_paragraph(document: Document, text: str) -> None:
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("#"):
            paragraph = document.add_heading(stripped.lstrip("#").strip(), level=2)
        elif stripped[0:2].isdigit() and stripped[2:3] == ".":
            paragraph = document.add_heading("", level=2)
            add_math_aware_paragraph(paragraph, stripped)
        elif stripped.startswith("-"):
            paragraph = document.add_paragraph(style="List Bullet")
            add_math_aware_paragraph(paragraph, stripped[1:].strip())
        else:
            paragraph = document.add_paragraph()
            add_math_aware_paragraph(paragraph, stripped)


def set_cell_text(cell, text: str) -> None:
    cell.text = ""
    lines = text.splitlines() or [""]
    first_paragraph = cell.paragraphs[0]
    add_math_aware_paragraph(first_paragraph, lines[0])
    for line in lines[1:]:
        paragraph = cell.add_paragraph()
        add_math_aware_paragraph(paragraph, line)
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(11)


def add_table_row(table, values: list[str]) -> None:
    row = table.add_row()
    for index, value in enumerate(values):
        set_cell_text(row.cells[index], value)


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


def add_section_table(document: Document, title: str, instruction: str, body: str) -> None:
    table = document.add_table(rows=2, cols=1)
    table.style = "Table Grid"
    set_cell_text(table.rows[0].cells[0], f"{title}\n{instruction}".strip())
    set_cell_text(table.rows[1].cells[0], body)


def add_activity_table(document: Document, title: str, instruction: str, rows: list[list[str]]) -> None:
    header = document.add_table(rows=1, cols=1)
    header.style = "Table Grid"
    set_cell_text(header.rows[0].cells[0], f"{title}\n{instruction}".strip())

    table = document.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    for index, value in enumerate(["Thời gian thực", "Giáo viên và Học sinh", "Nội dung"]):
        set_cell_text(table.rows[0].cells[index], value)
    for row in rows:
        add_table_row(table, row)


def build_docx(plan: TDSWeekPlan | MoetWeekPlan, lesson_text: str, program: str = "TDS") -> Path:
    GENERATED_DIR.mkdir(parents=True, exist_ok=True)
    output_path = GENERATED_DIR / f"{program}_G{plan.grade}_Tuan_{plan.week:02d}_draft.docx"

    document = Document()
    apply_base_style(document)

    notes = getattr(plan, "notes", "")
    title = lesson_title(plan)
    document.add_heading("KẾ HOẠCH DẠY HỌC", level=1)

    info_table = document.add_table(rows=2, cols=6)
    info_table.style = "Table Grid"
    rows = [
        ["Lớp", str(plan.grade), "Tên bài học", title, "Môn học", "Toán"],
        ["Giáo viên", "TDS THT", "Tuần học", str(plan.week), "Năm học", "2025 – 2026"],
    ]
    for row, values in zip(info_table.rows, rows, strict=True):
        for index, value in enumerate(values):
            set_cell_text(row.cells[index], value)

    section = document.add_table(rows=1, cols=1)
    section.style = "Table Grid"
    set_cell_text(section.rows[0].cells[0], "I. THÔNG TIN CHUNG")

    competencies = "\n".join([
        "- Năng lực tư duy và lập luận toán học: phân tích giả thiết, nhận diện quan hệ toán học, lập luận và kiểm chứng kết quả.",
        "- Năng lực mô hình hóa toán học: chuyển tình huống hoặc bài toán sang biểu diễn toán học khi phù hợp.",
        "- Năng lực giải quyết vấn đề toán học: lựa chọn chiến lược, triển khai lời giải, đánh giá tính hợp lí.",
        "- Năng lực giao tiếp toán học: trình bày, trao đổi, phản biện lời giải bằng ngôn ngữ toán học.",
        "- Năng lực sử dụng công cụ và phương tiện học toán: sử dụng máy tính, hình vẽ, bảng phụ hoặc phần mềm khi cần.",
    ])
    add_section_table(document, "1. Tiêu chuẩn năng lực cốt lõi (1c)", "Liệt kê các tiêu chuẩn năng lực cốt lõi cho mỗi tiết dạy", competencies)

    objectives = "\n".join([
        "Mục tiêu tối thiểu (Đa số học sinh đạt được):",
        f"- Tôi có thể trình bày được kiến thức/kĩ năng trọng tâm của tuần {plan.week}: {title}.",
        "- Tôi có thể thực hiện các nhiệm vụ nhận biết, thông hiểu và vận dụng trực tiếp theo PPCT.",
        "Mục tiêu phân hóa (Học sinh khá, giỏi):",
        "- Tôi có thể giải thích, so sánh cách làm và mở rộng bài toán trong tình huống mới.",
    ])
    add_section_table(document, "2. Mục tiêu học tập (1c)", "Liệt kê các mục tiêu học tập theo hình thức tuyên bố 'Tôi có thể...' cho mỗi tiết dạy.", objectives)

    materials = "\n".join([
        "- Kế hoạch dạy học, sách giáo khoa/sách bài tập theo chương trình hiện hành.",
        "- Slide hoặc bảng phụ tóm tắt câu hỏi định hướng và nhiệm vụ học tập.",
        "- Phiếu bài tập phân hóa; máy tính cầm tay, thước, phần mềm/hình vẽ nếu phù hợp.",
        f"- Nội dung PPCT {program}:\n{ppct_summary(plan)}",
    ])
    if notes:
        materials += f"\n- Ghi chú PPCT: {notes}"
    add_section_table(document, "3. Tài liệu dạy học (1d)", "Liệt kê tất cả các tài liệu cần thiết cho tiết dạy, bao gồm: sách, PPT, PBT, video,...", materials)

    progress = document.add_table(rows=1, cols=1)
    progress.style = "Table Grid"
    set_cell_text(progress.rows[0].cells[0], "II. TIẾN TRÌNH HOẠT ĐỘNG (1a, 1b, 1e, 1f)")

    add_activity_table(document, "1. ÔN CÁI ĐÃ BIẾT/KHỞI ĐỘNG", "Lựa chọn hoạt động mở đầu ngắn gọn, tạo nhu cầu học tập và kết nối kiến thức nền.", [["3 phút", "GV: Nêu tình huống/câu hỏi gợi mở liên quan bài học.\nHS: Suy nghĩ cá nhân và chia sẻ nhanh.", f"Khởi động bằng ví dụ hoặc câu hỏi dẫn tới nội dung: {title}."], ["2 phút", "GV: Chốt vấn đề cần khám phá.\nHS: Ghi nhận nhiệm vụ học tập.", "Hình thành nhu cầu học tập và xác định hướng tiếp cận bài học."]])
    add_activity_table(document, "2. XÁC ĐỊNH MỤC TIÊU HỌC TẬP", "Mục tiêu học tập được đặt ra rõ ràng, học sinh hiểu mình cần đạt được điều gì sau bài học.", [["3 phút", "GV: Công bố mục tiêu theo ngôn ngữ 'Tôi có thể...'.\nHS: Đối chiếu với hiểu biết ban đầu.", objectives]])

    main_rows = [["35-40 phút", "GV: Tổ chức câu hỏi định hướng, nhiệm vụ cá nhân/nhóm, hỗ trợ phân hóa và chốt kiến thức.\nHS: Thực hiện nhiệm vụ, trình bày, phản biện, tự sửa lỗi.", f"Tiết {lesson.period}: {lesson.content}\n- Hình thành kiến thức/kĩ năng trọng tâm.\n- Luyện tập tối thiểu 3 nhiệm vụ từ cơ bản đến vận dụng.\n- Sơ kết, nêu lỗi thường gặp và giao nhiệm vụ tiếp nối."] for lesson in plan.lessons]
    add_activity_table(document, "3. CÁC HOẠT ĐỘNG HỌC TẬP CHÍNH", "Tổ chức chuỗi hoạt động hình thành kiến thức, luyện tập, vận dụng và đánh giá thường xuyên theo PPCT.", main_rows)

    document.add_heading("PHẦN SINH NỘI DUNG CHI TIẾT TỪ AI", level=2)
    add_multiline_paragraph(document, lesson_text)

    document.save(output_path)
    return output_path


def get_existing_or_create_week_folder(client: GoogleDriveClient, parent_id: str, week: int) -> dict[str, str]:
    existing_folder = find_week_folder(client, parent_id, week)
    if existing_folder:
        return existing_folder
    return client.get_or_create_child_folder(parent_id, f"Tuần {week}")


def generate_tds_docx(grade: int, week: int, track: str, upload: bool = False, notify: bool = False) -> Path:
    plan = extract_tds_week(TDS_EXCEL_PATH, grade, week, track)
    lesson_text = generate_lesson_text(plan, "TDS")
    output_path = build_docx(plan, lesson_text, "TDS")
    print(f"Generated DOCX: {output_path}")

    uploaded_link = ""
    if upload:
        parent_id = tds_grade_output_folder_id(grade)
        client = GoogleDriveClient()
        week_folder = get_existing_or_create_week_folder(client, parent_id, week)
        uploaded = client.upload_file(
            output_path,
            week_folder["id"],
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        uploaded_link = uploaded.get("webViewLink", "")
        print(f"Uploaded DOCX: {uploaded.get('name')} | {uploaded.get('id')} | {uploaded_link}")

    if notify:
        message = f"Đã tạo giáo án nháp TDS G{grade} tuần {week:02d}: {output_path.name}"
        if uploaded_link:
            message += f"\n{uploaded_link}"
        build_notifier().send_message(message)
        print("Telegram notification sent.")

    return output_path


def generate_moet_docx(grade: int, week: int, upload: bool = False, notify: bool = False) -> Path:
    plan = extract_moet_week(grade, week)
    lesson_text = generate_lesson_text(plan, "MOET")
    output_path = build_docx(plan, lesson_text, "MOET")
    print(f"Generated DOCX: {output_path}")

    uploaded_link = ""
    if upload:
        parent_id = moet_grade_output_folder_id(grade)
        client = GoogleDriveClient()
        week_folder = get_existing_or_create_week_folder(client, parent_id, week)
        uploaded = client.upload_file(
            output_path,
            week_folder["id"],
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        uploaded_link = uploaded.get("webViewLink", "")
        print(f"Uploaded DOCX: {uploaded.get('name')} | {uploaded.get('id')} | {uploaded_link}")

    if notify:
        message = f"Đã tạo giáo án nháp MOET G{grade} tuần {week:02d}: {output_path.name}"
        if uploaded_link:
            message += f"\n{uploaded_link}"
        build_notifier().send_message(message)
        print("Telegram notification sent.")

    return output_path


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
