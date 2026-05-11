from __future__ import annotations

import html
import re
from pathlib import Path

from playwright.sync_api import sync_playwright

from app.config import BASE_DIR
from app.moet_parser import MoetWeekPlan
from app.ppct_parser import TDSWeekPlan


GENERATED_DIR = BASE_DIR / "outputs" / "generated"


def _escape_text_keep_math(text: str) -> str:
    parts: list[str] = []
    pattern = re.compile(r"(\\\(.+?\\\)|\\\[.+?\\\]|\$\$.+?\$\$|\$(?!\$).+?(?<!\\)\$)", re.DOTALL)
    position = 0
    for match in pattern.finditer(text):
        parts.append(html.escape(text[position:match.start()]))
        parts.append(match.group(0))
        position = match.end()
    parts.append(html.escape(text[position:]))
    return "".join(parts)


def _paragraph_html(line: str) -> str:
    stripped = line.strip()
    if not stripped:
        return ""
    body = _escape_text_keep_math(stripped)
    if stripped.startswith("###"):
        return f"<h3>{_escape_text_keep_math(stripped.lstrip('#').strip())}</h3>"
    if stripped.startswith("##"):
        return f"<h2>{_escape_text_keep_math(stripped.lstrip('#').strip())}</h2>"
    if stripped.startswith("#"):
        return f"<h2>{_escape_text_keep_math(stripped.lstrip('#').strip())}</h2>"
    if stripped.startswith("-"):
        return f"<p class='bullet'>• {body[1:].strip()}</p>"
    if stripped[:2].isdigit() and stripped[2:3] == ".":
        return f"<h2>{body}</h2>"
    return f"<p>{body}</p>"


def _lesson_title(plan: TDSWeekPlan | MoetWeekPlan) -> str:
    if not plan.lessons:
        return "Kế hoạch dạy học theo PPCT"
    first = plan.lessons[0].content.strip()
    if len(plan.lessons) == 1:
        return first
    return f"{first} và các nội dung tuần {plan.week}"


def _ppct_rows(plan: TDSWeekPlan | MoetWeekPlan) -> str:
    rows = []
    for lesson in plan.lessons:
        rows.append(
            "<tr>"
            f"<td>{html.escape(str(lesson.period))}</td>"
            f"<td>{html.escape(lesson.subject)}</td>"
            f"<td>{_escape_text_keep_math(lesson.content)}</td>"
            "</tr>"
        )
    return "\n".join(rows)


def build_lesson_html(plan: TDSWeekPlan | MoetWeekPlan, lesson_text: str, program: str) -> str:
    title = _lesson_title(plan)
    week_label = getattr(plan, "week_label", f"Tuần {plan.week}")
    notes = getattr(plan, "notes", "")
    detail_html = "\n".join(_paragraph_html(line) for line in lesson_text.splitlines())
    return f"""<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <title>{html.escape(program)} G{plan.grade} Tuần {plan.week}</title>
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css">
  <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.js"></script>
  <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/contrib/auto-render.min.js"></script>
  <script>
    document.addEventListener("DOMContentLoaded", function() {{
      renderMathInElement(document.body, {{
        delimiters: [
          {{left: "$$", right: "$$", display: true}},
          {{left: "\\\\[", right: "\\\\]", display: true}},
          {{left: "\\\\(", right: "\\\\)", display: false}},
          {{left: "$", right: "$", display: false}}
        ],
        throwOnError: false
      }});
    }});
  </script>
  <style>
    @page {{ size: A4; margin: 18mm 15mm; }}
    body {{ font-family: "Times New Roman", "DejaVu Serif", serif; font-size: 12pt; line-height: 1.35; color: #111; }}
    h1 {{ text-align: center; font-size: 18pt; margin: 0 0 12px; text-transform: uppercase; }}
    h2 {{ font-size: 14pt; margin: 14px 0 7px; }}
    h3 {{ font-size: 13pt; margin: 10px 0 5px; }}
    p {{ margin: 5px 0; text-align: justify; }}
    .bullet {{ margin-left: 18px; }}
    table {{ width: 100%; border-collapse: collapse; margin: 8px 0 12px; }}
    th, td {{ border: 1px solid #333; padding: 6px; vertical-align: top; }}
    th {{ background: #f2f2f2; text-align: center; }}
    .meta td {{ text-align: center; }}
    .katex {{ font-size: 1.05em; }}
  </style>
</head>
<body>
  <h1>KẾ HOẠCH DẠY HỌC</h1>
  <table class="meta">
    <tr><td><b>Lớp</b></td><td>{plan.grade}</td><td><b>Tên bài học</b></td><td>{html.escape(title)}</td><td><b>Môn học</b></td><td>Toán</td></tr>
    <tr><td><b>Giáo viên</b></td><td>TDS THT</td><td><b>Tuần học</b></td><td>{html.escape(str(week_label))}</td><td><b>Năm học</b></td><td>2025 – 2026</td></tr>
  </table>
  <h2>I. THÔNG TIN CHUNG</h2>
  <h3>Nội dung PPCT {html.escape(program)}</h3>
  <table>
    <tr><th>Tiết</th><th>Môn/chủ đề</th><th>Nội dung</th></tr>
    {_ppct_rows(plan)}
  </table>
  {f"<p><b>Ghi chú PPCT:</b> {html.escape(notes)}</p>" if notes else ""}
  <h2>II. TIẾN TRÌNH VÀ NỘI DUNG CHI TIẾT</h2>
  {detail_html}
</body>
</html>"""


def render_pdf(
    plan: TDSWeekPlan | MoetWeekPlan,
    lesson_text: str,
    program: str,
    filename_prefix: str | None = None,
) -> Path:
    GENERATED_DIR.mkdir(parents=True, exist_ok=True)
    output_stem = filename_prefix or f"{program}_G{plan.grade}_Tuan_{plan.week:02d}"
    output_path = GENERATED_DIR / f"{output_stem}.pdf"
    html_content = build_lesson_html(plan, lesson_text, program)
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_content(html_content, wait_until="networkidle")
        page.pdf(path=str(output_path), format="A4", print_background=True)
        browser.close()
    return output_path
