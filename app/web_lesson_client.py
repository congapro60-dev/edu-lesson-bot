from __future__ import annotations

import base64
import re
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import unquote

import requests

from app.config import load_settings


DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
PDF_MIME_TYPE = "application/pdf"
DEFAULT_WORD_RENDER_ENDPOINT = "https://giaoandewey.vercel.app/api/render-word"
DEFAULT_EXPORT_LESSON_ENDPOINT = "https://giaoandewey.vercel.app/api/export-lesson"
FILENAME_RE = re.compile(r"filename\*=UTF-8''([^;]+)|filename=\"?([^\";]+)\"?", re.IGNORECASE)


class WebLessonRenderError(RuntimeError):
    pass


@dataclass(frozen=True)
class WebRenderedLessonFiles:
    docx_path: Path
    pdf_path: Path


def _safe_filename(value: str, fallback: str = "giao-an.docx") -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", value).strip()
    return cleaned or fallback


def _filename_from_disposition(disposition: str, fallback: str) -> str:
    match = FILENAME_RE.search(disposition or "")
    if not match:
        return fallback
    encoded = match.group(1) or match.group(2) or fallback
    return _safe_filename(unquote(encoded), fallback)


def web_render_enabled() -> bool:
    settings = load_settings()
    return bool(settings.web_word_render_url)


def export_lesson_endpoint() -> str:
    settings = load_settings()
    endpoint = (settings.web_word_render_url or DEFAULT_WORD_RENDER_ENDPOINT).rstrip("/")
    if endpoint.endswith("/render-word"):
        return f"{endpoint[: -len('/render-word')]}/export-lesson"
    if endpoint.endswith("/export-lesson"):
        return endpoint
    return DEFAULT_EXPORT_LESSON_ENDPOINT


def render_docx_with_web(
    title: str,
    content: str,
    output_dir: Path,
    orientation: str = "portrait",
    filename: str | None = None,
) -> Path:
    settings = load_settings()
    endpoint = settings.web_word_render_url or DEFAULT_WORD_RENDER_ENDPOINT
    output_dir.mkdir(parents=True, exist_ok=True)

    try:
        response = requests.post(
            endpoint,
            json={"title": title, "content": content, "orientation": orientation},
            timeout=180,
        )
    except requests.RequestException as exc:
        raise WebLessonRenderError(f"Không gọi được web render Word tại {endpoint}: {exc}") from exc

    if response.status_code != 200:
        snippet = response.text[:500]
        raise WebLessonRenderError(f"Web render Word lỗi HTTP {response.status_code}: {snippet}")

    content_type = response.headers.get("content-type", "")
    if DOCX_MIME_TYPE not in content_type and not response.content.startswith(b"PK"):
        raise WebLessonRenderError(f"Web render Word không trả về DOCX hợp lệ: {content_type}")

    fallback_name = filename or f"{_safe_filename(title, 'giao-an')}_A4.docx"
    output_name = filename or _filename_from_disposition(response.headers.get("content-disposition", ""), fallback_name)
    if not output_name.lower().endswith(".docx"):
        output_name += ".docx"

    output_path = output_dir / output_name
    output_path.write_bytes(response.content)
    return output_path


def _write_base64_payload(file_data: dict, output_dir: Path, fallback_name: str, expected_suffix: str) -> Path:
    filename = _safe_filename(fallback_name, fallback_name)
    if not filename.lower().endswith(expected_suffix):
        filename += expected_suffix

    encoded = file_data.get("base64")
    if not encoded or not isinstance(encoded, str):
        raise WebLessonRenderError(f"Web export không trả về base64 cho {filename}")

    try:
        content = base64.b64decode(encoded, validate=True)
    except ValueError as exc:
        raise WebLessonRenderError(f"Web export trả về base64 không hợp lệ cho {filename}") from exc

    output_path = output_dir / filename
    output_path.write_bytes(content)
    return output_path


def render_lesson_files_with_web(
    title: str,
    content: str,
    output_dir: Path,
    grade: int,
    week: int,
    program: str,
    lesson_name: str,
    orientation: str = "portrait",
    filename_prefix: str | None = None,
) -> WebRenderedLessonFiles:
    endpoint = export_lesson_endpoint()
    output_dir.mkdir(parents=True, exist_ok=True)

    try:
        response = requests.post(
            endpoint,
            json={
                "grade": grade,
                "week": week,
                "type": program,
                "lessonName": lesson_name,
                "title": title,
                "content": content,
                "orientation": orientation,
            },
            timeout=240,
        )
    except requests.RequestException as exc:
        raise WebLessonRenderError(f"Không gọi được web export lesson tại {endpoint}: {exc}") from exc

    if response.status_code != 200:
        snippet = response.text[:500]
        raise WebLessonRenderError(f"Web export lesson lỗi HTTP {response.status_code}: {snippet}")

    try:
        payload = response.json()
    except ValueError as exc:
        raise WebLessonRenderError("Web export lesson không trả về JSON hợp lệ") from exc

    prefix = filename_prefix or f"{program}_G{grade}_Tuan_{week:02d}"
    word_data = payload.get("word") or {}
    pdf_data = payload.get("pdf") or {}
    docx_path = _write_base64_payload(word_data, output_dir, f"{prefix}.docx", ".docx")
    pdf_path = _write_base64_payload(pdf_data, output_dir, f"{prefix}.pdf", ".pdf")
    return WebRenderedLessonFiles(docx_path=docx_path, pdf_path=pdf_path)
