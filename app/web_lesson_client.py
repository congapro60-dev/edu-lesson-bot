from __future__ import annotations

import re
from pathlib import Path
from urllib.parse import unquote

import requests

from app.config import load_settings


DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
DEFAULT_WORD_RENDER_ENDPOINT = "http://localhost:3000/api/render-word"
FILENAME_RE = re.compile(r"filename\*=UTF-8''([^;]+)|filename=\"?([^\";]+)\"?", re.IGNORECASE)


class WebLessonRenderError(RuntimeError):
    pass


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
