from __future__ import annotations

import hmac
import os
import shutil
import tempfile
from pathlib import Path
from typing import Annotated

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from pydantic import BaseModel

from app.config import load_settings
from app.drive_client import GoogleDriveClient

app = FastAPI(title="Lesson Bot API", docs_url=None, redoc_url=None)

# Allow the soangiaoan web app (any origin) to call these endpoints from the browser.
# Auth is handled via X-API-Token header, not cookies, so allow_credentials=False is correct.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["X-API-Token", "Content-Type"],
)

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
PDF_MIME = "application/pdf"
_EXT_MIME: dict[str, str] = {"docx": DOCX_MIME, "pdf": PDF_MIME}


def _root_folder_id(lesson_type: str, grade: int) -> str:
    settings = load_settings()
    key = f"{lesson_type.lower()}_g{grade}_folder_id"
    folder_id: str = getattr(settings, key, "") or ""
    if not folder_id:
        raise HTTPException(
            status_code=400,
            detail=f"Drive folder not configured for {lesson_type.upper()} G{grade}. "
                   f"Set {key.upper()} in Railway env vars.",
        )
    return folder_id


def verify_token(x_api_token: Annotated[str | None, Header()] = None) -> None:
    expected = os.getenv("WEB_API_TOKEN", "").strip()
    if not expected:
        raise HTTPException(status_code=500, detail="WEB_API_TOKEN not configured on server")
    if not hmac.compare_digest(x_api_token or "", expected):
        raise HTTPException(status_code=401, detail="Invalid or missing X-API-Token")


class CheckRequest(BaseModel):
    lesson_type: str
    grade: int
    week: int
    filename: str | None = None


@app.post("/api/lessons/check", dependencies=[Depends(verify_token)])
async def check_lesson(body: CheckRequest) -> JSONResponse:
    """Return whether the week folder and optionally a specific file already exist on Drive."""
    if body.grade not in (10, 11, 12):
        raise HTTPException(status_code=400, detail="grade must be 10, 11, or 12")
    if not 1 <= body.week <= 40:
        raise HTTPException(status_code=400, detail="week must be 1-40")

    root_id = _root_folder_id(body.lesson_type, body.grade)
    client = GoogleDriveClient()
    week_folder = client.find_child_folder(root_id, f"Tuần {body.week:02d}")

    if not week_folder:
        return JSONResponse({
            "folder_exists": False,
            "folder_url": None,
            "files": [],
            "filename_exists": False if body.filename else None,
        })

    folder_id = week_folder["id"]
    folder_url = week_folder.get("webViewLink") or f"https://drive.google.com/drive/folders/{folder_id}"
    raw_files = client.list_files(
        query=(
            f"'{folder_id}' in parents and trashed = false "
            "and mimeType != 'application/vnd.google-apps.folder'"
        ),
        page_size=100,
    )
    files = [
        {"name": f["name"], "id": f["id"], "url": f.get("webViewLink", "")}
        for f in raw_files
    ]
    filename_exists = any(f["name"] == body.filename for f in files) if body.filename else None

    return JSONResponse({
        "folder_exists": True,
        "folder_url": folder_url,
        "files": files,
        "filename_exists": filename_exists,
    })


@app.post("/api/drive/upload", dependencies=[Depends(verify_token)])
async def upload_lesson(
    file: Annotated[UploadFile, File()],
    lesson_type: Annotated[str, Form()],
    grade: Annotated[int, Form()],
    week: Annotated[int, Form()],
    replace_existing: Annotated[bool, Form()] = True,
) -> JSONResponse:
    """Upload a DOCX or PDF lesson file into the correct Drive week sub-folder."""
    if grade not in (10, 11, 12):
        raise HTTPException(status_code=400, detail="grade must be 10, 11, or 12")
    if not 1 <= week <= 40:
        raise HTTPException(status_code=400, detail="week must be 1-40")
    if not file.filename:
        raise HTTPException(status_code=400, detail="filename required")

    ext = Path(file.filename).suffix.lower().lstrip(".")
    mime_type = _EXT_MIME.get(ext) or file.content_type or "application/octet-stream"

    root_id = _root_folder_id(lesson_type, grade)
    client = GoogleDriveClient()
    week_folder = client.get_or_create_child_folder(root_id, f"Tuần {week:02d}")
    folder_id = week_folder["id"]
    folder_url = week_folder.get("webViewLink") or f"https://drive.google.com/drive/folders/{folder_id}"

    content = await file.read()
    tmp_dir = Path(tempfile.mkdtemp())
    try:
        tmp_path = tmp_dir / file.filename
        tmp_path.write_bytes(content)
        result = client.upload_file(tmp_path, folder_id, mime_type=mime_type, replace_existing=replace_existing)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    file_id = result.get("id", "")
    return JSONResponse({
        "success": True,
        "drive_file_id": file_id,
        "drive_url": result.get("webViewLink") or f"https://drive.google.com/file/d/{file_id}/view",
        "folder_url": folder_url,
        "filename": result.get("name"),
    })


@app.get("/health")
async def health() -> JSONResponse:
    return JSONResponse({"status": "ok"})
