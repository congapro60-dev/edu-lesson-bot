from __future__ import annotations

import argparse
import io
from pathlib import Path
from typing import Any

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

from app.config import load_settings


SCOPES = ["https://www.googleapis.com/auth/drive"]


class GoogleDriveClient:
    def __init__(self) -> None:
        self.settings = load_settings()
        self.service = build("drive", "v3", credentials=self._load_credentials())

    def _load_credentials(self) -> Credentials:
        token_file = self.settings.google_token_file
        credentials_file = self.settings.google_credentials_file
        creds: Credentials | None = None

        if token_file.exists():
            creds = Credentials.from_authorized_user_file(str(token_file), SCOPES)

        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())

        if not creds or not creds.valid:
            if not credentials_file.exists():
                raise FileNotFoundError(
                    f"Google credentials file not found: {credentials_file}"
                )
            flow = InstalledAppFlow.from_client_secrets_file(str(credentials_file), SCOPES)
            creds = flow.run_local_server(port=0)
            token_file.parent.mkdir(parents=True, exist_ok=True)
            token_file.write_text(creds.to_json(), encoding="utf-8")

        return creds

    def list_files(self, query: str | None = None, page_size: int = 10) -> list[dict[str, Any]]:
        params: dict[str, Any] = {
            "pageSize": page_size,
            "fields": "files(id, name, mimeType, parents, webViewLink, shortcutDetails)",
            "supportsAllDrives": True,
            "includeItemsFromAllDrives": True,
        }
        if query:
            params["q"] = query

        result = self.service.files().list(**params).execute()
        return result.get("files", [])

    def get_file(self, file_id: str) -> dict[str, Any]:
        return (
            self.service.files()
            .get(
                fileId=file_id,
                fields="id, name, mimeType, parents, webViewLink, shortcutDetails",
                supportsAllDrives=True,
            )
            .execute()
        )

    def export_google_doc(self, file_id: str, output_path: Path) -> Path:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        request = self.service.files().export_media(
            fileId=file_id,
            mimeType="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        with output_path.open("wb") as file_handle:
            downloader = MediaIoBaseDownload(file_handle, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
        return output_path

    def create_folder(self, name: str, parent_id: str) -> dict[str, Any]:
        metadata = {
            "name": name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_id],
        }
        return (
            self.service.files()
            .create(
                body=metadata,
                fields="id, name, webViewLink",
                supportsAllDrives=True,
            )
            .execute()
        )

    def find_child_folder(self, parent_id: str, name: str) -> dict[str, Any] | None:
        escaped_name = name.replace("'", "\\'")
        query = (
            f"'{parent_id}' in parents and "
            "mimeType = 'application/vnd.google-apps.folder' and "
            f"name = '{escaped_name}' and trashed = false"
        )
        files = self.list_files(query=query, page_size=10)
        return files[0] if files else None

    def get_or_create_child_folder(self, parent_id: str, name: str) -> dict[str, Any]:
        existing = self.find_child_folder(parent_id, name)
        if existing:
            return existing
        return self.create_folder(name=name, parent_id=parent_id)

    def download_file(self, file_id: str, output_path: Path) -> Path:
        request = self.service.files().get_media(fileId=file_id, supportsAllDrives=True)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with output_path.open("wb") as file_handle:
            downloader = MediaIoBaseDownload(file_handle, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
        return output_path

    def upload_file(self, local_path: Path, parent_id: str, mime_type: str | None = None) -> dict[str, Any]:
        media = MediaFileUpload(str(local_path), mimetype=mime_type, resumable=True)
        metadata = {"name": local_path.name, "parents": [parent_id]}
        return (
            self.service.files()
            .create(
                body=metadata,
                media_body=media,
                fields="id, name, webViewLink, parents",
                supportsAllDrives=True,
            )
            .execute()
        )


def test_drive_connection() -> None:
    client = GoogleDriveClient()
    files = client.list_files(page_size=10)
    print("Google Drive connection successful.")
    print("Recent accessible files/folders:")
    for file in files:
        line = f"- {file.get('name')} | {file.get('id')} | {file.get('mimeType')}"
        print(line.encode("utf-8", errors="replace").decode("utf-8"))


def list_folder(folder_id: str) -> None:
    client = GoogleDriveClient()
    files = client.list_files(query=f"'{folder_id}' in parents and trashed = false", page_size=100)
    print(f"Found {len(files)} items in folder {folder_id}:")
    for file in files:
        shortcut = file.get("shortcutDetails") or {}
        target_id = shortcut.get("targetId", "")
        target_mime_type = shortcut.get("targetMimeType", "")
        shortcut_suffix = f" -> {target_id} | {target_mime_type}" if target_id else ""
        line = f"- {file.get('name')} | {file.get('id')} | {file.get('mimeType')}{shortcut_suffix}"
        print(line.encode("utf-8", errors="replace").decode("utf-8"))


def main() -> None:
    parser = argparse.ArgumentParser(description="Google Drive helper")
    parser.add_argument("--test", action="store_true", help="Run OAuth flow and list accessible files")
    parser.add_argument("--list-folder", help="List direct children of a Google Drive folder ID")
    args = parser.parse_args()

    if args.test:
        test_drive_connection()
    elif args.list_folder:
        list_folder(args.list_folder)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
