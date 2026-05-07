from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv


BASE_DIR = Path(__file__).resolve().parents[1]
load_dotenv(BASE_DIR / ".env")


@dataclass(frozen=True)
class Settings:
    anthropic_api_key: str
    claude_model: str
    telegram_bot_token: str
    telegram_chat_id: str
    google_credentials_file: Path
    google_token_file: Path
    drive_folder_lich_bao_giang: str
    drive_folder_giao_an_mau: str
    drive_folder_giao_an_2526: str
    drive_folder_ppct_2526: str
    tds_root_folder_id: str
    moet_root_folder_id: str
    tds_g10_folder_id: str
    tds_g11_folder_id: str
    tds_g12_folder_id: str
    moet_g10_folder_id: str
    moet_g11_folder_id: str
    moet_g12_folder_id: str
    ppct_tds_excel_file_id: str
    ppct_moet_g10_pdf_file_id: str
    ppct_moet_g11_pdf_file_id: str
    ppct_moet_g12_pdf_file_id: str
    default_week: str
    default_year: str


def _getenv(name: str, default: str = "") -> str:
    return os.getenv(name, default).strip()


def load_settings() -> Settings:
    return Settings(
        anthropic_api_key=_getenv("ANTHROPIC_API_KEY"),
        claude_model=_getenv("CLAUDE_MODEL", "claude-sonnet-4-5"),
        telegram_bot_token=_getenv("TELEGRAM_BOT_TOKEN"),
        telegram_chat_id=_getenv("TELEGRAM_CHAT_ID"),
        google_credentials_file=BASE_DIR / _getenv("GOOGLE_CREDENTIALS_FILE", "credentials/google_credentials.json"),
        google_token_file=BASE_DIR / _getenv("GOOGLE_TOKEN_FILE", "credentials/google_token.json"),
        drive_folder_lich_bao_giang=_getenv("DRIVE_FOLDER_LICH_BAO_GIANG"),
        drive_folder_giao_an_mau=_getenv("DRIVE_FOLDER_GIAO_AN_MAU"),
        drive_folder_giao_an_2526=_getenv("DRIVE_FOLDER_GIAO_AN_2526"),
        drive_folder_ppct_2526=_getenv("DRIVE_FOLDER_PPCT_2526"),
        tds_root_folder_id=_getenv("TDS_ROOT_FOLDER_ID"),
        moet_root_folder_id=_getenv("MOET_ROOT_FOLDER_ID"),
        tds_g10_folder_id=_getenv("TDS_G10_FOLDER_ID"),
        tds_g11_folder_id=_getenv("TDS_G11_FOLDER_ID"),
        tds_g12_folder_id=_getenv("TDS_G12_FOLDER_ID"),
        moet_g10_folder_id=_getenv("MOET_G10_FOLDER_ID"),
        moet_g11_folder_id=_getenv("MOET_G11_FOLDER_ID"),
        moet_g12_folder_id=_getenv("MOET_G12_FOLDER_ID"),
        ppct_tds_excel_file_id=_getenv("PPCT_TDS_EXCEL_FILE_ID"),
        ppct_moet_g10_pdf_file_id=_getenv("PPCT_MOET_G10_PDF_FILE_ID"),
        ppct_moet_g11_pdf_file_id=_getenv("PPCT_MOET_G11_PDF_FILE_ID"),
        ppct_moet_g12_pdf_file_id=_getenv("PPCT_MOET_G12_PDF_FILE_ID"),
        default_week=_getenv("DEFAULT_WEEK", "01"),
        default_year=_getenv("DEFAULT_YEAR", "2025"),
    )


def require_values(settings: Settings, names: list[str]) -> None:
    missing: list[str] = []
    for name in names:
        value = getattr(settings, name)
        if isinstance(value, Path):
            if not str(value):
                missing.append(name)
        elif not value:
            missing.append(name)

    if missing:
        joined = ", ".join(missing)
        raise RuntimeError(f"Missing required configuration values: {joined}")
