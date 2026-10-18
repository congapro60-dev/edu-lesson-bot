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
    web_word_render_url: str
    ai_provider: str
    openai_compatible_base_url: str
    openai_compatible_api_key: str
    openai_compatible_model: str
    gemini_api_key: str
    gemini_model: str
    deepseek_base_url: str
    deepseek_api_key: str
    deepseek_model: str
    telegram_bot_token: str
    telegram_chat_id: str
    smartplan_assistant_url: str
    smartplan_assistant_secret: str
    smartplan_assistant_timeout_seconds: float
    smartplan_assistant_poll_interval_seconds: float
    telegram_account_id: str
    zalo_oa_id: str
    zalo_oa_allowed_ids: str
    zalo_oa_access_token: str
    zalo_oa_webhook_secret: str
    zalo_openapi_base_url: str
    zalo_webhook_replay_db: Path
    zalo_webhook_max_age_seconds: float
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
    value = os.getenv(name)
    if value is None or not value.strip():
        return default
    return value.strip()


def _getenv_float(name: str, default: float) -> float:
    raw = _getenv(name)
    if not raw:
        return default
    try:
        value = float(raw)
    except ValueError as exc:
        raise RuntimeError(f"Invalid numeric configuration: {name}") from exc
    if value <= 0:
        raise RuntimeError(f"Configuration must be positive: {name}")
    return value


def load_settings() -> Settings:
    return Settings(
        anthropic_api_key=_getenv("ANTHROPIC_API_KEY"),
        claude_model=_getenv("CLAUDE_MODEL", "claude-sonnet-4-5"),
        web_word_render_url=_getenv("WEB_WORD_RENDER_URL"),
        ai_provider=_getenv("AI_PROVIDER", "anthropic").lower(),
        openai_compatible_base_url=_getenv("OPENAI_COMPATIBLE_BASE_URL"),
        openai_compatible_api_key=_getenv("OPENAI_COMPATIBLE_API_KEY"),
        openai_compatible_model=_getenv("OPENAI_COMPATIBLE_MODEL"),
        gemini_api_key=_getenv("GEMINI_API_KEY"),
        gemini_model=_getenv("GEMINI_MODEL", "gemini-3.5-flash"),
        deepseek_base_url=_getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com"),
        deepseek_api_key=_getenv("DEEPSEEK_API_KEY"),
        deepseek_model=_getenv("DEEPSEEK_MODEL", "deepseek-chat"),
        telegram_bot_token=_getenv("TELEGRAM_BOT_TOKEN"),
        telegram_chat_id=_getenv("TELEGRAM_CHAT_ID"),
        smartplan_assistant_url=_getenv("SMARTPLAN_ASSISTANT_URL"),
        smartplan_assistant_secret=_getenv("SMARTPLAN_ASSISTANT_SECRET"),
        smartplan_assistant_timeout_seconds=_getenv_float("SMARTPLAN_ASSISTANT_TIMEOUT_SECONDS", 20.0),
        smartplan_assistant_poll_interval_seconds=_getenv_float("SMARTPLAN_ASSISTANT_POLL_INTERVAL_SECONDS", 2.0),
        telegram_account_id=_getenv("TELEGRAM_ACCOUNT_ID"),
        zalo_oa_id=_getenv("ZALO_OA_ID"),
        zalo_oa_allowed_ids=_getenv("ZALO_OA_ALLOWED_IDS"),
        zalo_oa_access_token=_getenv("ZALO_OA_ACCESS_TOKEN"),
        zalo_oa_webhook_secret=_getenv("ZALO_OA_WEBHOOK_SECRET", _getenv("ZALO_OA_SECRET_KEY")),
        zalo_openapi_base_url=_getenv("ZALO_OPENAPI_BASE_URL", "https://openapi.zalo.me"),
        zalo_webhook_replay_db=BASE_DIR / _getenv("ZALO_WEBHOOK_REPLAY_DB", "outputs/zalo_webhook_replay.sqlite3"),
        zalo_webhook_max_age_seconds=_getenv_float("ZALO_WEBHOOK_MAX_AGE_SECONDS", 300.0),
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
