from __future__ import annotations

import hashlib
import hmac
import json
import sqlite3
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import requests

from app.config import Settings


class ZaloClientError(RuntimeError):
    """Safe transport/configuration error. Never includes token, secret, or response body."""

    def __init__(self, message: str, *, code: str = "ZALO_REQUEST_FAILED", status_code: int | None = None) -> None:
        super().__init__(message)
        self.code = code
        self.status_code = status_code


def parse_allowed_oa_ids(value: str | None) -> frozenset[str]:
    return frozenset(part.strip() for part in (value or "").split(",") if part.strip())


def verify_webhook_signature(raw_body: bytes, signature: str | None, secret_key: str, *, prefix: str = "sha256=") -> bool:
    if not raw_body or not signature or not secret_key:
        return False
    expected = hmac.new(secret_key.encode("utf-8"), raw_body, hashlib.sha256).hexdigest()
    supplied = signature.strip().lower()
    for marker in (prefix, "mac="):
        if marker and supplied.startswith(marker):
            supplied = supplied[len(marker):]
            break
    return bool(len(supplied) == len(expected)) and hmac.compare_digest(supplied, expected)


@dataclass(frozen=True)
class ZaloInboundEvent:
    oa_id: str
    user_id: str
    text: str
    event_id: str
    timestamp: int | None = None
    has_provider_event_id: bool = True


class ZaloReplayStore:
    """Durable one-shot event claim; stores only a hash and expiry."""

    def __init__(self, path: str | Path, *, ttl_seconds: float = 600.0, clock: Any = time.time) -> None:
        self.path = Path(path)
        self.ttl_seconds = ttl_seconds
        self.clock = clock
        if str(self.path) != ":memory:":
            self.path.parent.mkdir(parents=True, exist_ok=True)

    def claim(self, event_key: str) -> bool:
        if not event_key.strip():
            return False
        now = float(self.clock())
        key = hashlib.sha256(event_key.encode("utf-8")).hexdigest()
        with sqlite3.connect(str(self.path), timeout=5) as db:
            db.execute("CREATE TABLE IF NOT EXISTS zalo_webhook_events (event_key TEXT PRIMARY KEY, expires_at REAL NOT NULL)")
            db.execute("DELETE FROM zalo_webhook_events WHERE expires_at <= ?", (now,))
            cursor = db.execute(
                "INSERT OR IGNORE INTO zalo_webhook_events(event_key, expires_at) VALUES (?, ?)",
                (key, now + self.ttl_seconds),
            )
            return cursor.rowcount == 1


def _first_text(*values: object) -> str:
    for value in values:
        if isinstance(value, str) and value.strip():
            return value.strip()
    return ""


def _nested_text(record: dict[str, Any], *keys: str) -> str:
    for key in keys:
        value = record.get(key)
        if isinstance(value, dict):
            found = _first_text(value.get("id"), value.get("user_id"), value.get("userId"), value.get("oa_id"), value.get("oaId"))
            if found:
                return found
        else:
            found = _first_text(value)
            if found:
                return found
    return ""


def normalize_webhook_event(raw_body: bytes | str, *, allowed_oa_ids: frozenset[str]) -> ZaloInboundEvent | None:
    try:
        raw_bytes = raw_body if isinstance(raw_body, bytes) else raw_body.encode("utf-8")
        payload = json.loads(raw_bytes.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError, TypeError) as exc:
        raise ZaloClientError("Zalo webhook không hợp lệ", code="INVALID_WEBHOOK") from exc
    if not isinstance(payload, dict):
        raise ZaloClientError("Zalo webhook không hợp lệ", code="INVALID_WEBHOOK")

    source = payload.get("source") if isinstance(payload.get("source"), dict) else {}
    recipient = payload.get("recipient") if isinstance(payload.get("recipient"), dict) else {}
    message = payload.get("message") if isinstance(payload.get("message"), dict) else {}
    oa_id = _first_text(
        payload.get("oa_id"), payload.get("oaId"), payload.get("official_account_id"),
        source.get("oa_id"), source.get("oaId"), recipient.get("id"), recipient.get("oa_id"), recipient.get("oaId"),
    )
    user_id = _first_text(
        payload.get("user_id_by_app"), payload.get("userIdByApp"), payload.get("user_id"),
        payload.get("uid"), source.get("user_id_by_app"), source.get("user_id"),
        _nested_text(payload, "sender", "user"),
    )
    if not oa_id or not user_id or not allowed_oa_ids or oa_id not in allowed_oa_ids:
        raise ZaloClientError("Zalo webhook không được phép", code="UNAUTHORIZED_WEBHOOK")

    text = _first_text(message.get("text"), message.get("content"), payload.get("text"), payload.get("content"))
    if not text:
        return None

    provider_event_id = _first_text(
        payload.get("event_id"), payload.get("eventId"), payload.get("message_id"), payload.get("messageId"),
        message.get("msg_id"), message.get("message_id"),
    )
    event_id = provider_event_id or hashlib.sha256(raw_bytes).hexdigest()[:32]
    raw_timestamp = payload.get("timestamp") or message.get("timestamp")
    try:
        timestamp = int(raw_timestamp) if raw_timestamp is not None else None
    except (TypeError, ValueError):
        timestamp = None
    return ZaloInboundEvent(oa_id=oa_id, user_id=user_id, text=text, event_id=event_id, timestamp=timestamp, has_provider_event_id=bool(provider_event_id))


class ZaloOAClient:
    def __init__(self, settings: Settings, *, session: requests.Session | Any | None = None) -> None:
        self.oa_id = str(getattr(settings, "zalo_oa_id", "") or "").strip()
        self.allowed_oa_ids = parse_allowed_oa_ids(getattr(settings, "zalo_oa_allowed_ids", ""))
        self.access_token = str(getattr(settings, "zalo_oa_access_token", "") or "").strip()
        self.secret_key = str(getattr(settings, "zalo_oa_webhook_secret", getattr(settings, "zalo_oa_secret_key", "")) or "").strip()
        self.base_url = str(getattr(settings, "zalo_openapi_base_url", "https://openapi.zalo.me") or "").rstrip("/")
        self.session = session or requests.Session()

    def verify_webhook(self, raw_body: bytes, signature: str | None) -> bool:
        if not self.oa_id or self.oa_id not in self.allowed_oa_ids:
            return False
        return verify_webhook_signature(raw_body, signature, self.secret_key)

    def send_text(self, user_id: str, text: str) -> dict[str, Any]:
        if not self.oa_id or self.oa_id not in self.allowed_oa_ids or not self.access_token or not self.base_url:
            raise ZaloClientError("Zalo OA chưa được cấu hình", code="ZALO_NOT_CONFIGURED")
        if not user_id.strip() or not text.strip():
            raise ZaloClientError("Zalo message không hợp lệ", code="INVALID_MESSAGE")
        try:
            response = self.session.post(
                f"{self.base_url}/v3.0/oa/message/cs",
                headers={"access_token": self.access_token, "Content-Type": "application/json"},
                json={"recipient": {"user_id": user_id}, "message": {"text": text}},
                timeout=20,
            )
        except requests.RequestException as exc:
            raise ZaloClientError("Không kết nối được Zalo OpenAPI", code="ZALO_UNAVAILABLE") from exc
        try:
            result = response.json()
        except (ValueError, TypeError) as exc:
            raise ZaloClientError("Zalo OpenAPI trả dữ liệu không hợp lệ", code="INVALID_RESPONSE", status_code=response.status_code) from exc
        if not isinstance(result, dict):
            raise ZaloClientError("Zalo OpenAPI trả dữ liệu không hợp lệ", code="INVALID_RESPONSE", status_code=response.status_code)
        api_error = result.get("error", 0)
        if response.status_code >= 400 or (api_error not in (0, "0", None)):
            raise ZaloClientError("Zalo OpenAPI từ chối yêu cầu", code="ZALO_API_ERROR", status_code=response.status_code)
        return result

