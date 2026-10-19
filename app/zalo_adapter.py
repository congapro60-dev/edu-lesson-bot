from __future__ import annotations

import json
import re
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from app.assistant_client import AssistantClient, AssistantClientError, AssistantIdentity, redact_assistant_text
from app.config import Settings
from app.zalo_client import ZaloClientError, ZaloInboundEvent, ZaloOAClient, ZaloReplayStore, normalize_webhook_event


ZALO_HELP = "Dùng /link CODE để liên kết, /confirm ID HASH để xác nhận, /cancel ID HASH để hủy, /status JOB_ID để xem trạng thái."


@dataclass(frozen=True)
class ZaloCommand:
    action: str
    args: tuple[str, ...] = ()


def parse_zalo_command(text: str) -> ZaloCommand | None:
    parts = text.strip().split()
    if not parts or not parts[0].startswith("/"):
        return None
    action = parts[0].lower().split("@", 1)[0]
    if action in {"/help", "/start"} and len(parts) == 1:
        return ZaloCommand("help")
    if action in {"/link", "/confirm", "/cancel", "/status"}:
        return ZaloCommand(action[1:], tuple(parts[1:]))
    return ZaloCommand("unknown")


def _request_id(prefix: str, event_id: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._:-]", "-", event_id).strip("-") or "event"
    return f"zalo-{prefix}-{value}"[:128]


def _safe_result_text(response: dict[str, Any]) -> str:
    result = response.get("result")
    if isinstance(result, dict):
        if result.get("linked") is True:
            return "Đã liên kết tài khoản Zalo OA với SmartPlan."
        for key in ("answer", "text"):
            value = result.get(key)
            if isinstance(value, str) and value.strip():
                return redact_assistant_text(value)
        if isinstance(result.get("confirmationId"), str) and isinstance(result.get("requestHash"), str):
            preview = redact_assistant_text(result.get("preview")) or "Yêu cầu cần xác nhận."
            return f"{preview}\nDùng /confirm {result['confirmationId']} {result['requestHash']} để xác nhận, hoặc /cancel {result['confirmationId']} {result['requestHash']} để hủy."
        if isinstance(result.get("status"), str):
            progress = result.get("progress")
            suffix = f" ({progress}%)" if isinstance(progress, (int, float)) else ""
            return f"Trạng thái công việc: {result['status']}{suffix}."
    return "Đã xử lý yêu cầu trợ lý."


class ZaloAdapter:
    def __init__(self, settings: Settings, *, client: ZaloOAClient | Any | None = None, assistant: AssistantClient | Any | None = None, replay_store: ZaloReplayStore | Any | None = None, clock: Any = time.time) -> None:
        self.settings = settings
        self.client = client or ZaloOAClient(settings)
        self.assistant = assistant or AssistantClient(settings)
        replay_path = getattr(settings, "zalo_webhook_replay_db", Path("outputs/zalo_webhook_replay.sqlite3"))
        self.replay_store = replay_store or ZaloReplayStore(replay_path)
        self.clock = clock
        self.max_age_seconds = float(getattr(settings, "zalo_webhook_max_age_seconds", 300.0) or 300.0)

    def identity(self, event: ZaloInboundEvent) -> AssistantIdentity:
        if not self.client.oa_id or event.oa_id != self.client.oa_id or event.oa_id not in self.client.allowed_oa_ids:
            raise ZaloClientError("Zalo webhook không được phép", code="UNAUTHORIZED_WEBHOOK")
        return AssistantIdentity(channel="zalo_oa", channel_account_id=event.oa_id, channel_user_id=event.user_id)

    def _assistant_reply(self, event: ZaloInboundEvent, response: dict[str, Any]) -> str:
        try:
            return _safe_result_text(response)
        except Exception:
            return "Không xử lý được yêu cầu trợ lý."

    def handle_event(self, event: ZaloInboundEvent) -> str:
        identity = self.identity(event)
        command = parse_zalo_command(event.text)
        request_id = _request_id(command.action if command else "chat", event.event_id)
        try:
            if command and command.action == "help":
                return ZALO_HELP
            if command and command.action == "unknown":
                return ZALO_HELP
            if command and command.action == "link":
                if len(command.args) != 1:
                    return "Cú pháp: /link CODE"
                return self._assistant_reply(event, self.assistant.send_request("consumeLinkCode", identity, {"code": command.args[0]}, request_id=request_id))
            if command and command.action in {"confirm", "cancel"}:
                if len(command.args) != 2:
                    return f"Cú pháp: /{command.action} ID HASH"
                method = self.assistant.confirm if command.action == "confirm" else self.assistant.cancel
                return self._assistant_reply(event, method(identity, command.args[0], command.args[1], request_id=request_id))
            if command and command.action == "status":
                if len(command.args) != 1:
                    return "Cú pháp: /status JOB_ID"
                return self._assistant_reply(event, self.assistant.get_job_status(identity, command.args[0], request_id=request_id))
            response = self.assistant.send_request("chat", identity, {"text": event.text}, request_id=request_id)
            return self._assistant_reply(event, response)
        except AssistantClientError as exc:
            if exc.code in {"ASSISTANT_NOT_CONFIGURED", "ASSISTANT_UNAVAILABLE"}:
                return "Trợ lý hiện chưa sẵn sàng."
            if exc.code in {"UNAUTHORIZED", "INVALID_REQUEST", "CHANNEL_NOT_LINKED", "NOT_LINKED"}:
                return "Không thể xác thực yêu cầu trợ lý. Hãy liên kết lại trong SmartPlan."
            if exc.code == "NOT_IMPLEMENTED":
                return "Tính năng trợ lý này chưa được bật."
            return "Không xử lý được yêu cầu trợ lý."

    def handle_webhook(self, raw_body: bytes, signature: str | None) -> str | None:
        if not self.client.verify_webhook(raw_body, signature):
            raise ZaloClientError("Zalo webhook không được phép", code="UNAUTHORIZED_WEBHOOK")
        event = normalize_webhook_event(raw_body, allowed_oa_ids=self.client.allowed_oa_ids)
        if event is None:
            return None
        if not event.has_provider_event_id or event.timestamp is None:
            raise ZaloClientError("Zalo webhook thiếu định danh/thời gian sự kiện", code="INVALID_WEBHOOK")
        event_seconds = event.timestamp / 1000 if event.timestamp > 100_000_000_000 else event.timestamp
        if abs(float(self.clock()) - event_seconds) > self.max_age_seconds:
            raise ZaloClientError("Zalo webhook đã hết hạn", code="STALE_WEBHOOK")
        if not self.replay_store.claim(f"{event.oa_id}:{event.user_id}:{event.event_id}"):
            return None
        reply = self.handle_event(event)
        self.client.send_text(event.user_id, reply)
        return reply


def register_zalo_webhook(app: Any, adapter: ZaloAdapter) -> None:
    from fastapi import Request
    from fastapi.responses import JSONResponse

    if any(getattr(route, "path", None) == "/api/zalo/webhook" for route in getattr(app, "routes", [])):
        return

    @app.post("/api/zalo/webhook")
    async def zalo_webhook(request: Request) -> JSONResponse:
        raw_body = await request.body()
        signature = (
            request.headers.get("x-zevent-signature")
            or request.headers.get("x-zalo-oa-signature")
            or request.headers.get("x-zalo-signature")
        )
        try:
            adapter.handle_webhook(raw_body, signature)
        except ZaloClientError as exc:
            status = 401 if exc.code == "UNAUTHORIZED_WEBHOOK" else 400
            return JSONResponse({"ok": False, "code": exc.code}, status_code=status)
        return JSONResponse({"ok": True})
