from __future__ import annotations

import hashlib
import hmac
import json
import re
import time
import uuid
from dataclasses import dataclass
from typing import Any

import requests

from app.config import Settings


ASSISTANT_SERVICE_ID = "smartplan-assistant"
TERMINAL_JOB_STATUSES = {"succeeded", "failed", "cancelled"}
_PRIVATE_URL = re.compile(r"https?://[^\s<>'\"]*(?:[?&](?:amp;)?(?:token|signature|sig|expires|expiry|x-amz-[^=]+)=|%2F[^\s<>'\"]*\?(?:alt|token)=)[^\s<>'\"]*", re.IGNORECASE)
_SENSITIVE_TEXT = re.compile(r"(?:access\s*token|api\s*key|m[aã]\s*pin|b[aà]i\s*l[aà]m|n[oộ]i\s*dung\s*b[aà]i\s*n[oộ]p|student\s*(?:submission|answer))", re.IGNORECASE)


def redact_assistant_text(value: object, *, max_length: int = 2000) -> str:
    text = value.strip() if isinstance(value, str) else ""
    if not text:
        return ""
    if _SENSITIVE_TEXT.search(text):
        return "Nội dung trợ lý đã được ẩn vì chứa dữ liệu riêng tư."
    redacted = _PRIVATE_URL.sub("[liên kết riêng tư đã ẩn]", text)
    return redacted[:max_length] + ("…" if len(redacted) > max_length else "")


class AssistantClientError(RuntimeError):
    def __init__(self, message: str, *, code: str = "ASSISTANT_REQUEST_FAILED", status_code: int | None = None) -> None:
        super().__init__(message)
        self.code = code
        self.status_code = status_code


@dataclass(frozen=True)
class AssistantIdentity:
    channel: str
    channel_account_id: str
    channel_user_id: str
    channel_chat_id: str | None = None


class AssistantClient:
    def __init__(self, settings: Settings, *, session: requests.Session | Any | None = None, clock: Any = time.time, nonce_factory: Any = uuid.uuid4) -> None:
        self.base_url = settings.smartplan_assistant_url.rstrip("/")
        self.secret = settings.smartplan_assistant_secret
        self.timeout = settings.smartplan_assistant_timeout_seconds
        self.poll_interval = settings.smartplan_assistant_poll_interval_seconds
        self.session = session or requests.Session()
        self.clock = clock
        self.nonce_factory = nonce_factory

    def _validate_config(self) -> None:
        if not self.base_url or not self.secret:
            raise AssistantClientError("SmartPlan Assistant chưa được cấu hình", code="ASSISTANT_NOT_CONFIGURED")

    def _signed_headers(self, raw_body: str) -> dict[str, str]:
        timestamp = str(int(self.clock()))
        nonce = str(self.nonce_factory())
        message = f"{timestamp}.{nonce}.{raw_body}.{ASSISTANT_SERVICE_ID}".encode("utf-8")
        signature = hmac.new(self.secret.encode("utf-8"), message, hashlib.sha256).hexdigest()
        return {
            "Content-Type": "application/json",
            "X-Assistant-Timestamp": timestamp,
            "X-Assistant-Nonce": nonce,
            "X-Assistant-Signature": signature,
        }

    def send_request(self, action: str, identity: AssistantIdentity, payload: dict[str, Any] | None = None, *, request_id: str | None = None) -> dict[str, Any]:
        self._validate_config()
        if not action or not identity.channel or not identity.channel_account_id or not identity.channel_user_id:
            raise AssistantClientError("Yêu cầu Assistant không hợp lệ", code="INVALID_REQUEST")
        body: dict[str, Any] = {
            "action": action,
            "requestId": request_id or str(uuid.uuid4()),
            "channel": identity.channel,
            "channelAccountId": identity.channel_account_id,
            "channelUserId": identity.channel_user_id,
        }
        if identity.channel_chat_id is not None:
            body["channelChatId"] = identity.channel_chat_id
        if payload:
            body.update(payload)
        raw_body = json.dumps(body, ensure_ascii=False, separators=(",", ":"), sort_keys=True)
        try:
            response = self.session.post(self.base_url, data=raw_body.encode("utf-8"), headers=self._signed_headers(raw_body), timeout=self.timeout)
        except requests.RequestException as exc:
            raise AssistantClientError("Không kết nối được SmartPlan Assistant", code="ASSISTANT_UNAVAILABLE") from exc
        try:
            result = response.json()
        except ValueError as exc:
            raise AssistantClientError("SmartPlan Assistant trả dữ liệu không hợp lệ", code="INVALID_RESPONSE", status_code=response.status_code) from exc
        if not isinstance(result, dict):
            raise AssistantClientError("SmartPlan Assistant trả dữ liệu không hợp lệ", code="INVALID_RESPONSE", status_code=response.status_code)
        if response.status_code >= 400 or result.get("kind") == "error":
            code = str(result.get("code") or "ASSISTANT_REQUEST_FAILED")
            raise AssistantClientError(f"Assistant request failed: {code}", code=code, status_code=response.status_code)
        return result

    def prepare_confirmation(self, identity: AssistantIdentity, *, tool_name: str, args: dict[str, Any], preview: str, resource_ids: list[str], resource_scopes: list[dict[str, Any]], request_id: str, recipients: list[str] | None = None) -> dict[str, Any]:
        if not request_id or not request_id.strip():
            raise AssistantClientError("prepareConfirmation cần request_id ổn định", code="INVALID_REQUEST")
        return self.send_request("prepareConfirmation", identity, {
            "toolName": tool_name, "args": args, "preview": preview,
            "resourceIds": resource_ids, "resourceScopes": resource_scopes, "recipients": recipients or [],
        }, request_id=request_id)

    def confirm(self, identity: AssistantIdentity, confirmation_id: str, request_hash: str, *, request_id: str | None = None) -> dict[str, Any]:
        return self.send_request("confirm", identity, {"confirmationId": confirmation_id, "requestHash": request_hash}, request_id=request_id)

    def cancel(self, identity: AssistantIdentity, confirmation_id: str, request_hash: str, *, request_id: str | None = None) -> dict[str, Any]:
        return self.send_request("cancel", identity, {"confirmationId": confirmation_id, "requestHash": request_hash}, request_id=request_id)

    def get_job_status(self, identity: AssistantIdentity, job_id: str, *, request_id: str | None = None) -> dict[str, Any]:
        return self.send_request("jobStatus", identity, {"jobId": job_id}, request_id=request_id)

    def wait_for_job(self, identity: AssistantIdentity, job_id: str, *, max_polls: int = 30) -> dict[str, Any]:
        if max_polls < 1 or max_polls > 300:
            raise ValueError("max_polls must be between 1 and 300")
        for poll in range(max_polls):
            response = self.get_job_status(identity, job_id)
            result = response.get("result") if isinstance(response.get("result"), dict) else {}
            if result.get("status") in TERMINAL_JOB_STATUSES:
                return response
            if poll + 1 < max_polls:
                time.sleep(self.poll_interval)
        raise AssistantClientError("Assistant job chưa hoàn tất trong thời gian chờ", code="JOB_POLL_TIMEOUT")
