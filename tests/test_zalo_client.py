from __future__ import annotations

import hashlib
import hmac
import json
import tempfile
import unittest
from types import SimpleNamespace

from app.zalo_client import ZaloClientError, ZaloOAClient, ZaloReplayStore, normalize_webhook_event, verify_webhook_signature


class FakeResponse:
    def __init__(self, payload: object, status_code: int = 200) -> None:
        self.payload = payload
        self.status_code = status_code

    def json(self) -> object:
        return self.payload


class FakeSession:
    def __init__(self, response: FakeResponse) -> None:
        self.response = response
        self.calls: list[dict[str, object]] = []

    def post(self, url: str, *, headers: dict[str, str], json: dict[str, object], timeout: float) -> FakeResponse:
        self.calls.append({"url": url, "headers": headers, "json": json, "timeout": timeout})
        return self.response


def make_settings(**overrides: object) -> SimpleNamespace:
    values = {
        "zalo_oa_id": "oa-1",
        "zalo_oa_allowed_ids": "oa-1",
        "zalo_oa_access_token": "access-token",
        "zalo_oa_webhook_secret": "webhook-secret",
        "zalo_openapi_base_url": "https://openapi.zalo.me",
    }
    values.update(overrides)
    return SimpleNamespace(**values)


class ZaloClientTests(unittest.TestCase):
    def test_webhook_signature_accepts_exact_body_and_sha256_prefix(self) -> None:
        body = b'{"event_name":"user_send_text"}'
        digest = hmac.new(b"webhook-secret", body, hashlib.sha256).hexdigest()
        self.assertTrue(verify_webhook_signature(body, digest, "webhook-secret"))
        self.assertTrue(verify_webhook_signature(body, f"sha256={digest}", "webhook-secret"))
        self.assertFalse(verify_webhook_signature(body + b"x", digest, "webhook-secret"))
        self.assertFalse(verify_webhook_signature(body, digest, "wrong-secret"))

    def test_normalize_event_enforces_oa_allowlist_and_uid(self) -> None:
        payload = {
            "event_name": "user_send_text",
            "recipient": {"id": "oa-1"},
            "sender": {"id": "uid-1"},
            "message": {"msg_id": "msg-1", "text": "/link CODE"},
        }
        event = normalize_webhook_event(json.dumps(payload).encode(), allowed_oa_ids=frozenset({"oa-1"}))
        self.assertIsNotNone(event)
        assert event is not None
        self.assertEqual((event.oa_id, event.user_id, event.text, event.event_id), ("oa-1", "uid-1", "/link CODE", "msg-1"))
        with self.assertRaises(ZaloClientError) as raised:
            normalize_webhook_event(json.dumps({**payload, "recipient": {"id": "oa-2"}}).encode(), allowed_oa_ids=frozenset({"oa-1"}))
        self.assertEqual(raised.exception.code, "UNAUTHORIZED_WEBHOOK")

    def test_reply_uses_openapi_body_and_does_not_put_token_in_url(self) -> None:
        session = FakeSession(FakeResponse({"error": 0, "message": "Success"}))
        client = ZaloOAClient(make_settings(), session=session)

        result = client.send_text("uid-1", "Xin chào")

        self.assertEqual(result["error"], 0)
        request = session.calls[0]
        self.assertEqual(request["url"], "https://openapi.zalo.me/v3.0/oa/message/cs")
        self.assertNotIn("access-token", request["url"])
        self.assertEqual(request["headers"]["access_token"], "access-token")  # type: ignore[index]
        self.assertEqual(request["json"], {"recipient": {"user_id": "uid-1"}, "message": {"text": "Xin chào"}})

    def test_missing_config_and_http_error_are_safe(self) -> None:
        client = ZaloOAClient(make_settings(zalo_oa_access_token=""), session=FakeSession(FakeResponse({"error": 0})))
        with self.assertRaises(ZaloClientError) as raised:
            client.send_text("uid-1", "hello")
        self.assertEqual(raised.exception.code, "ZALO_NOT_CONFIGURED")
        self.assertNotIn("access-token", str(raised.exception))

        failed = ZaloOAClient(make_settings(), session=FakeSession(FakeResponse({"error": -1, "message": "access-token"}, status_code=401)))
        with self.assertRaises(ZaloClientError) as raised:
            failed.send_text("uid-1", "hello")
        self.assertEqual(raised.exception.code, "ZALO_API_ERROR")
        self.assertNotIn("access-token", str(raised.exception))

    def test_replay_claim_survives_a_new_store_instance(self) -> None:
        path = tempfile.mktemp()
        self.assertTrue(ZaloReplayStore(path, clock=lambda: 100).claim("oa-1:uid-1:event-1"))
        self.assertFalse(ZaloReplayStore(path, clock=lambda: 101).claim("oa-1:uid-1:event-1"))


if __name__ == "__main__":
    unittest.main()
