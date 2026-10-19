from __future__ import annotations

import hashlib
import hmac
import json
import tempfile
import time
import unittest
from types import SimpleNamespace

from app.assistant_client import AssistantIdentity
from app.zalo_adapter import ZaloAdapter, parse_zalo_command
from app.zalo_client import ZaloInboundEvent, ZaloOAClient, ZaloReplayStore


class FakeAssistant:
    def __init__(self) -> None:
        self.calls: list[tuple[str, AssistantIdentity, dict[str, object] | None, str | None]] = []

    def send_request(self, action: str, identity: AssistantIdentity, payload: dict[str, object] | None = None, *, request_id: str | None = None) -> dict[str, object]:
        self.calls.append((action, identity, payload, request_id))
        if action == "consumeLinkCode":
            return {"kind": "answer", "result": {"linked": True}}
        return {"kind": "answer", "result": {"message": "ok"}}

    def confirm(self, identity: AssistantIdentity, confirmation_id: str, request_hash: str, *, request_id: str | None = None) -> dict[str, object]:
        self.calls.append(("confirm", identity, {"confirmationId": confirmation_id, "requestHash": request_hash}, request_id))
        return {"kind": "confirmation", "result": {"message": "confirmed"}}

    def cancel(self, identity: AssistantIdentity, confirmation_id: str, request_hash: str, *, request_id: str | None = None) -> dict[str, object]:
        self.calls.append(("cancel", identity, {"confirmationId": confirmation_id, "requestHash": request_hash}, request_id))
        return {"kind": "confirmation", "result": {"message": "cancelled"}}

    def get_job_status(self, identity: AssistantIdentity, job_id: str, *, request_id: str | None = None) -> dict[str, object]:
        self.calls.append(("jobStatus", identity, {"jobId": job_id}, request_id))
        return {"kind": "job", "result": {"status": "running"}}


class FakeReplyClient:
    def __init__(self) -> None:
        self.oa_id = "oa-1"
        self.allowed_oa_ids = frozenset({"oa-1"})
        self.replies: list[tuple[str, str]] = []

    def verify_webhook(self, raw_body: bytes, signature: str | None) -> bool:
        return signature == "valid"

    def send_text(self, user_id: str, text: str) -> dict[str, object]:
        self.replies.append((user_id, text))
        return {"error": 0}


EVENT = ZaloInboundEvent(oa_id="oa-1", user_id="uid-1", text="", event_id="msg-1")


class ZaloAdapterTests(unittest.TestCase):
    def test_command_parser_is_strict_and_keeps_arguments(self) -> None:
        self.assertEqual(parse_zalo_command("/link CODE"), parse_zalo_command("/link CODE"))
        self.assertEqual(parse_zalo_command("/confirm c-1 h-1").args, ("c-1", "h-1"))  # type: ignore[union-attr]
        self.assertIsNone(parse_zalo_command("hello"))

    def test_link_confirm_cancel_status_keep_oa_and_uid_scoped_identity(self) -> None:
        assistant = FakeAssistant()
        adapter = ZaloAdapter(SimpleNamespace(), client=FakeReplyClient(), assistant=assistant)
        for text in ("/link CODE", "/confirm c-1 h-1", "/cancel c-1 h-1", "/status job-1"):
            adapter.handle_event(EVENT.__class__(**{**EVENT.__dict__, "text": text}))

        self.assertEqual([call[0] for call in assistant.calls], ["consumeLinkCode", "confirm", "cancel", "jobStatus"])
        for _, identity, _, request_id in assistant.calls:
            self.assertEqual((identity.channel, identity.channel_account_id, identity.channel_user_id), ("zalo_oa", "oa-1", "uid-1"))
            self.assertTrue(request_id and request_id.startswith("zalo-"))

    def test_webhook_rejects_invalid_signature_before_reply(self) -> None:
        reply = FakeReplyClient()
        adapter = ZaloAdapter(SimpleNamespace(), client=reply, assistant=FakeAssistant())
        raw = json.dumps({"event_name": "user_send_text"}).encode()
        with self.assertRaises(Exception):
            adapter.handle_webhook(raw, "invalid")
        self.assertEqual(reply.replies, [])

    def test_webhook_replies_only_after_verified_inbound_event(self) -> None:
        reply = FakeReplyClient()
        assistant = FakeAssistant()
        adapter = ZaloAdapter(SimpleNamespace(), client=reply, assistant=assistant, replay_store=ZaloReplayStore(tempfile.mktemp()))
        payload = {"event_name": "user_send_text", "timestamp": int(time.time()), "recipient": {"id": "oa-1"}, "sender": {"id": "uid-1"}, "message": {"msg_id": "m-1", "text": "/status job-1"}}
        raw = json.dumps(payload).encode()
        self.assertEqual(adapter.handle_webhook(raw, "valid"), "Trạng thái công việc: running.")
        self.assertEqual(reply.replies, [("uid-1", "Trạng thái công việc: running.")])

    def test_duplicate_or_stale_webhook_is_not_dispatched(self) -> None:
        reply = FakeReplyClient()
        assistant = FakeAssistant()
        adapter = ZaloAdapter(SimpleNamespace(), client=reply, assistant=assistant, replay_store=ZaloReplayStore(tempfile.mktemp()), clock=lambda: 1_700_000_000)
        payload = {"event_name": "user_send_text", "timestamp": 1_700_000_000, "recipient": {"id": "oa-1"}, "sender": {"id": "uid-1"}, "message": {"msg_id": "m-replay", "text": "/status job-1"}}
        raw = json.dumps(payload).encode()
        self.assertEqual(adapter.handle_webhook(raw, "valid"), "Trạng thái công việc: running.")
        self.assertIsNone(adapter.handle_webhook(raw, "valid"))
        self.assertEqual(len(assistant.calls), 1)
        stale = {**payload, "message": {**payload["message"], "msg_id": "m-stale"}, "timestamp": 1_600_000_000}
        with self.assertRaises(Exception):
            adapter.handle_webhook(json.dumps(stale).encode(), "valid")

    def test_webhook_rejects_allowlisted_but_unconfigured_oa(self) -> None:
        reply = FakeReplyClient()
        reply.allowed_oa_ids = frozenset({"oa-1", "oa-2"})
        adapter = ZaloAdapter(SimpleNamespace(), client=reply, assistant=FakeAssistant())
        event = ZaloInboundEvent(oa_id="oa-2", user_id="uid-1", text="/help", event_id="m-2")
        with self.assertRaises(Exception):
            adapter.handle_event(event)


if __name__ == "__main__":
    unittest.main()
