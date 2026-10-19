from __future__ import annotations

import hashlib
import hmac
import json
import unittest
from types import SimpleNamespace
from unittest.mock import call, patch

from app.assistant_client import (
    ASSISTANT_SERVICE_ID,
    AssistantClient,
    AssistantClientError,
    AssistantIdentity,
)


class FakeResponse:
    def __init__(self, payload: object, status_code: int = 200) -> None:
        self.payload = payload
        self.status_code = status_code

    def json(self) -> object:
        return self.payload


class FakeSession:
    def __init__(self, responses: list[FakeResponse]) -> None:
        self.responses = list(responses)
        self.calls: list[dict[str, object]] = []

    def post(self, url: str, *, data: bytes, headers: dict[str, str], timeout: float) -> FakeResponse:
        self.calls.append({"url": url, "data": data, "headers": headers, "timeout": timeout})
        if not self.responses:
            raise AssertionError("Unexpected extra POST")
        return self.responses.pop(0)


def settings(*, url: str = "https://assistant.example/gateway", secret: str = "test-secret", poll_interval: float = 0.5) -> SimpleNamespace:
    return SimpleNamespace(
        smartplan_assistant_url=url,
        smartplan_assistant_secret=secret,
        smartplan_assistant_timeout_seconds=7.5,
        smartplan_assistant_poll_interval_seconds=poll_interval,
    )


IDENTITY = AssistantIdentity(
    channel="telegram",
    channel_account_id="oa-account-1",
    channel_user_id="zalo-user-2",
    channel_chat_id="chat-3",
)


class AssistantClientTests(unittest.TestCase):
    def test_prepare_confirmation_requires_stable_request_id(self) -> None:
        client = AssistantClient(settings(), session=FakeSession([]))
        with self.assertRaises(AssistantClientError) as raised:
            client.prepare_confirmation(
                IDENTITY,
                tool_name="exportLesson",
                args={"lessonPlanId": "lesson-1"},
                preview="Xuất giáo án",
                resource_ids=["lesson-1"],
                resource_scopes=[{"resourceType": "lessonPlan", "resourceId": "lesson-1"}],
                request_id="",
            )
        self.assertEqual(raised.exception.code, "INVALID_REQUEST")

    def test_request_uses_canonical_hmac_envelope(self) -> None:
        session = FakeSession([FakeResponse({"kind": "success", "result": {"accepted": True}})])
        client = AssistantClient(
            settings(),
            session=session,
            clock=lambda: 1_700_000_000.9,
            nonce_factory=lambda: "nonce-1",
        )

        response = client.send_request(
            "lookup",
            IDENTITY,
            {"prompt": "Xin chào", "empty": None},
            request_id="request-1",
        )

        self.assertEqual(response, {"kind": "success", "result": {"accepted": True}})
        self.assertEqual(len(session.calls), 1)
        request = session.calls[0]
        self.assertEqual(request["url"], "https://assistant.example/gateway")
        self.assertEqual(request["timeout"], 7.5)

        raw_body = request["data"].decode("utf-8")  # type: ignore[union-attr]
        body = json.loads(raw_body)
        self.assertEqual(
            body,
            {
                "action": "lookup",
                "channel": "telegram",
                "channelAccountId": "oa-account-1",
                "channelChatId": "chat-3",
                "channelUserId": "zalo-user-2",
                "empty": None,
                "prompt": "Xin chào",
                "requestId": "request-1",
            },
        )
        self.assertEqual(
            raw_body,
            json.dumps(body, ensure_ascii=False, separators=(",", ":"), sort_keys=True),
        )

        headers = request["headers"]
        self.assertEqual(headers["Content-Type"], "application/json")  # type: ignore[index]
        self.assertEqual(headers["X-Assistant-Timestamp"], "1700000000")  # type: ignore[index]
        self.assertEqual(headers["X-Assistant-Nonce"], "nonce-1")  # type: ignore[index]
        signed = "1700000000.nonce-1." + raw_body + "." + ASSISTANT_SERVICE_ID
        expected_signature = hmac.new(b"test-secret", signed.encode("utf-8"), hashlib.sha256).hexdigest()
        self.assertEqual(headers["X-Assistant-Signature"], expected_signature)  # type: ignore[index]

    def test_prepare_confirm_cancel_and_job_status_use_action_contract(self) -> None:
        session = FakeSession([
            FakeResponse({"kind": "success", "result": {"confirmationId": "c-1"}}),
            FakeResponse({"kind": "success", "result": {"confirmed": True}}),
            FakeResponse({"kind": "success", "result": {"cancelled": True}}),
            FakeResponse({"kind": "success", "result": {"status": "running"}}),
        ])
        client = AssistantClient(settings(), session=session, clock=lambda: 1_700_000_001, nonce_factory=lambda: "nonce")

        client.prepare_confirmation(
            IDENTITY,
            tool_name="send_parent_notice",
            args={"classId": "class-1", "message": "Nội dung"},
            preview="Gửi thông báo cho phụ huynh",
            resource_ids=["class-1", "notice-1"],
            resource_scopes=[{"resourceType": "class", "resourceId": "class-1"}],
            recipients=["parent-1"],
            request_id="prepare-1",
        )
        client.confirm(IDENTITY, "c-1", "hash-1", request_id="confirm-1")
        client.cancel(IDENTITY, "c-1", "hash-1", request_id="cancel-1")
        client.get_job_status(IDENTITY, "job-1", request_id="status-1")

        bodies = [json.loads(call_data["data"].decode("utf-8")) for call_data in session.calls]  # type: ignore[union-attr]
        self.assertEqual([body["action"] for body in bodies], ["prepareConfirmation", "confirm", "cancel", "jobStatus"])
        self.assertEqual(bodies[0]["toolName"], "send_parent_notice")
        self.assertEqual(bodies[0]["args"], {"classId": "class-1", "message": "Nội dung"})
        self.assertEqual(bodies[0]["resourceIds"], ["class-1", "notice-1"])
        self.assertEqual(bodies[0]["resourceScopes"], [{"resourceType": "class", "resourceId": "class-1"}])
        self.assertEqual(bodies[0]["recipients"], ["parent-1"])
        self.assertEqual(bodies[1]["confirmationId"], "c-1")
        self.assertEqual(bodies[1]["requestHash"], "hash-1")
        self.assertEqual(bodies[2]["confirmationId"], "c-1")
        self.assertEqual(bodies[3]["jobId"], "job-1")
        self.assertEqual([body["requestId"] for body in bodies], ["prepare-1", "confirm-1", "cancel-1", "status-1"])

    def test_wait_for_job_returns_on_each_terminal_status(self) -> None:
        for status in ("succeeded", "failed", "cancelled"):
            with self.subTest(status=status):
                session = FakeSession([FakeResponse({"kind": "success", "result": {"status": status}})])
                client = AssistantClient(settings(), session=session)
                result = client.wait_for_job(IDENTITY, "job-1", max_polls=1)
                self.assertEqual(result["result"]["status"], status)  # type: ignore[index]
                self.assertEqual(len(session.calls), 1)

    def test_wait_for_job_polls_until_terminal_and_uses_interval(self) -> None:
        session = FakeSession([
            FakeResponse({"kind": "success", "result": {"status": "queued"}}),
            FakeResponse({"kind": "success", "result": {"status": "running"}}),
            FakeResponse({"kind": "success", "result": {"status": "succeeded", "output": {"ok": True}}}),
        ])
        client = AssistantClient(settings(poll_interval=0.25), session=session)

        with patch("app.assistant_client.time.sleep") as sleep:
            result = client.wait_for_job(IDENTITY, "job-1", max_polls=3)

        self.assertEqual(result["result"]["status"], "succeeded")  # type: ignore[index]
        self.assertEqual(sleep.call_args_list, [call(0.25), call(0.25)])
        self.assertEqual(len(session.calls), 3)

    def test_wait_for_job_raises_timeout_after_max_polls(self) -> None:
        session = FakeSession([
            FakeResponse({"kind": "success", "result": {"status": "queued"}}),
            FakeResponse({"kind": "success", "result": {"status": "running"}}),
        ])
        client = AssistantClient(settings(poll_interval=0.25), session=session)

        with patch("app.assistant_client.time.sleep") as sleep:
            with self.assertRaises(AssistantClientError) as raised:
                client.wait_for_job(IDENTITY, "job-1", max_polls=2)

        self.assertEqual(raised.exception.code, "JOB_POLL_TIMEOUT")
        self.assertEqual(sleep.call_args_list, [call(0.25)])
        self.assertEqual(len(session.calls), 2)


if __name__ == "__main__":
    unittest.main()
