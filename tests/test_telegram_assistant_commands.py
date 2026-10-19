from __future__ import annotations

import sys
import types
import unittest
import requests
from unittest.mock import patch

from app.assistant_client import AssistantIdentity


def _stub_legacy_imports() -> dict[str, types.ModuleType]:
    modules: dict[str, types.ModuleType] = {}
    drive_audit = types.ModuleType("app.drive_audit")
    drive_audit.WeekAudit = object
    drive_audit.audit_week = lambda *args, **kwargs: None
    drive_audit.format_week_audit = lambda *args, **kwargs: ""
    modules["app.drive_audit"] = drive_audit
    drive_client = types.ModuleType("app.drive_client")
    drive_client.GoogleDriveClient = object
    modules["app.drive_client"] = drive_client
    generator = types.ModuleType("app.lesson_generator")
    generator.ExistingOutputConflict = object
    generator.GeneratedLessonBatch = object
    generator.find_existing_output_conflicts = lambda *args, **kwargs: []
    generator.generate_moet_docx = lambda *args, **kwargs: None
    generator.generate_tds_docx = lambda *args, **kwargs: None
    modules["app.lesson_generator"] = generator
    notify = types.ModuleType("app.telegram_notify")
    notify.build_notifier = lambda: None
    modules["app.telegram_notify"] = notify
    return modules


class FakeAssistant:
    def __init__(self) -> None:
        self.calls: list[tuple[str, AssistantIdentity, object, str | None]] = []

    def send_request(self, action: str, identity: AssistantIdentity, payload: object = None, *, request_id: str | None = None) -> dict:
        self.calls.append((action, identity, payload, request_id))
        return {"kind": "answer", "result": {"linked": True} if action == "consumeLinkCode" else {"status": "running"}}

    def get_job_status(self, identity: AssistantIdentity, job_id: str, *, request_id: str | None = None) -> dict:
        self.calls.append(("jobStatus", identity, {"jobId": job_id}, request_id))
        return {"kind": "job", "result": {"status": "running"}}


class TelegramAssistantCommandTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.modules = _stub_legacy_imports()
        cls.patcher = patch.dict(sys.modules, cls.modules)
        cls.patcher.start()
        import app.telegram_bot as telegram_bot
        cls.telegram_bot = telegram_bot

    @classmethod
    def tearDownClass(cls) -> None:
        cls.patcher.stop()

    def test_parse_assistant_commands_preserves_hash_and_arguments(self) -> None:
        parser = self.telegram_bot.parse_assistant_command
        self.assertEqual(parser("/link CODE").args, ("CODE",))
        self.assertEqual(parser("/confirm c-1 HASH-ABC").args, ("c-1", "HASH-ABC"))
        self.assertEqual(parser("/status job-1").action, "status")
        self.assertIsNone(parser("soạn TDS lớp 10 tuần 1"))

    def test_private_link_and_status_route_to_assistant_with_stable_message_id(self) -> None:
        bot = self.telegram_bot.TelegramPollingBot.__new__(self.telegram_bot.TelegramPollingBot)
        bot.assistant_account_id = "telegram-bot-1"
        bot.assistant_client = FakeAssistant()
        bot.pending_requests = {}
        bot.pending_overwrites = {}
        sent: list[tuple[str, str]] = []
        bot.send_message = lambda chat_id, text: sent.append((chat_id, text))
        message = {"chat": {"id": 123, "type": "private"}, "from": {"id": 123}, "message_id": 77, "text": "/link CODE"}

        self.assertTrue(bot._handle_assistant_message(message, "123", "/link CODE"))
        self.assertEqual(bot.assistant_client.calls[0][0], "consumeLinkCode")
        self.assertEqual(bot.assistant_client.calls[0][3], "telegram-link-123-77")
        self.assertIn("Đã liên kết", sent[-1][1])

        message["text"] = "/status job-1"
        message["message_id"] = 78
        self.assertTrue(bot._handle_assistant_message(message, "123", "/status job-1"))
        self.assertEqual(bot.assistant_client.calls[-1][0], "jobStatus")
        self.assertIn("running", sent[-1][1])

    def test_group_chat_does_not_call_assistant(self) -> None:
        bot = self.telegram_bot.TelegramPollingBot.__new__(self.telegram_bot.TelegramPollingBot)
        bot.assistant_account_id = "telegram-bot-1"
        bot.assistant_client = FakeAssistant()
        bot.pending_requests = {}
        bot.pending_overwrites = {}
        sent: list[str] = []
        bot.send_message = lambda _chat_id, text: sent.append(text)
        message = {"chat": {"id": -99, "type": "group"}, "from": {"id": 123}, "message_id": 77, "text": "/status job-1"}

        self.assertTrue(bot._handle_assistant_message(message, "-99", "/status job-1"))
        self.assertEqual(bot.assistant_client.calls, [])
        self.assertIn("chat riêng", sent[-1])

    def test_duplicate_message_id_is_ignored_and_transport_errors_do_not_leak_token(self) -> None:
        bot = self.telegram_bot.TelegramPollingBot.__new__(self.telegram_bot.TelegramPollingBot)
        bot.assistant_account_id = "telegram-bot-1"
        bot.assistant_client = FakeAssistant()
        bot.pending_requests = {}
        bot.pending_overwrites = {}
        bot.seen_assistant_messages = set()
        sent: list[str] = []
        bot.send_message = lambda _chat_id, text: sent.append(text)
        message = {"chat": {"id": 123, "type": "private"}, "from": {"id": 123}, "message_id": 77, "text": "/status job-1"}
        self.assertTrue(bot._handle_assistant_message(message, "123", message["text"]))
        self.assertTrue(bot._handle_assistant_message(message, "123", message["text"]))
        self.assertEqual(len(bot.assistant_client.calls), 1)

        bot.api_base = "https://api.telegram.org/botSECRET/getUpdates"
        with patch.object(self.telegram_bot.requests, "get", side_effect=requests.RequestException("https://api.telegram.org/botSECRET/getUpdates")):
            with self.assertRaises(RuntimeError) as raised:
                bot.get_updates(None)
        self.assertNotIn("SECRET", str(raised.exception))

    def test_assistant_reply_redacts_private_data(self) -> None:
        result = self.telegram_bot._safe_assistant_result({"result": {"answer": "https://storage.example/x?token=secret"}})
        self.assertNotIn("secret", result)
        private = self.telegram_bot._safe_assistant_result({"result": {"answer": "Mã PIN: 1234; nội dung bài nộp của học sinh"}})
        self.assertIn("ẩn", private)


if __name__ == "__main__":
    unittest.main()
