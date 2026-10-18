# Lesson Bot

Python automation core for weekly Math lesson-plan generation.

## Current workflow

1. Configure environment variables in `.env` based on `.env.example`.
2. Test Telegram notification.
3. Configure Google Drive OAuth credentials.
4. Download and parse PPCT files from Google Drive.
5. Audit existing lesson-plan folders in Google Drive.
6. Send audit summaries and missing-plan suggestions through Telegram.
7. Generate DOCX drafts for TDS or MOET.
8. Upload generated DOCX files into the correct grade/week folder.

## Useful commands

Run all commands from the repository root.

### Test Telegram

```cmd
cd lesson-bot && set PYTHONIOENCODING=utf-8 && .venv\Scripts\python.exe -m app.telegram_notify
```

### Test Google Drive access

```cmd
cd lesson-bot && set PYTHONIOENCODING=utf-8 && .venv\Scripts\python.exe -m app.drive_client --test
```

### Download PPCT source files

```cmd
cd lesson-bot && set PYTHONIOENCODING=utf-8 && .venv\Scripts\python.exe -m app.ppct_parser --download
```

### Inspect or extract TDS PPCT

```cmd
cd lesson-bot && set PYTHONIOENCODING=utf-8 && .venv\Scripts\python.exe -m app.ppct_parser --extract-tds-week --grade 10 --week 1 --track dgs
```

### Inspect or extract MOET PPCT

```cmd
cd lesson-bot && set PYTHONIOENCODING=utf-8 && .venv\Scripts\python.exe -m app.moet_parser --week 1 --grade 10
```

### Audit existing TDS lesson plans

```cmd
cd lesson-bot && set PYTHONIOENCODING=utf-8 && .venv\Scripts\python.exe -m app.drive_audit --tds --start-week 1 --end-week 1 --notify
```

### Audit existing MOET lesson plans

```cmd
cd lesson-bot && set PYTHONIOENCODING=utf-8 && .venv\Scripts\python.exe -m app.drive_audit --moet --start-week 1 --end-week 1 --notify
```

### Generate one TDS DOCX and upload it

```cmd
cd lesson-bot && set PYTHONIOENCODING=utf-8 && .venv\Scripts\python.exe -m app.lesson_generator --tds --grade 10 --week 1 --track dgs --upload --notify
```

### Generate one MOET DOCX and upload it

```cmd
cd lesson-bot && set PYTHONIOENCODING=utf-8 && .venv\Scripts\python.exe -m app.lesson_generator --moet --grade 10 --week 1 --upload --notify
```

### Audit first, then generate only fully empty missing weekly bundles

By default this avoids regenerating a week folder that already has some DOCX files.

```cmd
cd lesson-bot && set PYTHONIOENCODING=utf-8 && .venv\Scripts\python.exe -m app.lesson_generator --missing --tds --moet --start-week 1 --end-week 1 --upload --notify
```

### Force regeneration for partially filled weeks

Use this only when you intentionally want the bot to generate a weekly draft even though the week folder already contains some DOCX files.

```cmd
cd lesson-bot && set PYTHONIOENCODING=utf-8 && .venv\Scripts\python.exe -m app.lesson_generator --missing --moet --start-week 1 --end-week 1 --partial-weeks --upload --notify
```

## Current behavior notes

- TDS uses the Excel PPCT file and supports grade 10, 11, and 12.
- MOET uses the PDF PPCT files and groups 5 periods per week.
- Existing/missing detection currently estimates completeness by comparing PPCT lesson count with DOCX/Google Docs count in each week folder.
- The safer `--missing` mode skips partially filled weeks unless `--partial-weeks` is explicitly provided.
- If the AI API call fails, the generator creates a structured fallback draft instead of stopping the workflow.

## SmartPlan Assistant gateway client

`app.assistant_client.AssistantClient` is the server-side bridge to the SmartPlan Assistant gateway. It sends JSON requests to `SMARTPLAN_ASSISTANT_URL` and requires these identity fields on every request:

```json
{
  "channel": "telegram",
  "channelAccountId": "<OA or bot account ID>",
  "channelUserId": "<channel-scoped user ID>",
  "channelChatId": "<optional chat ID>"
}
```

The request body also contains `action` and a generated `requestId` (or the caller's explicit id). The client serializes the body as compact, UTF-8 JSON with sorted keys. It signs the exact bytes with HMAC-SHA256 using `SMARTPLAN_ASSISTANT_SECRET` and this signing string:

```text
<unix_timestamp>.<nonce>.<raw_json_body>.smartplan-assistant
```

It sends the hexadecimal signature in `X-Assistant-Signature`, with the timestamp in `X-Assistant-Timestamp` and nonce in `X-Assistant-Nonce`. Supported actions are `prepareConfirmation`, `confirm`, `cancel`, and `jobStatus`:

- `prepareConfirmation`: stable caller-supplied `requestId` plus `toolName`, `args`, `preview`, `resourceIds`, `resourceScopes`, optional `recipients`; retries must reuse the same ID.
- `confirm` / `cancel`: `confirmationId`, `requestHash`.
- `jobStatus`: `jobId`.

`wait_for_job()` polls until `succeeded`, `failed`, or `cancelled`, then raises `JOB_POLL_TIMEOUT` when the poll limit is reached. The default timeout is 20 seconds per request and the default poll interval is 2 seconds; override them with the two numeric environment variables.

## Channel adapters

The bot now exposes `POST /api/zalo/webhook` when `app.main --run` starts. `ZaloAdapter` accepts signed inbound text events only when the event OA ID equals `ZALO_OA_ID` and is present in `ZALO_OA_ALLOWED_IDS`. It verifies `X-ZEVENT-SIGNATURE` (HMAC-SHA256 over the exact raw body), requires a fresh provider event ID/timestamp, claims the event in the SQLite replay store, normalizes the OA ID plus OA-scoped Zalo UID, and replies through `https://openapi.zalo.me/v3.0/oa/message/cs` using the server-only access token.

Supported reactive Zalo commands: `/link CODE`, `/confirm ID HASH`, `/cancel ID HASH`, `/status JOB_ID`, `/help`. Other text is sent to the shared `chat` gateway action. No startup or proactive Zalo message is sent. Errors never include access tokens, webhook secrets, or response bodies.

Telegram keeps the existing TDS/MOET parser. In private chats only, `/link`, `/confirm`, `/cancel`, `/status`, `/unlink`, and non-legacy natural language can route to the Assistant gateway when `TELEGRAM_ACCOUNT_ID` and the SmartPlan Assistant settings are present. `/unlink` points the teacher to SmartPlan because channel unlink requires the authenticated web flow. Group chats are rejected for Assistant routing. Duplicate/edited assistant messages are ignored, transport errors are normalized, and assistant replies redact private URLs, PINs, tokens, and submission markers.

Keep `SMARTPLAN_ASSISTANT_SECRET` only in the Railway/server environment or a local `.env` that is excluded from Git. Do not put it in browser settings, client bundles, Telegram messages, logs, request previews, or committed files. Rotate the secret at the gateway and bot together. Run the contract tests without external calls:

```text
python -m unittest discover -s tests -p "test_*.py" -v
```

## Do not commit secrets

Never commit `.env`, Google credential files, OAuth token files, generated DOCX outputs, or API keys.
