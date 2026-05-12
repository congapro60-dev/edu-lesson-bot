# HANDOFF — edu-lesson-bot

**Cập nhật**: 2026-05-12
**Mục đích**: Session mới đọc file này trước khi làm bất kỳ thứ gì.

---

## 1. Tổng quan

**App**: Telegram bot + FastAPI server chạy trên Railway
**Stack**: Python 3.11 + python-telegram-bot + FastAPI + uvicorn + Google Drive API + Anthropic Claude
**Mục đích**: (1) Bot Telegram soạn giáo án; (2) FastAPI làm cầu nối để web soangiaoan đẩy file lên Google Drive

---

## 2. Những gì đã làm xong

### Session 2026-05-12 (session hiện tại)

| Thay đổi | File | Mô tả |
|------------|------|-------|
| Rút gọn tên file giáo án | `app/lesson_generator.py` | `lesson_filename_prefix()` chỉ lấy tên bài (trước " - HS/GV..."), bỏ yêu cầu cần đạt + "(tiết X)" |
| Force web pipeline | `app/lesson_generator.py` | `render_single_lesson_files()` raise RuntimeError nếu `WEB_WORD_RENDER_URL` chưa set, bỏ fallback local renderer |

### Session 2026-05-11

| Thay đổi | File | Mô tả |
|------------|------|-------|
| Đồng bộ prompt bot với web | `app/lesson_generator.py` | Thay prompt TDS 4 bước bằng CLAUDE_FORMAT giống web: WALT/WILF + Danielson + 5 HĐ + 3 mức 🌶️ |
| max_tokens tăng | `app/lesson_generator.py` | 10000 → 16000 |
| XML extraction | `app/lesson_generator.py` | Parse `<lesson_content>` tag, fallback bỏ `<thinking>` |
| FastAPI server | `app/bot_api_server.py` | `POST /api/lessons/check` + `POST /api/drive/upload` + `/health` |
| CORS | `app/bot_api_server.py` | `allow_origins=["*"]`, `Authorization` + `X-API-Token` |
| Auth dual support | `app/bot_api_server.py` | Hỗ trợ cả `Authorization: Bearer` (mới) lẫn `X-API-Token` (legacy) |
| Default no-replace | `app/bot_api_server.py` | `replace_existing` mặc định `False` — không ghi đè tự động |
| 409 Conflict | `app/bot_api_server.py` | Trả 409 nếu file đã tồn tại và `replace_existing=False` |
| Sanitize filename | `app/bot_api_server.py` | `Path(file.filename).name` — tránh path traversal |
| Optional metadata | `app/bot_api_server.py` | Thêm `period`, `subject` fields (optional) |
| find_week_folder | `app/drive_client.py` | Xử lý cả 'Tuần 1' lẫn 'Tuần 01' — không tạo trùng thư mục |
| upload_to_week_folder | `app/drive_client.py` | Helper: tìm/tạo folder tuần + check conflict + upload |
| requirements.txt | `requirements.txt` | `fastapi`, `uvicorn[standard]`, `python-multipart`, `anthropic` |
| main.py | `app/main.py` | Start FastAPI (thread) + Telegram bot (main thread) song song |
| Force web pipeline | `app/lesson_generator.py` | `render_single_lesson_files()` BẮT BUỘC `WEB_WORD_RENDER_URL`. Bỏ fallback local renderer (gây lỗi DOCX bảng 1 ký tự + PDF LaTeX raw). Raise lỗi rõ nếu env var thiếu. |

---

## 3. Kiến trúc hiện tại

```
Railway service: edu-lesson-bot
  app/main.py  (entrypoint)
    ├─ FastAPI server  (uvicorn, thread)  ─ lắng nghe POST từ web
    └─ TelegramPollingBot  (main thread)  ─ nhận tin nhắn Telegram

FastAPI endpoints:
  POST /api/lessons/check   ─ kiểm tra thư mục tuần + file tồn tại
  POST /api/drive/upload    ─ upload DOCX/PDF lên Drive (trả 409 nếu trùng)
  GET  /health              ─ Railway health probe

Auth:
  - Authorization: Bearer <WEB_API_TOKEN>  (preferred)
  - X-API-Token: <WEB_API_TOKEN>           (legacy, vẫn hỗ trợ)

Drive structure:
  Root folder
    └─ Tuần 01/  (tự tạo nếu chưa có)
         ├─ giao_an_toan_10_tuan01.docx
         └─ giao_an_toan_10_tuan01.pdf
```

---

## 4. Biến môi trường Railway cần thiết

| Biến | Mô tả |
|------|-------|
| `WEB_API_TOKEN` | Token xác thực từ web — sinh bằng `python -c "import secrets; print(secrets.token_hex(32))"` |
| `TDS_G10_FOLDER_ID` | ID thư mục Drive giáo án TDS Lớp 10 |
| `TDS_G11_FOLDER_ID` | ID thư mục Drive giáo án TDS Lớp 11 |
| `TDS_G12_FOLDER_ID` | ID thư mục Drive giáo án TDS Lớp 12 |
| `MOET_G10_FOLDER_ID` | (tuỳ chọn) MOET Lớp 10 |
| `MOET_G11_FOLDER_ID` | (tuỳ chọn) MOET Lớp 11 |
| `MOET_G12_FOLDER_ID` | (tuỳ chọn) MOET Lớp 12 |
| `GOOGLE_TOKEN_JSON` | Nội dung file `token.json` (Google OAuth) |
| `GOOGLE_CREDENTIALS_JSON` | Nội dung file `credentials.json` (Google OAuth) |
| `ANTHROPIC_API_KEY` | API key Claude |
| `TELEGRAM_BOT_TOKEN` | Token Telegram bot |

---

## 5. Flow đẩy giáo án từ web lên Drive

```
Web (React)
  1. POST /api/export-lesson  →  nhận DOCX/PDF base64
  2. POST bot/api/lessons/check  →  kiểm tra tuần có gì không
  3. POST bot/api/drive/upload  →  upload file
     - Nếu trùng tên và replace_existing=false → 409 conflict
     - Web hiện hộp thoại xác nhận
     - User bấm "Ghi đè" → gửi lại với replace_existing=true
```

---

## 6. Biến môi trường bổ sung (proxy Claude)

| Biến | Giá trị | Mô tả |
|------|---------|-------|
| `ANTHROPIC_BASE_URL` | `https://digishop-api.io.vn` | Proxy — KHÔNG có `/v1` ở cuối |
| `CLAUDE_MODEL` | `claude-sonnet-4-6` | Model name proxy nhận (không dùng shorthand) |
| `WEB_WORD_RENDER_URL` | `https://<vercel>.vercel.app/api/render-word` | BẮT BUỘC — bot raise lỗi nếu thiếu |

## 7. Lỗi thường gặp

| Lỗi | Nguyên nhân | Xử lý |
|------|------------|-------|
| `401 Invalid or missing auth token` | Sai `WEB_API_TOKEN` | Kiểm tra token trong Railway và Settings web |
| `400 Drive folder not configured` | Thiếu `TDS_G1X_FOLDER_ID` | Thêm biến môi trường |
| `409 '...' đã tồn tại` | File trùng tên | Bấm "Ghi đè" trong modal web |
| `500 Google credentials` | Thiếu OAuth token | Cấu hình `GOOGLE_TOKEN_JSON` + `GOOGLE_CREDENTIALS_JSON` |
| `401 invalid x-api-key` | Sai key hoặc thiếu `ANTHROPIC_BASE_URL` | Set `ANTHROPIC_BASE_URL=https://digishop-api.io.vn` |
| `404 /v1/v1/messages` | `ANTHROPIC_BASE_URL` có trailing `/v1` | Bỏ `/v1` khỏi base URL |
| `403 model not accessible` | Dùng shorthand model | Set `CLAUDE_MODEL=claude-sonnet-4-6` |
| `RuntimeError: WEB_WORD_RENDER_URL chưa được cấu hình` | Thiếu env var | Set `WEB_WORD_RENDER_URL` trên Railway |

---

## 7. Quy tắc workflow

1. KHÔNG push thẳng main — luôn dùng feature branch + PR
2. Branch convention: `claude/...`
3. Cập nhật HANDOFF.md sau mỗi session
4. Hỏi trước khi tạo file mới
