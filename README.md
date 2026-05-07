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

## Do not commit secrets

Never commit `.env`, Google credential files, OAuth token files, generated DOCX outputs, or API keys.
