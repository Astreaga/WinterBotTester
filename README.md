# Winter Bot Tester

A PySide6 + Playwright desktop app that QA-tests the Telarus **Expedient**
supplier-recommendation bot (an Open WebUI instance).

## What it does

For each test question, the app:

1. Opens the bot in a real Chromium window (persistent login profile).
2. Types the question, sends it, and waits for the answer.
3. Downloads the bot's Word (`.docx`) answer.
4. Repeats for N rounds per question, across many suppliers.
5. Judges each answer — did the expected supplier surface? — and optionally
   submits a thumbs up/down rating back to the site.
6. Builds a per-supplier combined Word report.

## Requirements

- Python 3.11+
- macOS (Apple Silicon tested)

## Setup

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
playwright install chromium
```

## Run

```bash
python app.py
```

Sign in to Expedient once in the Chromium window that opens; the login profile
persists across runs.

## Notes

- No credentials or API keys are stored in this repository. Anything sensitive
  lives in the app's local settings or your environment.
- Test question templates are `.xlsx` files kept outside the repo.
