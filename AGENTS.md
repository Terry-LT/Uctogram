# IncomeTelegramBot

IncomeTelegramBot is a private, single-owner Telegram bot for tracking expected and actual income and expenses in one local Excel workbook.

## Architecture and stack

- `run.py` is the CLI entry point. It passes the optional backend argument to `bot.run()`. The only supported backend is `local`, which is also the default.
- `bot.py` builds the `python-telegram-bot` application and starts long polling. Async handlers manage inline keyboards, input validation, conversation state, replies, and backups. The owner guard runs in handler group `-1` before command and callback handlers.
- `config.py` uses `python-dotenv` to load `TELEGRAM_TOKEN`, `OWNER_ID` (an integer Telegram user ID), and `EXCEL_FILE` (defaults to `finance.xlsx`). `DATA_DIR` is relative to the current working directory.
- `backend.py` resolves the configured workbook under `data/` and saves changes in place. `excel_ops.py` uses `openpyxl` to read and write the workbook, including calculating actual totals from transactions grouped by trimmed category name.
- `requirements.txt` pins dependencies, including `python-telegram-bot`, `openpyxl`, `python-dotenv`, and APScheduler. The Telegram JobQueue handles the Sunday 18:00 UTC backup when available. The existing local virtual environment uses Python 3.15.

## Run and verify

Run these PowerShell commands from the project root.

```powershell
python -m venv venv
.\venv\Scripts\python.exe -m pip install -r requirements.txt
# For first-time setup only
Copy-Item .env.example .env
```

Fill in `.env` with the bot token, the owner's numeric Telegram user ID, and the filename of a compatible workbook placed in `data/`. Set `EXCEL_FILE` explicitly because the example leaves it blank. Keep `.env` and live finance workbooks out of commits. Use a disposable copy of the workbook when testing writes.

```powershell
.\venv\Scripts\python.exe run.py local
```

The same command without `local` selects the default backend. There is no configured test suite or linter. For a syntax check, run `.\venv\Scripts\python.exe -m py_compile run.py config.py backend.py bot.py excel_ops.py`. For a live smoke test from the owner's Telegram account, use `/start` to check workbook access and `/help` to see the registered user workflows. Exercise `/income`, `/expenses`, `/view_income`, and `/view_expense` for reads. Use a disposable workbook to exercise `/addincome`, `/addexpense`, `/record_income`, `/record_expense`, and `/cancel`. `/backup` sends the configured workbook to the current chat.

## Workbook layout

The month keyboard uses exact worksheet names from `excel_ops.INCOME_SHEETS`. The included `IncomeTemplateDATASET.xlsx` has the expected English month names (September through August) and a `Template` sheet. Rename `Template` to `Cil` in the working copy to use the optional `Cil` menu entry.

| Block | Range | Column order |
| --- | --- | --- |
| Income categories | `C5:E40` | Name, expected, actual |
| Expense categories | `G5:I40` | Name, expected, actual |
| Income transactions | `K5:M40` | Category, date, amount |
| Expense transactions | `O5:Q103` | Category, date, amount |

## Coding conventions

- Follow the existing four-space indentation, `snake_case` functions, uppercase constants, and type hints on function boundaries. Keep Telegram presentation and conversation logic in `bot.py`, worksheet logic in `excel_ops.py`, and path and save operations in `backend.py`.
- Keep the owner guard ahead of other handlers and register conversations before ordinary commands. Use distinct callback prefixes with matching handler patterns, acknowledge callback queries, and store temporary selections in `ctx.user_data`. Escape dynamic content with `html.escape()` in HTML replies.
- Preserve the workbook's fixed ranges and keep reads, writes, and total calculations aligned when changing a range. Read displayed values with `data_only=True` where appropriate. Load the workbook normally before writing to preserve formulas, retain number formats, and persist through `backend.save()`.
- Preserve amount input with decimal commas or dots and optional spaces. Normalize manually entered transaction dates to `DD.MM.YYYY` and store them as text. Keep the matching income and expense flows consistent when changing shared behavior.
