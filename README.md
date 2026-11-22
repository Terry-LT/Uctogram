# IncomeTelegramBot

Uctogram  is a private Telegram bot for logging income and expenses on the go. It stores transactions in a local Excel workbook and lets you compare expected amounts with actual totals by month.

## Why I built it

I created this project because I was tired of paying monthly subscriptions for commercial finance tracking apps. I wanted a simple way to log my transactions on the go (via Telegram), but I didn't want to pay $5–$15 a month just to do it.

## Setup

1. Create a Telegram bot through BotFather and save its token. You'll also need your numeric Telegram user ID. The bot only accepts commands from the configured owner.
2. From the project root, create a Python virtual environment and install the dependencies. These commands are for PowerShell.

   ```powershell
   python -m venv venv
   .\venv\Scripts\python.exe -m pip install -r requirements.txt
   Copy-Item .env.example .env
   ```

   Copy the example only when creating a new `.env`.

3. Put a compatible `.xlsx` workbook in `data/`. If you use the included `IncomeTemplateDATASET.xlsx` as a starting point, copy it into `data/` and rename its Czech month tabs to their English equivalents so they match the bot's month menu. Rename `Template` to `Cil` if you want to use that menu option. Keep the workbook's existing column layout.
4. Fill in `.env` and keep it private.

   | Variable | Value |
   | --- | --- |
   | `TELEGRAM_TOKEN` | Token from BotFather |
   | `OWNER_ID` | Your numeric Telegram user ID |
   | `EXCEL_FILE` | Workbook filename in `data/`, such as `finance.xlsx` |

5. Start the bot from the project root.

   ```powershell
   .\venv\Scripts\python.exe run.py
   ```

   Open a private chat with the bot, send `/start` to check that it can find the workbook, then send `/help` for the command list. Keep the process running to receive Telegram updates and send scheduled backups.

## Telegram commands

| Command | What it does |
| --- | --- |
| `/income`, `/expenses` | Compare expected and actual totals by category for a selected month. |
| `/addincome`, `/addexpense` | Create a category and set its expected amount for a selected month. |
| `/record_income`, `/record_expense` | Select a month and an existing category, then enter an amount and choose today's date or enter a date manually. |
| `/view_income`, `/view_expense` | View the transactions for a selected month. |
| `/backup` | Send a copy of the workbook to the current chat. |
| `/cancel` | Exit the active category or transaction flow. |
| `/help` | Show the command list. |

Amounts accept a decimal comma or point. Enter manual transaction dates as `DD.MM.YYYY`. The bot sends a weekly workbook backup to the owner on Sundays at 18:00 UTC when the job queue is available.

## Tech stack

Python, `python-telegram-bot` for Telegram polling and conversations, `openpyxl` for Excel storage, and APScheduler for scheduled backups. See [AGENTS.md](AGENTS.md) for the code structure and workbook cell ranges.
