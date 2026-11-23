import html
import os
from datetime import datetime, time
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.constants import FileSizeLimit
from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler, MessageHandler,
    ConversationHandler, filters, ContextTypes, ApplicationHandlerStop, TypeHandler
)
from telegram.error import BadRequest, TelegramError

import config
import excel_ops
from backend import get_backend
from datetime import datetime


# Conversation states
CHOOSE_MONTH, ENTER_NAME, ENTER_AMOUNT = range(3)       # for add category
T_MONTH, T_CATEGORY, T_AMOUNT, T_DATE = range(10, 14)   # for income tx
TE_MONTH, TE_CATEGORY, TE_AMOUNT, TE_DATE = range(20, 24) # for expense tx
UPLOAD_FILE = 30
# reuse the same three state names — the conversations are independent,
# but each ConversationHandler tracks its own state per user, so this is fine.

# ---------- helpers for the income table ----------

def _fmt_number(value, currency: str) -> str:
    if value is None or value == "":
        return "—"
    if isinstance(value, (int, float)):
        if float(value).is_integer():
            s = f"{int(value):,}".replace(",", " ")          # 15 000
        else:
            s = f"{value:,.2f}".replace(",", " ").replace(".", ",")  # 1 234,56
        return f"{s} {currency}".strip() if currency else s
    return str(value)
def _parse_date(raw: str) -> str | None:
    """Accept 27.10.2025 / 27. 10. 2025 / 27/10/2025 / 27-10-2025.
    Returns normalized 'DD.MM.YYYY' or None if unparseable.
    """
    raw = raw.strip().replace(" ", "")
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%d.%m.%Y")
        except ValueError:
            continue
    return None

def _format_table(rows: list[dict], sheet: str, title: str) -> str:
    name_w, exp_w, act_w = 22, 14, 14
    header = f"{'Name':<{name_w}} {'Expected':>{exp_w}} {'Actual':>{act_w}}"
    sep = "─" * len(header)

    lines = [header, sep]
    for r in rows:
        name = r["name"][:name_w]
        exp = _fmt_number(r["expected"], r["currency"])
        act = _fmt_number(r["actual"], r["currency"])
        lines.append(f"{name:<{name_w}} {exp:>{exp_w}} {act:>{act_w}}")

    table = html.escape("\n".join(lines))
    return f"📊 <b>{html.escape(title)} — {html.escape(sheet)}</b>\n\n<pre>{table}</pre>"

def _format_transactions_table(rows: list[dict], sheet: str, title: str) -> str:
    """Build a monospace table for individual transactions."""
    date_w, cat_w, amt_w = 12, 20, 14
    header = f"{'Date':<{date_w}} {'Category':<{cat_w}} {'Amount':>{amt_w}}"
    sep = "─" * len(header)

    lines = [header, sep]
    for r in rows:
        date = r["date"][:date_w]
        cat = r["category"][:cat_w]
        amt = _fmt_number(r["amount"], r["currency"])
        lines.append(f"{date:<{date_w}} {cat:<{cat_w}} {amt:>{amt_w}}")

    table = html.escape("\n".join(lines))
    return f"📋 <b>{html.escape(title)} — {html.escape(sheet)}</b>\n\n<pre>{table}</pre>"

# ---------- the app ----------

def build_app(backend):
    app = Application.builder().token(config.TELEGRAM_TOKEN).build()

    async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        try:
            sheets = excel_ops.op4_get_sheet_names(backend)
            await update.message.reply_text(
                f"📊 Welcome, Owner!\nFile: {config.EXCEL_FILE}\n"
                f"Backend: {type(backend).__name__}\n"
                f"Sheets: {sheets}"
            )
        except FileNotFoundError as e:
            await update.message.reply_text(f"⚠️ {e}")
    async def help_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        help_text = (
            "🤖 <b>Help - Command List</b>\n\n"
            "<b>📊 Monthly Summaries:</b>\n"
            "/income - View expected vs. actual income\n"
            "/expenses - View expected vs. actual expenses\n\n"
            
            "<b>📋 Transaction History:</b>\n"
            "/view_income - List all income transactions\n"
            "/view_expense - List all expense transactions\n\n"
            
            "<b>➕ Manage Categories:</b>\n"
            "/addincome - Add a new income category\n"
            "/addexpense - Add a new expense category\n\n"
            
            "<b>✍️ Record Transactions:</b>\n"
            "/record_income - Log a new income transaction\n"
            "/record_expense - Log a new expense transaction\n\n"
            
            "<b>🛠 Data & Settings:</b>\n"
            "/backup - Send the Excel file to this chat\n"
            "/upload - Replace the finance workbook with an Excel file\n"
            "/cancel - Cancel the current action\n"
            "/help - Show this help menu"
        )
        
        await update.message.reply_text(help_text, parse_mode="HTML")
    # ---- shared: month keyboard + table reply ----

    def _month_keyboard(prefix: str) -> InlineKeyboardMarkup:
        keyboard, row = [], []
        for s in excel_ops.INCOME_SHEETS:
            row.append(InlineKeyboardButton(s, callback_data=f"{prefix}:{s}"))
            if len(row) == 3:
                keyboard.append(row)
                row = []
        if row:
            keyboard.append(row)
        return InlineKeyboardMarkup(keyboard)

    async def _send_month_table(query, sheet: str, kind: str):
        """kind = 'income' | 'expenses'"""
        try:
            if kind == "income":
                rows = excel_ops.get_income(backend, sheet)
                title = "Income"
                empty_msg = f"There are no incomes in sheet „{sheet}“ (C5:E40 is empty)."
            else:
                rows = excel_ops.get_expenses(backend, sheet)
                title = "Expense"
                empty_msg = f"There are no expenses in sheet „{sheet}“ (G5:I40 is empty)."
        except (FileNotFoundError, KeyError) as e:
            await query.edit_message_text(f"⚠️ {e}")
            return
        except Exception as e:
            await query.edit_message_text(f"⚠️ Error: {e}")
            return

        if not rows:
            await query.edit_message_text(empty_msg)
            return

        await query.edit_message_text(
            _format_table(rows, sheet, title),
            parse_mode="HTML",
        )

    # ---- /income ----

    async def income_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        await update.message.reply_text(
            "Select a month for income:",
            reply_markup=_month_keyboard("inc"),
        )

    async def income_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        query = update.callback_query
        await query.answer()
        sheet = query.data.split(":", 1)[1]
        await _send_month_table(query, sheet, kind="income")

    # ---- /expenses ----

    async def expenses_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        await update.message.reply_text(
           "Select a month for expenses:",
            reply_markup=_month_keyboard("exp"),
        )

    async def expenses_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        query = update.callback_query
        await query.answer()
        sheet = query.data.split(":", 1)[1]
        await _send_month_table(query, sheet, kind="expenses")
        # ---- /addincome (conversation) ----

    async def addincome_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        await update.message.reply_text(
            "Select a month for new income:",
            reply_markup=_month_keyboard("addinc_month"),
        )
        return CHOOSE_MONTH

    async def addincome_month(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        query = update.callback_query
        await query.answer()
        sheet = query.data.split(":", 1)[1]
        ctx.user_data["sheet"] = sheet
        await query.edit_message_text(
            f"Month: <b>{html.escape(sheet)}</b>\n\n"
            f"Enter the name of the income category:",
            parse_mode="HTML",
        )
        return ENTER_NAME

    async def addincome_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        name = update.message.text.strip()
        if not name:
            await update.message.reply_text("The name cannot be empty. Try again:")
            return ENTER_NAME
        ctx.user_data["name"] = name
        await update.message.reply_text(
            f"Name: <b>{html.escape(name)}</b>\n\n"
            f"Enter the expected amount (e.g. <code>15000</code> or <code>15000,50</code>):",
            parse_mode="HTML",
        )
        return ENTER_AMOUNT

    async def addincome_amount(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        raw = update.message.text.strip().replace(" ", "").replace(",", ".")
        try:
            amount = float(raw)
        except ValueError:
            await update.message.reply_text(
                "I don't understand the amount. Try again (e.g. 15000 or 15000,50):"
            )
            return ENTER_AMOUNT

        sheet = ctx.user_data.pop("sheet", None)
        name = ctx.user_data.pop("name", None)
        if not sheet or not name:
            await update.message.reply_text("⚠️ Something got lost. Run /addincome again.")
            return ConversationHandler.END

        try:
            row = excel_ops.add_income(backend, sheet, name, amount)
        except (FileNotFoundError, KeyError, ValueError) as e:
            await update.message.reply_text(f"⚠️ {e}")
            return ConversationHandler.END

        await update.message.reply_text(
            f"✅ Added to <b>{html.escape(sheet)}</b>, row {row}:\n"
            f"{html.escape(name)} — {_fmt_number(amount, '')}",
            parse_mode="HTML",
        )
        return ConversationHandler.END

    async def cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        ctx.user_data.clear()
        await update.message.reply_text("Cancelled.")
        return ConversationHandler.END

    addincome_conv = ConversationHandler(
        entry_points=[CommandHandler("addincome", addincome_cmd)],
        states={
            CHOOSE_MONTH: [
                CallbackQueryHandler(addincome_month, pattern=r"^addinc_month:"),
            ],
            ENTER_NAME: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, addincome_name),
            ],
            ENTER_AMOUNT: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, addincome_amount),
            ],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        # If the user abandons the flow for more than 5 min, reset state
        conversation_timeout=300,
    )
        # ---- /addexpense (conversation) ----

    async def addexpense_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        await update.message.reply_text(
            "Select a month for new expense:",
            reply_markup=_month_keyboard("addexp_month"),
        )
        return CHOOSE_MONTH

    async def addexpense_month(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        query = update.callback_query
        await query.answer()
        sheet = query.data.split(":", 1)[1]
        ctx.user_data["sheet"] = sheet
        await query.edit_message_text(
            f"Month: <b>{html.escape(sheet)}</b>\n\n"
            f"Enter the name of the expense category:",
            parse_mode="HTML",
        )
        return ENTER_NAME

    async def addexpense_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        name = update.message.text.strip()
        if not name:
            await update.message.reply_text("The name cannot be empty. Try again:")
            return ENTER_NAME
        ctx.user_data["name"] = name
        await update.message.reply_text(
            f"Name: <b>{html.escape(name)}</b>\n\n"
            f"Enter the expected amount (e.g. <code>2500</code> or <code>2500,50</code>):",
            parse_mode="HTML",
        )
        return ENTER_AMOUNT

    async def addexpense_amount(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        raw = update.message.text.strip().replace(" ", "").replace(",", ".")
        try:
            amount = float(raw)
        except ValueError:
            await update.message.reply_text(
                "I don't understand the amount. Try again (e.g. 2500 or 2500,50):"
            )
            return ENTER_AMOUNT

        sheet = ctx.user_data.pop("sheet", None)
        name = ctx.user_data.pop("name", None)
        if not sheet or not name:
            await update.message.reply_text("⚠️ Something got lost. Run /addexpense again.")
            return ConversationHandler.END

        try:
            row = excel_ops.add_expense(backend, sheet, name, amount)
        except (FileNotFoundError, KeyError, ValueError) as e:
            await update.message.reply_text(f"⚠️ {e}")
            return ConversationHandler.END

        await update.message.reply_text(
            f"✅ Added to <b>{html.escape(sheet)}</b>, row {row}:\n"
            f"{html.escape(name)} — {_fmt_number(amount, '')}",
            parse_mode="HTML",
        )
        return ConversationHandler.END

    addexpense_conv = ConversationHandler(
        entry_points=[CommandHandler("addexpense", addexpense_cmd)],
        states={
            CHOOSE_MONTH: [
                CallbackQueryHandler(addexpense_month, pattern=r"^addexp_month:"),
            ],
            ENTER_NAME: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, addexpense_name),
            ],
            ENTER_AMOUNT: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, addexpense_amount),
            ],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        conversation_timeout=300,
    )

        # ---- /addtransaction (income) ----

    async def addtraninc_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        await update.message.reply_text(
            "Select a month for a new income transaction:",
            reply_markup=_month_keyboard("tin_month"),
        )
        return T_MONTH

    async def addtraninc_month(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        query = update.callback_query
        await query.answer()
        sheet = query.data.split(":", 1)[1]
        ctx.user_data["sheet"] = sheet

        try:
            categories = excel_ops.get_income_categories(backend, sheet)
        except (FileNotFoundError, KeyError) as e:
            await query.edit_message_text(f"⚠️ {e}")
            return ConversationHandler.END

        if not categories:
            await query.edit_message_text(
                f"There are no income categories in sheet „{sheet}“.\n"
                f"First add them via /addincome."
            )
            return ConversationHandler.END

        ctx.user_data["categories"] = categories
        keyboard = [
            [InlineKeyboardButton(c, callback_data=f"tin_cat:{i}")]
            for i, c in enumerate(categories)
        ]
        await query.edit_message_text(
            f"Month: <b>{html.escape(sheet)}</b>\n\nSelect a category:",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup(keyboard),
        )
        return T_CATEGORY

    async def addtraninc_category(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        query = update.callback_query
        await query.answer()
        idx = int(query.data.split(":", 1)[1])
        categories = ctx.user_data.get("categories", [])
        if idx >= len(categories):
            await query.edit_message_text("⚠️ Invalid choice.")
            return ConversationHandler.END

        ctx.user_data["category"] = categories[idx]
        await query.edit_message_text(
            f"Month: <b>{html.escape(ctx.user_data['sheet'])}</b>\n"
            f"Category: <b>{html.escape(categories[idx])}</b>\n\n"
            f"Enter the transaction amount:",
            parse_mode="HTML",
        )
        return T_AMOUNT

    async def addtraninc_amount(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        raw = update.message.text.strip().replace(" ", "").replace(",", ".")
        try:
            amount = float(raw)
        except ValueError:
            await update.message.reply_text("I don't understand the amount. Try again:")
            return T_AMOUNT
        ctx.user_data["amount"] = amount

        today = datetime.now().strftime("%d.%m.%Y")
        await update.message.reply_text(
            f"Amount: <b>{_fmt_number(amount, '')}</b>\n\n"
            f"Select „Today“ or enter the date in the format <code>DD.MM.YYYY</code>:",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton(f"Dnes ({today})", callback_data="tin_date:today"),
            ]]),
        )
        return T_DATE

    async def addtraninc_date_today(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        query = update.callback_query
        await query.answer()
        today = datetime.now().strftime("%d.%m.%Y")
        await _save_income_tran(ctx, today, query.edit_message_text)
        return ConversationHandler.END  # <--- ADD THIS LINE!

    async def addtraninc_date_manual(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        date_str = _parse_date(update.message.text)
        if date_str is None:
            await update.message.reply_text(
                "I don't understand the date. Format: DD.MM.YYYY (e.g. 27.10.2025):"
            )
            return T_DATE
        await _save_income_tran(ctx, date_str, update.message.reply_text)
        return ConversationHandler.END

    async def _save_income_tran(ctx, date_str: str, reply):
        sheet = ctx.user_data.pop("sheet", None)
        category = ctx.user_data.pop("category", None)
        amount = ctx.user_data.pop("amount", None)
        ctx.user_data.pop("categories", None)

        if not sheet or not category or amount is None:
            await reply("⚠️ Something got lost. Run /addtransaction again.")
            return

        try:
            row = excel_ops.add_income_transaction(
                backend, sheet, category, amount, date_str
            )
        except (FileNotFoundError, KeyError, ValueError) as e:
            await reply(f"⚠️ {e}")
            return

        await reply(
            f"✅ Transaction recorded in <b>{html.escape(sheet)}</b>, row {row}:\n"
            f"Date: {html.escape(date_str)}\n"
            f"Category: {html.escape(category)}\n"
            f"Amount: {_fmt_number(amount, '')}",
            parse_mode="HTML",
        )

    addtraninc_conv = ConversationHandler(
        entry_points=[CommandHandler("record_income", addtraninc_cmd)],
        states={
            T_MONTH:    [CallbackQueryHandler(addtraninc_month,    pattern=r"^tin_month:")],
            T_CATEGORY: [CallbackQueryHandler(addtraninc_category, pattern=r"^tin_cat:")],
            T_AMOUNT:   [MessageHandler(filters.TEXT & ~filters.COMMAND, addtraninc_amount)],
            T_DATE: [
                CallbackQueryHandler(addtraninc_date_today, pattern=r"^tin_date:today$"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, addtraninc_date_manual),
            ],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        conversation_timeout=300,
    )
    # ---- /addexpensetx (expense transaction) ----

    async def addexpensetx_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        await update.message.reply_text(
            "Select a month for a new expense transaction:",
            reply_markup=_month_keyboard("tex_month"),
        )
        return TE_MONTH

    async def addexpensetx_month(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        query = update.callback_query
        await query.answer()
        sheet = query.data.split(":", 1)[1]
        ctx.user_data["sheet"] = sheet

        try:
            categories = excel_ops.get_expense_categories(backend, sheet)
        except (FileNotFoundError, KeyError) as e:
            await query.edit_message_text(f"⚠️ {e}")
            return ConversationHandler.END

        if not categories:
            await query.edit_message_text(
                f"There are no expense categories in sheet „{sheet}“.\n",
                f"First add them via /addexpense."
            )
            return ConversationHandler.END

        ctx.user_data["categories"] = categories
        keyboard = [[InlineKeyboardButton(c, callback_data=f"tex_cat:{i}")] for i, c in enumerate(categories)]
        await query.edit_message_text(
            f"Month: <b>{html.escape(sheet)}</b>\n\nSelect an expense category:",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup(keyboard),
        )
        return TE_CATEGORY

    async def addexpensetx_category(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        query = update.callback_query
        await query.answer()
        idx = int(query.data.split(":", 1)[1])
        categories = ctx.user_data.get("categories", [])
        
        ctx.user_data["category"] = categories[idx]
        await query.edit_message_text(
            f"Month: <b>{html.escape(ctx.user_data['sheet'])}</b>\n"
            f"Category: <b>{html.escape(categories[idx])}</b>\n\n"
            f"Enter the transaction amount",
            parse_mode="HTML",
        )
        return TE_AMOUNT

    async def addexpensetx_amount(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        raw = update.message.text.strip().replace(" ", "").replace(",", ".")
        try:
            amount = float(raw)
        except ValueError:
            await update.message.reply_text("I don't understand the amount. Try again:")
            return TE_AMOUNT
        ctx.user_data["amount"] = amount

        today = datetime.now().strftime("%d.%m.%Y")
        await update.message.reply_text(
            f"Amount: <b>{_fmt_number(amount, '')}</b>\n\n"
            f"Select „Today“ or enter the date in the format <code>DD.MM.YYYY</code>:",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton(f"Dnes ({today})", callback_data="tex_date:today"),
            ]]),
        )
        return TE_DATE

    async def addexpensetx_date_today(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        query = update.callback_query
        await query.answer()
        today = datetime.now().strftime("%d.%m.%Y")
        await _save_expense_tran(ctx, today, query.edit_message_text)
        return ConversationHandler.END  # <--- ADD THIS LINE!

    async def addexpensetx_date_manual(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        date_str = _parse_date(update.message.text)
        if date_str is None:
            await update.message.reply_text("I don't understand the date. Format: DD.MM.YYYY (e.g. 27.10.2025):")
            return TE_DATE
        await _save_expense_tran(ctx, date_str, update.message.reply_text)
        return ConversationHandler.END

    async def _save_expense_tran(ctx, date_str: str, reply):
        sheet = ctx.user_data.pop("sheet", None)
        category = ctx.user_data.pop("category", None)
        amount = ctx.user_data.pop("amount", None)
        ctx.user_data.pop("categories", None)

        try:
            row = excel_ops.add_expense_transaction(backend, sheet, category, amount, date_str)
        except (FileNotFoundError, KeyError, ValueError) as e:
            await reply(f"⚠️ {e}")
            return

        await reply(
            f"✅ Expense transaction recorded in <b>{html.escape(sheet)}</b>, row {row}:\n"
            f"Date: {html.escape(date_str)}\n"
            f"Category: {html.escape(category)}\n"
            f"Amount: {_fmt_number(amount, '')}",
            parse_mode="HTML",
        )

    addexpensetx_conv = ConversationHandler(
        entry_points=[CommandHandler("record_expense", addexpensetx_cmd)],
        states={
            TE_MONTH:    [CallbackQueryHandler(addexpensetx_month,    pattern=r"^tex_month:")],
            TE_CATEGORY: [CallbackQueryHandler(addexpensetx_category, pattern=r"^tex_cat:")],
            TE_AMOUNT:   [MessageHandler(filters.TEXT & ~filters.COMMAND, addexpensetx_amount)],
            TE_DATE: [
                CallbackQueryHandler(addexpensetx_date_today, pattern=r"^tex_date:today$"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, addexpensetx_date_manual),
            ],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        conversation_timeout=300,
        allow_reentry=True,  # <--- ADD THIS TO ALL 4 CONVERSATIONS
    )
    # ---- /view_income ----

    async def view_income_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        await update.message.reply_text(
            "Select a month to display income transactions:",
            reply_markup=_month_keyboard("viewinc_month"),
        )

    async def view_income_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        query = update.callback_query
        await query.answer()
        sheet = query.data.split(":", 1)[1]

        try:
            rows = excel_ops.get_income_transactions(backend, sheet)
        except (FileNotFoundError, KeyError) as e:
            await query.edit_message_text(f"⚠️ {e}")
            return

        if not rows:
            try:
                await query.edit_message_text(
                    f"There are no income transactions in sheet „{sheet}“."
                )
            except BadRequest:
                pass # Ignore if user clicks the same month twice
            return

        try:
            await query.edit_message_text(
                _format_transactions_table(rows, sheet, "Income transactions"),
                parse_mode="HTML",
            )
        except BadRequest as e:
            if "Message is not modified" in str(e):
                pass # Ignore this specific error
            else:
                raise e # Raise other actual errors
    # ---- /view_expense ----

    async def view_expense_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        await update.message.reply_text(
            "Select a month to display expense transactions:",
            reply_markup=_month_keyboard("viewexp_month"),
        )

    async def view_expense_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        query = update.callback_query
        await query.answer()
        sheet = query.data.split(":", 1)[1]

        try:
            rows = excel_ops.get_expense_transactions(backend, sheet)
        except (FileNotFoundError, KeyError) as e:
            await query.edit_message_text(f"⚠️ {e}")
            return

        if not rows:
            await query.edit_message_text(
                f"There are no expense transactions in sheet „{sheet}“."
            )
            return

        # Reusing the same table formatter, just passing a different title
        await query.edit_message_text(
            _format_transactions_table(rows, sheet, "Expense transactions"),
            parse_mode="HTML",
        )
    async def get_my_id(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        user_id = update.effective_user.id
        await update.message.reply_text(f"Your permanent User ID is: `{user_id}`", parse_mode="Markdown")
    # ---- Security Guard ----
    
    async def restrict_access(update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Block anyone who is not the OWNER_ID."""
        user = update.effective_user
        if not user or user.id != config.OWNER_ID:
            # If a stranger messages the bot, send them a warning (optional)
            if update.effective_message:
                await update.effective_message.reply_text("⛔ Unauthorized. This is a private bot.")
            
            # Stop the bot from processing any further handlers
            raise ApplicationHandlerStop
    # ---- Workbook upload ----

    async def upload_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
        await update.message.reply_text(
            f"Send an .xlsx document to replace {config.EXCEL_FILE} in data/. "
            "Use /backup first if you want a copy, or /cancel to stop."
        )
        return UPLOAD_FILE

    async def upload_document(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
        message = update.message
        document = message.document
        if not (document.file_name or "").lower().endswith(".xlsx"):
            await message.reply_text("Send an .xlsx document, or use /cancel.")
            return UPLOAD_FILE
        if document.file_size and document.file_size > FileSizeLimit.FILESIZE_DOWNLOAD:
            await message.reply_text("Send a workbook of 20 MB or less, or use /cancel.")
            return UPLOAD_FILE

        try:
            telegram_file = await document.get_file()
            contents = await telegram_file.download_as_bytearray()
        except TelegramError:
            await message.reply_text("Could not download the file. Send it again or use /cancel.")
            return UPLOAD_FILE

        try:
            excel_ops.validate_upload(contents)
            backend.replace_workbook(contents)
        except ValueError as exc:
            await message.reply_text(f"{exc} Send a compatible workbook or use /cancel.")
            return UPLOAD_FILE
        except OSError:
            await message.reply_text(
                "Could not replace the workbook. If it is open in Excel, close it and try again."
            )
            return UPLOAD_FILE

        await message.reply_text(f"Uploaded {config.EXCEL_FILE}. The bot is now using it.")
        return ConversationHandler.END

    async def cancel_upload(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
        await update.message.reply_text("Upload cancelled.")
        return ConversationHandler.END

    upload_conv = ConversationHandler(
        entry_points=[CommandHandler("upload", upload_cmd)],
        states={UPLOAD_FILE: [MessageHandler(filters.Document.ALL, upload_document)]},
        fallbacks=[CommandHandler("cancel", cancel_upload)],
    )

    # ---- Backups ----

    async def backup_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        """Manual backup triggered by the /backup command."""
        try:
            file_path = backend.get_file()
            with open(file_path, "rb") as f:
                await update.message.reply_document(
                    document=f, 
                    filename=config.EXCEL_FILE,
                    caption="📁 Here is your manual backup."
                )
        except Exception as e:
            await update.message.reply_text(f"⚠️ Failed to backup: {e}")

    async def send_scheduled_backup(context: ContextTypes.DEFAULT_TYPE):
        """Automated background backup."""
        try:
            file_path = backend.get_file()
            with open(file_path, "rb") as f:
                await context.bot.send_document(
                    chat_id=config.OWNER_ID,
                    document=f,
                    filename=config.EXCEL_FILE,
                    caption="📁 Here is your automated weekly backup."
                )
            print("✅ Weekly backup sent successfully.")
        except Exception as e:
            print(f"⚠️ Failed to send scheduled backup: {e}")
    # ---- register ----
    # Register the guard in group -1 (it runs BEFORE all other handlers)
    app.add_handler(TypeHandler(Update, restrict_access), group=-1)
    # Register it temporarily
    #app.add_handler(CommandHandler("myid", get_my_id))
    # 1. Conversations go first! They need to intercept messages 
    # before the generic command handlers do.
    app.add_handler(upload_conv)
    app.add_handler(addincome_conv)
    app.add_handler(addexpense_conv)
    app.add_handler(addtraninc_conv)
    app.add_handler(addexpensetx_conv)

    # 2. Standard Commands
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("income", income_cmd))
    app.add_handler(CommandHandler("expenses", expenses_cmd))
    app.add_handler(CommandHandler("view_income", view_income_cmd))
    app.add_handler(CommandHandler("view_expense", view_expense_cmd))
    app.add_handler(CommandHandler("backup", backup_cmd))

    # 3. Callbacks for the months (Income / Expense summaries)
    app.add_handler(CallbackQueryHandler(income_callback, pattern=r"^inc:"))
    app.add_handler(CallbackQueryHandler(expenses_callback, pattern=r"^exp:"))

    # 4. Callbacks for the month (View Income / View Expense lists)
    app.add_handler(CallbackQueryHandler(view_income_callback, pattern=r"^viewinc_month:"))
    app.add_handler(CallbackQueryHandler(view_expense_callback, pattern=r"^viewexp_month:"))

    # 5. Schedule the backup (Every Sunday at 18:00 Server Time)
    if app.job_queue:
        app.job_queue.run_daily(
            send_scheduled_backup,
            time=time(hour=18, minute=0), 
            days=(6,) # 6 = Sunday
        )
    else:
        print("⚠️ WARNING: JobQueue is missing. Automated backups will not run.")

    return app

    return app


def run(backend_name: str):
    backend = get_backend(backend_name)
    app = build_app(backend)
    print(f"Starting bot with {backend_name} backend, file={config.EXCEL_FILE}")
    app.run_polling()
