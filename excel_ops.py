from openpyxl import load_workbook

INCOME_SHEETS = [
    "Cil",        # Note: Change to "Target" if you renamed this sheet in Excel
    "September", 
    "October", 
    "November", 
    "December", 
    "January", 
    "February", 
    "March", 
    "April", 
    "May", 
    "June", 
    "July", 
    "August"
]

INCOME_COLS  = {"name": "C", "expected": "D", "actual": "E"}
EXPENSE_COLS = {"name": "G", "expected": "H", "actual": "I"}

# Transaction columns
INCOME_TX_COLS  = {"category": "K", "date": "L", "amount": "M"}
EXPENSE_TX_COLS = {"category": "O", "date": "P", "amount": "Q"}

FIRST_ROW = 5
LAST_ROW = 40
INCOME_TX_LAST_ROW = 40
EXPENSE_TX_LAST_ROW = 103  # <-- Added this for your 103 limit

def _detect_currency(number_format: str) -> str:
    if not number_format:
        return ""
    for symbol in ("Kč", "CZK", "€", "EUR", "$", "USD", "£", "GBP"):
        if symbol in number_format:
            return symbol
    return ""

def _read_block(backend, sheet: str, cols: dict, tx_cols: dict | None = None, tx_last_row: int = LAST_ROW):
    wb = load_workbook(backend.get_file(), data_only=True)
    if sheet not in wb.sheetnames:
        raise KeyError(f"Sheet '{sheet}' not found. Available: {wb.sheetnames}")
    ws = wb[sheet]

    tx_sums: dict[str, float] = {}
    if tx_cols:
        for r in range(FIRST_ROW, tx_last_row + 1):
            cat = ws[f"{tx_cols['category']}{r}"].value
            amt = ws[f"{tx_cols['amount']}{r}"].value
            if cat and isinstance(amt, (int, float)):
                key = str(cat).strip()
                tx_sums[key] = tx_sums.get(key, 0) + amt

    rows = []
    for r in range(FIRST_ROW, LAST_ROW + 1):
        name = ws[f"{cols['name']}{r}"].value
        if name is None or str(name).strip() == "":
            continue
        name = str(name).strip()
        exp_cell = ws[f"{cols['expected']}{r}"]
        act_cell = ws[f"{cols['actual']}{r}"]
        currency = _detect_currency(exp_cell.number_format) or _detect_currency(act_cell.number_format)

        if tx_cols:
            actual = tx_sums.get(name, 0)
        else:
            actual = act_cell.value

        rows.append({
            "row": r, "name": name, "expected": exp_cell.value,
            "actual": actual, "currency": currency,
        })
    return rows

def add_income(backend, sheet: str, name: str, expected: float) -> int:
    """Write name → C<row>, expected → D<row> on first empty row in C5:C40.
    Returns the row number written. Raises ValueError if block is full.
    """
    path = backend.get_file()
    wb = load_workbook(path)                    # NOT data_only — we're writing
    if sheet not in wb.sheetnames:
        raise KeyError(f"Sheet '{sheet}' not found. Available: {wb.sheetnames}")

    ws = wb[sheet]
    for r in range(FIRST_ROW, LAST_ROW + 1):
        cell = ws[f"{INCOME_COLS['name']}{r}"]  # C{r}
        if cell.value is None or str(cell.value).strip() == "":
            ws[f"{INCOME_COLS['name']}{r}"] = name
            ws[f"{INCOME_COLS['expected']}{r}"] = expected

            # Copy the number format from the row above (or D5) so currency
            # formatting stays consistent with the rest of the column.
            src_row = r - 1 if r > FIRST_ROW else r
            src_fmt = ws[f"{INCOME_COLS['expected']}{src_row}"].number_format
            if src_fmt and src_fmt != "General":
                ws[f"{INCOME_COLS['expected']}{r}"].number_format = src_fmt

            backend.save(wb)
            return r

    raise ValueError(f"Sheet  „{sheet}“ is full (C5:C40).")
def add_expense(backend, sheet: str, name: str, expected: float) -> int:
    """Write name → G<row>, expected → H<row> on first empty row in G5:G40.
    Returns the row number written. Raises ValueError if block is full.
    """
    path = backend.get_file()
    wb = load_workbook(path)                    # NOT data_only — we're writing
    if sheet not in wb.sheetnames:
        raise KeyError(f"Sheet '{sheet}' not found. Available: {wb.sheetnames}")

    ws = wb[sheet]
    for r in range(FIRST_ROW, LAST_ROW + 1):
        cell = ws[f"{EXPENSE_COLS['name']}{r}"]  # G{r}
        if cell.value is None or str(cell.value).strip() == "":
            ws[f"{EXPENSE_COLS['name']}{r}"] = name
            ws[f"{EXPENSE_COLS['expected']}{r}"] = expected

            src_row = r - 1 if r > FIRST_ROW else r
            src_fmt = ws[f"{EXPENSE_COLS['expected']}{src_row}"].number_format
            if src_fmt and src_fmt != "General":
                ws[f"{EXPENSE_COLS['expected']}{r}"].number_format = src_fmt

            backend.save(wb)
            return r

    raise ValueError(f"V listu „{sheet}“ je již plno (G5:G40).")

def add_income_transaction(backend, sheet: str, category: str,
                           amount: float, date_str: str) -> int:
    """Write one income transaction to first free row in K5:M40.
    K = category, L = date ('DD.MM.YYYY' string), M = amount.
    Returns the row number written.
    """
    path = backend.get_file()
    wb = load_workbook(path)
    if sheet not in wb.sheetnames:
        raise KeyError(f"Sheet '{sheet}' not found. Available: {wb.sheetnames}")

    ws = wb[sheet]
    for r in range(FIRST_ROW, LAST_ROW + 1):
        if ws[f"K{r}"].value in (None, ""):
            ws[f"K{r}"] = category
            ws[f"L{r}"] = date_str              # text, no time
            ws[f"M{r}"] = amount

            # Inherit currency format from the cell above in M so the column
            # stays visually consistent (Kč, €, etc.)
            src_row = r - 1 if r > FIRST_ROW else r
            src_fmt = ws[f"M{src_row}"].number_format
            if src_fmt and src_fmt != "General":
                ws[f"M{r}"].number_format = src_fmt

            backend.save(wb)
            return r

    raise ValueError(f"V listu „{sheet}“ je již plno (K5:K40).")

def get_income_categories(backend, sheet: str) -> list[str]:
    return [r["name"] for r in get_income(backend, sheet)]

def get_income(backend, sheet: str):
    return _read_block(backend, sheet, INCOME_COLS, tx_cols=INCOME_TX_COLS, tx_last_row=INCOME_TX_LAST_ROW)

def get_expenses(backend, sheet: str):
    return _read_block(backend, sheet, EXPENSE_COLS, tx_cols=EXPENSE_TX_COLS, tx_last_row=EXPENSE_TX_LAST_ROW)

def get_expense_categories(backend, sheet: str) -> list[str]:
    return [r["name"] for r in get_expenses(backend, sheet)]

def add_expense_transaction(backend, sheet: str, category: str, amount: float, date_str: str) -> int:
    path = backend.get_file()
    wb = load_workbook(path)
    if sheet not in wb.sheetnames:
        raise KeyError(f"Sheet '{sheet}' not found.")

    ws = wb[sheet]
    for r in range(FIRST_ROW, EXPENSE_TX_LAST_ROW + 1):
        if ws[f"O{r}"].value in (None, ""):
            ws[f"O{r}"] = category
            ws[f"P{r}"] = date_str
            ws[f"Q{r}"] = amount

            # Inherit currency format from H (expected) for consistency
            src_fmt = ws[f"{EXPENSE_COLS['expected']}{r}"].number_format
            if src_fmt and src_fmt != "General":
                ws[f"Q{r}"].number_format = src_fmt

            backend.save(wb)
            return r
    raise ValueError(f"V listu „{sheet}“ je již plno (O5:O103).")

def get_income_transactions(backend, sheet: str) -> list[dict]:
    """Read K5:K40 (category), L5:L40 (date), M5:M40 (amount)."""
    wb = load_workbook(backend.get_file(), data_only=True)
    if sheet not in wb.sheetnames:
        raise KeyError(f"Sheet '{sheet}' not found.")

    ws = wb[sheet]
    transactions = []
    
    for r in range(FIRST_ROW, INCOME_TX_LAST_ROW + 1):
        cat = ws[f"{INCOME_TX_COLS['category']}{r}"].value  # Column K
        if cat is None or str(cat).strip() == "":
            continue
            
        date_val = ws[f"{INCOME_TX_COLS['date']}{r}"].value  # Column L
        amount_cell = ws[f"{INCOME_TX_COLS['amount']}{r}"]   # Column M
        currency = _detect_currency(amount_cell.number_format)
        
        transactions.append({
            "row": r,
            "category": str(cat).strip(),
            "date": str(date_val) if date_val else "—",
            "amount": amount_cell.value,
            "currency": currency,
        })
        
    return transactions

def get_expense_transactions(backend, sheet: str) -> list[dict]:
    """Read O5:O103 (category), P5:P103 (date), Q5:Q103 (amount)."""
    wb = load_workbook(backend.get_file(), data_only=True)
    if sheet not in wb.sheetnames:
        raise KeyError(f"Sheet '{sheet}' not found.")

    ws = wb[sheet]
    transactions = []
    
    # We use EXPENSE_TX_LAST_ROW here to read up to row 103
    for r in range(FIRST_ROW, EXPENSE_TX_LAST_ROW + 1):
        cat = ws[f"{EXPENSE_TX_COLS['category']}{r}"].value  # Column O
        if cat is None or str(cat).strip() == "":
            continue
            
        date_val = ws[f"{EXPENSE_TX_COLS['date']}{r}"].value  # Column P
        amount_cell = ws[f"{EXPENSE_TX_COLS['amount']}{r}"]   # Column Q
        currency = _detect_currency(amount_cell.number_format)
        
        transactions.append({
            "row": r,
            "category": str(cat).strip(),
            "date": str(date_val) if date_val else "—",
            "amount": amount_cell.value,
            "currency": currency,
        })
        
    return transactions

def op1_read_cell(backend, sheet: str, cell: str):
    wb = load_workbook(backend.get_file())
    return wb[sheet][cell].value


def op2_write_cell(backend, sheet: str, cell: str, value):
    wb = load_workbook(backend.get_file())
    wb[sheet][cell] = value
    backend.save(wb)


def op3_add_row(backend, sheet: str, row: list):
    wb = load_workbook(backend.get_file())
    wb[sheet].append(row)
    backend.save(wb)


def op4_get_sheet_names(backend):
    wb = load_workbook(backend.get_file())
    return wb.sheetnames


def op5_read_range(backend, sheet: str, cell_range: str):
    wb = load_workbook(backend.get_file(), data_only=True)
    return [[c.value for c in row] for row in wb[sheet][cell_range]]


def op6_delete_row(backend, sheet: str, row_number: int):
    wb = load_workbook(backend.get_file())
    wb[sheet].delete_rows(row_number)
    backend.save(wb)