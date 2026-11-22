import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
DATA_DIR = Path("data")
EXCEL_FILE = os.getenv("EXCEL_FILE", "finance.xlsx")
OWNER_ID = int(os.getenv("OWNER_ID", 0))