from pathlib import Path
from tempfile import TemporaryDirectory
from openpyxl import load_workbook

import config


class LocalBackend:
    def __init__(self):
        self.path = config.DATA_DIR / config.EXCEL_FILE   # 👈 the one file

    def get_file(self) -> Path:
        if not self.path.exists():
            raise FileNotFoundError(
                f"{self.path} not found. Set EXCEL_FILE in .env and put the file in data/"
            )
        return self.path

    def save(self, wb) -> None:
        wb.save(self.path)

    def replace_workbook(self, contents: bytes | bytearray) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with TemporaryDirectory(dir=self.path.parent) as temp_dir:
            staged_file = Path(temp_dir) / self.path.name
            staged_file.write_bytes(contents)
            staged_file.replace(self.path)

    def list_files(self) -> list[str]:        # still handy for debugging
        return [p.name for p in config.DATA_DIR.glob("*.xlsx")]



def get_backend(name: str):
    if name == "local":
        return LocalBackend()
    raise ValueError(f"Unknown backend: {name}")
