import pandas as pd
from pathlib import Path

class ExcelMgr:
    def __init__(self, path: Path) -> None:
        self.path = path

    def __contains__(self, sheet: str) -> bool:
        return sheet in self.sheets

    @property
    def sheets(self) -> list:
        xl = pd.ExcelFile(self.path)
        return (list(xl.sheet_names))


class SheetMgr():
    def __init__(self, path: Path, sheet: str) -> None:
        self.path = path
        self.sheet = sheet

    def __contains__(self, col: str) -> bool:
        return col in self.columns

    @property
    def data(self) -> pd.DataFrame:
        return pd.read_excel(self.path, sheet_name=self.sheet)

    @property
    def record_count(self) -> int:
        return len(self.data.index)

    @property
    def columns(self) -> list:
        return list(self.data.columns)

    def column_row_count(self, column: str) -> int:
        return self.data[column].count()
