import pandas as pd
from pathlib import Path

class ExcelMgr:
    def __init__(self, path: Path):
        """Create ExcelMgr object.

        path -- path to excel file
        """
        self.path = path

    def __contains__(self, sheet: str):
        """Return true if sheet present in excel file.

        sheet -- sheet to search for
        """
        return sheet in self.sheets

    def __getattr__(self, attr):
        """Return attribute or SheetMgr object when using dot notation."""
        attribute = attr in self.__dict__.keys()
        worksheet = attr in self.sheets

        if attribute and worksheet:
            raise ReferenceError(
                f'{attr} references both an ExcelMgr attribute and a sheet in the excel file.')
        elif attribute:
            return attr
        else:
            # if the sheet is not a worksheet, the SheetMgr class will throw an error
            return SheetMgr(self.path, attr)

    def __getitem__(self, item):
        """Return SheetMgr object when using brackets."""
        if item in self.__dict__.keys():
            raise ReferenceError(
                f'{item} is an attribute, not a sheet in the file. Use dot notation to access this objects attributes.')
        else:
            # if the sheet is not a worksheet, the SheetMgr class will throw an error
            return SheetMgr(self.path, item)

    @property
    def sheets(self):
        """Return sheets in excel file"""
        xl = pd.ExcelFile(self.path)
        return (list(xl.sheet_names))


class SheetMgr(ExcelMgr):
    def __init__(self, path: Path, sheet: str):
        """Create SheetMgr object.

        path -- path to excel file
        sheet -- sheet to read from
        """
        super().__init__(path)
        if sheet in self.sheets:
            self.sheet = sheet
        else:
            raise LookupError('That sheet name does not exist in file.')

    def __contains__(self, column: str):
        """Return True if column present in sheet."""
        return column in self.columns

    @property
    def columns(self):
        """Return list of columns in sheet."""
        return list(self.data.columns)

    @property
    def data(self):
        """Return a dataframe of the sheet"""
        return pd.read_excel(self.path, sheet_name=self.sheet)

    @property
    def record_count(self):
        """Return count of rows in sheet"""
        return len(self.data.index)

    def column_row_count(self, column: str) -> int:
        """Return count of items in column.
        
        column -- column containing items
        """
        return self.data[column].count()
