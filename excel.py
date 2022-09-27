import pandas as pd
import os
import re


class ExcelMgr:
    def __init__(self, path: str):
        """Create ExcelMgr object.
        path -- path to Excel file
        """
        self.path = path

    def __contains__(self, sheet: str):
        """Return true if sheet present in Excel file.
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
        """Return sheets in Excel file"""
        xl = pd.ExcelFile(self.path)
        return list(xl.sheet_names)


class SheetMgr(ExcelMgr):
    def __init__(self, path: str, sheet: str):
        """Create SheetMgr object.
        path -- path to excel file
        sheet -- sheet to read from
        """
        super().__init__(path)
        if sheet in self.sheets:
            self.sheet = sheet
        else:
            raise LookupError(f'{sheet} does not exist in file.')

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

    def _col_nexist(func):
        def wrapper(self, column, *args, **kwargs):
            if column not in self.columns:
                raise LookupError(f'That column does not exist in {self.sheet}')
            return func(self, column, *args, **kwargs)
        return wrapper

    @_col_nexist
    def column_row_count(self, column: str) -> int:
        """Return count of items in column.
        
        column -- column containing items
        """
        return self.data[column].count()

    @_col_nexist
    def column_unique_vals(self, column: str) -> list:
        """Return unique values in a column.

        column -- column containing items
        """
        return list(self.data[column].unique())

    def save_to_csv(self) -> None:
        """Save sheet data to csv in format FileName-SheetName."""
        pd.io.formats.excel.ExcelFormatter.header_style = None

        base_name = re.sub('.xlsx$', '', os.path.basename(self.path))
        bare_path = os.path.dirname(self.path)
        csv_path = os.path.join(bare_path, f"{base_name}-{self.sheet}.csv")
        self.data.to_csv(csv_path, index=False)
