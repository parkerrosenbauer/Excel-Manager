import pandas as pd

class ExcelMgr():
    def __init__(self, path: str):
        """Create excel object.

        path -- path to excel file
        """
        self.path = path
        
    @property
    def sheets(self):
        xl = pd.ExcelFile(self.path)
        return list(xl.sheet_names)
        
    def __contains__(self, sheet: str):
        """Return True if sheet is in the excel file."""
        return sheet in self.sheets

class SheetMgr(ExcelMgr):
    def __init__(self, path: str, sheet: str):
        """Create excel sheet object.

        path -- path to excel file
        sheet -- sheet in excel file
        """
        super().__init__(path)
        self.sheet = sheet
        self.data = pd.read_excel(self.path, sheet_name=self.sheet)
        self.columns = list(self.data.columns)
        self.record_count = len(self.data.index)

    def column_row_count(self, column: str):
        """Return count of non null values in the specified column."""
        return self.data[column].count()
