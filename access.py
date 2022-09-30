from sqlite3 import connect
import pyodbc
import warnings
import pandas as pd
import win32com.client as win
from pathlib import Path

class MSAccess():
    def __init__(self, db_path: Path):
        """Create an object to interact with Access data and forms.
        
        Positional arguments:
        db_path -- path to the Access DB
        """
        self.path = db_path
        self.conn_str = (
            r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={db_path}'
        )
    
    def download_to_excel(self, tbl_name: str, destination: Path, sheet=""):
        """Download Access table as an Excel sheet.

        Positional arguments:
        tbl_name -- the name of the Access table to download
        destination -- the file path to download to

        Keyword arguments:
        sheet -- the Excel sheet name to save to (default tbl_name)
        """
        sql = f"SELECT * FROM {tbl_name}"
        data = self.run_select_sql(sql_query=sql, method="df")
        pd.io.formats.excel.ExcelFormatter.header_style = None
        if sheet == "":
            sheet = tbl_name
        data.to_excel(destination, sheet_name=sheet, index=False)

    def form_fill_run(self, form: str, *fields: str):
        """Update the Access form and run the queries.
        
        Positional arguments:
        form -- name of the Access form
        *fields -- the form's arguments in the order of the form
        """
        try:
            access = win.Dispatch('Access.Application')
            db = access.OpenCurrentDatabase(self.path)

            access.DoCmd.OpenForm(form)
            
            access.Forms(form).Fill_Form(*fields)
            access.Forms(form).RunForm_Click()
            
        except Exception as e:
            print(e)
            
        finally:
            access.DoCmd.CloseDatabase()
            access.Quit()
        

    def run_select_sql(self, sql_query: str, method="print"):
        """Return the SELECT query.
        
        Keyword arguments:
        method -- print query or return as a df (default 'print')
        """
        cnxn = pyodbc.connect(self.conn_str)

        if method == "print":
            cursor = cnxn.execute(sql_query)
            for row in cursor:
                print(row)
        elif method == "df":
            data = pd.read_sql(sql_query, cnxn)
            return data
        else:
            warnings.warn("That is not a valid method.")

        cursor.close()
        cnxn.close()
    
    def run_sql(self, sql_query: str):
        """Run the SQL query in Access.
        
        Positional arguments:
        sql_query -- the SQL query
        """
        cnxn = pyodbc.connect(self.conn_str)
        cursor = cnxn.execute(sql_query)
        cnxn.commit()
        cursor.close()
        cnxn.close()
        if "select" in sql_query.lower():
            warnings.warn("To see the output of the SELECT statement, use run_select_sql(sql_query) instead")

    def run_access_query(self, access_query: str):
        """Run the predefined Access query.
        
        Positional arguments:
        access_query -- the name of the query in Access
        """
        cnxn = pyodbc.connect(self.conn_str)
        sql = f'\u007bCALL {access_query}\u007d'
        cursor = cnxn.execute(sql)
        cnxn.commit()
        cursor.close()
        cnxn.close()

    def upload_table(self, file_loc: Path, file_sheet: str, tbl_name: str):
        """Upload Excel file to Access.
        
        Positional arguments:
        file_loc -- location of the Excel file
        file_sheet -- name of the Excel sheet the data is located in
        tbl_name -- the name to give the table in Access
        """ 
        data = pd.read_excel(file_loc, sheet_name=file_sheet)
        for col in data.columns:
            if len(col) > 25:
                data.rename(columns={col: col[0:25]}, inplace=True)

        data.to_accessdb(self.path, tbl_name)    
