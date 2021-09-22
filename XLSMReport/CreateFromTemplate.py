#import openpyxl
import win32com.client as wc
import pyodbc, re, shutil, os
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

def create(name):
    xl = wc.Dispatch('Excel.Application')
    xl.Visible = 1
    xl.DisplayAlerts = 0
    wb = xl.Workbooks.Open("C:\\Users\\jeremyshank\\Documents\\BMSS Assets\\Reports\\Fname Lname Weekly Hours Report.xlsm")
    writeData = wb.Worksheets('Staff')
    writeData.Cells(1,1).Value = name
    xl.Application.Run("RunAllQueries")
    if re.search("''", name) != None:
        name = re.sub("''", "'", name)
    wb.SaveAs("C:\\Users\\jeremyshank\\Desktop\\No Macro\\" + name + " Weekly Hours Report.xlsx", FileFormat=51)
    wb.Close(SaveChanges=False)
    xl.Quit()
    xl = None
    shutil.rmtree("C:\\Users\\jeremyshank\\AppData\\Local\\Temp\\gen_py", ignore_errors=True)
    os.system("taskkill /f /im excel.exe")

namerow = conn.cursor()
namerow.execute("""SELECT [StaffName]
  FROM [dbo].[tblStaff]
  WHERE StaffName = StaffEnded IS NULL AND
  StaffOffice IN ('BHM', 'HSV', 'GAD')""")

names = namerow.fetchall()

for name in names:
    if re.search("'", name[0]) != None:
        nick = re.sub("'", "''", name[0])
        create(nick)
    else:
        create(name[0])
