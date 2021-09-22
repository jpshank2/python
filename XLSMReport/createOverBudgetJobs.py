import win32com.client as wc
import pyodbc, re, shutil, os
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

def create(name):
    xl = wc.Dispatch('Excel.Application')
    xl.Visible = 1
    xl.DisplayAlerts = 0
    wb = xl.Workbooks.Open("C:\\Users\\jeremyshank\\Documents\\BMSS Assets\\Reports\\OverBudgetJobsReport.xlsm")
    writeData = wb.Worksheets('Staff')
    writeData.Cells(1,1).Value = name[0]
    xl.Application.Run("DueDateQuery")
    wb.SaveAs("C:\\Users\\jeremyshank\\BMSS\\Business Intelligence - Documents\\Ad Hoc Reports\\Reports for PA\\" + name[1] + " Over Budget Jobs Report.xlsx", FileFormat=51)
    wb.Close(SaveChanges=False)
    xl.Quit()
    xl = None
    shutil.rmtree("C:\\Users\\jeremyshank\\AppData\\Local\\Temp\\gen_py", ignore_errors=True)
    os.system("taskkill /f /im excel.exe")

namerow = conn.cursor()
namerow.execute("""SELECT DISTINCT StaffIndex, StaffName
FROM dbo.tblStaff S
INNER JOIN dbo.tblJob_Header JH ON JH.Job_Manager = S.StaffIndex
WHERE S.StaffEnded IS NULL AND StaffIndex NOT IN (0, 22) 
ORDER BY StaffIndex
--OFFSET 80 ROWS""")

names = namerow.fetchall()

for name in names:
    create(name)
