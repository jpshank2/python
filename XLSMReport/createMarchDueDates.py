
import win32com.client as wc
import pyodbc, re, shutil, os
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')


def create(name):
    xl = wc.Dispatch('Excel.Application')
    xl.Visible = 1
    xl.DisplayAlerts = 0
    wb = xl.Workbooks.Open("C:\\Users\\jeremyshank\\Documents\\BMSS Assets\\Reports\\MarchDeadline.xlsm")
    writeData = wb.Worksheets('Staff')
    writeData.Cells(1,1).Value = name
    xl.Application.Run("DueDateSlice")
    if re.search("''", name) != None:
        name = re.sub("''", "'", name)
    wb.SaveAs("C:\\Users\\jeremyshank\\Desktop\\No Macro\\" + name + " March Due Dates Report.xlsx", FileFormat=51)
    wb.Close(SaveChanges=False)
    xl.Quit()
    xl = None
    shutil.rmtree("C:\\Users\\jeremyshank\\AppData\\Local\\Temp\\gen_py", ignore_errors=True)
    os.system("taskkill /f /im excel.exe")

namerow = conn.cursor()
namerow.execute("""Select			--P.PersonTitle AS Title
				ts.StaffName AS Employee
				,D.DeptName as Department
				,G.GradeDesc AS Level
				,C.CatName AS Homeroom
				,Loc.OfficeName AS Office
				--,Convert(VARCHAR,TS.StaffStarted,101) as [Start Date]
				--,Convert(VARCHAR,ts.StaffEnded,101) AS [End Date]
				
	FROM       tblStaff ts 
	inner join tblGrade G ON G.GradeCode = ts.StaffCategory
	INNER JOIN tblDepartment AS D ON TS.StaffDepartment = D.DeptIdx
	inner join tblOffices AS Loc ON ts.StaFFOffice = Loc.OfficeCode
	inner join tblContactAttributes AS CA ON ts.StaffIndex = CA.ContIndex
	inner join tblContacts AS TC ON ts.StaffIndex = TC.ContIndex
	inner join tblPerson AS P ON ts.StaffIndex = P.ContIndex
	inner Join tblCategory C ON C.Category = CA.AttrValid AND C.CatType = 'HOMEROOM'


WHERE Loc.OfficeName <> ' No Selection' and TS.StaffEnded is Null AND TS.StaffManager = '-1'""")

names = namerow.fetchall()


for name in names:
    if re.search("'", name[0]) != None:
        nick = re.sub("'", "''", name[0])
        create(nick)
    else:
        create(name[0])
