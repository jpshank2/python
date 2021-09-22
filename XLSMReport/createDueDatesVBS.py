import pyodbc, re, os
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')


def writeVBS(i):
    name = i[0]
    email = i[1]
    username = re.sub("@bmss.com", "", i[1])
    pathName = "C:\\Users\\jeremyshank\\Desktop\\No Macro\\" + username + "DueDates.vbs"
    p = open(pathName, "w")
    p.write("""Set objFSO = CreateObject("Scripting.FileSystemObject")\n\n""")
    p.write("""src_file = objFSO.GetAbsolutePathName("C:\\AutomatedReports\\Due Dates\\""" + name + """ Due Dates Report.xlsx")\n\n""")
    p.write("""Dim oExcel
Set oExcel = CreateObject("Excel.Application")

Dim oBook
Set oBook = oExcel.Workbooks.Open(src_file)
oExcel.DisplayAlerts = False
oExcel.AskToUpdateLinks = False
oExcel.AlertBeforeOverwriting = False


oBook.RefreshAll

wscript.sleep 100*100

oBook.Save

oBook.Close (False)
oExcel.Quit

'Dim outlook, email
Set outlook = CreateObject("Outlook.Application")
Set email = outlook.CreateItem(0)\n\n""")
    p.write("""with email
	.to = \"""" + email + """\"
	.bcc = "jeremyshank@bmss.com"
	.subject = \""""+ name + """ Due Dates Report"
	.Attachments.Add src_file
	.HTMLBody = "<p>See attached...</p>"
	.Send
wscript.sleep 100*100
End with


wscript.quit""")
    p.close()

namerow = conn.cursor()
namerow.execute("""Select			--P.PersonTitle AS Title
				ts.StaffName AS Employee
                ,ts.StaffEmail
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
    writeVBS(name)
