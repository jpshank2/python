import pyodbc, re, os
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

def writeVBS(i):
    name = i[0]
    email = i[1]
    username = re.sub("@bmss.com", "", i[1])
    pathName = "C:\\Users\\jeremyshank\\Desktop\\test\\" + username + "MarchDueDates.vbs"
    p = open(pathName, "w")
    p.write("""Set objFSO = CreateObject("Scripting.FileSystemObject")\n\n""")
    p.write("""src_file = objFSO.GetAbsolutePathName("C:\\AutomatedReports\\Due Dates\\""" + name + """ March Due Dates Report.xlsx")\n\n""")
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
	.subject = \""""+ name + """ March Due Dates Report"
	.Attachments.Add src_file
	.HTMLBody = "<h1>New March Due Dates Report</h1><br><h2 style=" + Chr(34) + "color:#366092" + Chr(34) + ">What to Expect</h2><p>When you open this report, you will recognize it as being very similar to the Due Dates report you receive on Monday mornings, however this report is only showing 2019 1065 Partnership Returns, 2019 1120S S-Corporation Returns, and 2019-A Foreign Trusts that have a 3/15 Deadline. This report is highlighting projects where the Job Status is " + Chr(34) + "Not Started" + Chr(34) + " yellow, and jobs where the Job Status is " + Chr(34) + "In Progress" + Chr(34) + " but the <i>Workflow Status</i> is " + Chr(34) + "Not Started" + Chr(34) + " orange. If these are jobs BMSS is no longer engaged to perform, engagement leaders should mark those jobs " + Chr(34) + "closed" + Chr(34) + " or " + Chr(34) + "completed" + Chr(34) + " in practice engine.</p><p>Refer to the Currently Applied Filters section below to see what is being presented, and how to customize it to suit your needs.</p><p><b>If you do not have any rows of data, you do not have any of these jobs due by 3/15.</b></p><br><h2 style=" + Chr(34) + "color:#366092" + Chr(34) + ">Currently Applied Filters</h2><ol><li>This report is filtered to show Client Projects that are " + Chr(34) + "2019 1065 Partnership Return," + Chr(34) + " " + Chr(34) + "2019 1120S S-Corporation Return," + Chr(34) + " or " + Chr(34) + "2019 3520-A Foreign Trust." + Chr(34) + "</li><li>This report is filtered to show Client Projects with a due date that is equal to 3/15/2020. If you would like to see all projects, remove the filter from the Due Before 3/15/20 column. Alternatively, if you are only interested in the projects due on 3/15/2020, you can select that date as a filter on The Due Date column.</li><li>This report is sorted to show the oldest due dates at the top.</li><li>This report is highlighting any Client Project where the Job Status is " + Chr(34) + "Not Started" + Chr(34) + "</li></ol><br><h2 style=" + Chr(34) + "color:#366092" + Chr(34) + ">How is This Report Different</h2><p>As stated above, this report is showing a slice of currently open jobs, and is highlighting jobs that are not started. This is intended to help engagement leaders identify what is in PE that needs to be started, and what is in PE that BMSS is no longer working on. When you mark those jobs " + Chr(34) + "in progress" + Chr(34) + " or " + Chr(34) + "closed" + Chr(34) + ", you can refresh this report by going to the Data ribbon at the top, and selecting the Refresh All button in the Queries & Connections section.</p>"
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
