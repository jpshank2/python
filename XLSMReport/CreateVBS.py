import pyodbc, re, os
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

def writeVBS(i):
    name = i[0]
    email = i[1]
    username = re.sub("@bmss.com", "", i[1])
    pathName = "C:\\Users\\jeremyshank\\Desktop\\No Macro\\VBS\\" + username + "Weekly.vbs"
    p = open(pathName, "w")
    p.write("""Set objFSO = CreateObject("Scripting.FileSystemObject")\n\n""")
    p.write("""src_file = objFSO.GetAbsolutePathName("C:\\AutomatedReports\\No Macro\\""" + name + """ Weekly Hours Report.xlsx")\n\n""")
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
	.subject = \""""+ name + """ Weekly Hours Report"
	.Attachments.Add src_file
	.HTMLBody = "See attached spreadsheet..."
	.Send
wscript.sleep 100*100
End with


wscript.quit""")
    p.close()

namerow = conn.cursor()
namerow.execute("""SELECT [StaffName],
  [StaffEMail]
  FROM [dbo].[tblStaff]
  WHERE StaffName = StaffEnded IS NULL AND
  --StaffOffice IN ('BHM', 'HSV', 'GAD')
  --last ran on 3/8/2021
  --AND StaffStarted > '1/24/2021'""")
    
names = namerow.fetchall()

for name in names:
    writeVBS(name)
