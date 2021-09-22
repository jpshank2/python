import re, pyodbc, os
import win32com.client as wc
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS'))

def writePY(i):
    name = i[0]
    email = i[1]
    username = re.sub("@bmss.com", "", i[1])
    pathName = "C:\\Users\\jeremyshank\\Desktop\\No Macro\\" + username + "Weekly.py"
    p = open(pathName, "w")
    p.write("import time")
    p.write("import win32com.client as wc\n\n")
    p.write("""path = r"C:\\AutomatedReports\\No Macro\\""" + name + """ Weekly Hours Report.xlsx"\n""")
    p.write("""xlapp = wc.DispatchEx("Excel.Application")\nwb = xlapp.Workbooks.Open(path)\nwb.RefreshAll()\n""")
    p.write("""xlapp.CalculateUntilAsyncQueriesDone()\nxlapp.DisplayAlerts = False\nwb.Save\nwb.Close\nxlapp.Quit()\nxlapp = None\n\n""")
    p.write("""outlook = wc.Dispatch("outlook.application")\nmail = outlook.CreateItem(0)\n""")
    p.write("""mail.To = \"""" + email + """\"\nmail.BCC = 'jeremyshank@bmss.com'\nmail.Subject = \"""" + name + """ Weekly Hours Report\"\n""")
    p.write("""mail.AttachmentAdd(path)\nmail.HTMLBody = "<p>See attached spreadsheet...</p>\"\n\n""")
    p.write("mail.Send()\n\nmail.Close\noutlook.Quit()\noutlook = None")

namerow = conn.cursor()
namerow.execute("""SELECT [StaffName],
  [StaffEMail]
  FROM [dbo].[tblStaff]
  WHERE StaffEnded IS NULL 
  AND StaffOffice IN ('BHM', 'HSV', 'GAD')""")
    
names = namerow.fetchall()

for name in names:
    writePY(name)
