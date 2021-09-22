from dotenv import load_dotenv
import win32com.client as wc
import pyodbc, re, shutil, os

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

def create(name):
    print("Starting " + name)
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
    print("Finished with " + name)

def email(user):
    print("Emailing " + user[0])
    path = "C:\\Users\\jeremyshank\\Desktop\\No Macro\\" + user[0] + " Weekly Hours Report.xlsx"
    outlook = wc.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.To = user[1]
    mail.Subject = user[0] + ' Weekly Hours Report'
    mail.Attachments.Add(path)
    mail.HTMLBody = '<p>See attached spreadsheet...</p><p>should go to ' + user[1] + '</p>'

    mail.Send()

def delete(file):
    print("Deleting " + file + "'s temp file")
    os.remove("C:\\Users\\jeremyshank\\Desktop\\No Macro\\" + file + " Weekly Hours Report.xlsx")

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

namerow = conn.cursor()
namerow.execute("""SELECT StaffName, StaffEMail
FROM dbo.tblStaff
WHERE StaffName = 'Caroline Gilmer'--StaffEnded IS NULL
--AND StaffOffice IN ('BHM', 'HSV', 'GAD')
--last ran on 3/8/2021
--AND StaffStarted > '1/24/2021'
ORDER BY StaffName""")

names = namerow.fetchall()

for name in names:
    if re.search("'", name[0]) != None:
        subName = re.sub("'", "''", name[0])
        create(subName)
        email(name)
        delete(name[0])
    else:
        create(name[0])
        email(name)
        delete(name[0])
