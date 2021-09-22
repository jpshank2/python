import win32com.client as wc
import pyodbc
import re
import shutil
import os


def create(name):
    xl = wc.Dispatch('Excel.Application')
    xl.Visible = 1
    xl.DisplayAlerts = 0
    wb = xl.Workbooks.Open(
        "C:\\Users\\jeremyshank\\Desktop\\Client Update Project.xlsm")
    writeData = wb.Worksheets('Staff')
    writeData.Cells(1, 1).Value = name
    xl.Application.Run("GetOriginators")
    if re.search("''", name) != None:
        name = re.sub("''", "'", name)
    wb.SaveAs("C:\\Users\\jeremyshank\\Desktop\\No Macro\\" +
              name + " Client Update Project.xlsx", FileFormat=51)
    wb.Close(SaveChanges=False)
    xl.Quit()
    xl = None
    shutil.rmtree(
        "C:\\Users\\jeremyshank\\AppData\\Local\\Temp\\gen_py", ignore_errors=True)
    os.system("taskkill /f /im excel.exe")


conn = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

namerow = conn.cursor()
namerow.execute("""select 
	T.Originator,
	S.StaffEmail
From
(select
	(CASE WHEN O.StaffCategory='1' THEN O.StaffName ELSE (CASE WHEN O.STAFFNAME='Bill House' THEN 'Tom Albright' WHEN O.StaffName='Randy Barnes' THEN 'Greg Barnes' ELSE P.STAFFNAME END) END) as [Originator]

from tblEngagement E
	Left Join tblCategory I on E.ClientIndustry = I.Category and I.CatType = 'INDUSTRY'
	Inner Join tblClientOrigination OWN on E.Contindex=OWN.ContIndex
	Inner Join tblStaff O on OWN.StaffIndex=O.StaffIndex
	Left Join tblOwnerType Ent on E.ClientOwnership = Ent.OwnerIndex
	Inner Join tblStaff P on E.ClientPartner = P.StaffIndex
	Inner Join tblStaff M on E.ClientManager = M.StaffIndex
where E.ClientStatus NOT IN ('Lost','Internal','Suspended') and E.ClientCode<>'RLBTest' and (CASE WHEN O.StaffCategory='1' THEN O.StaffName ELSE P.STAFFNAME END) <> 'Administration'
group by (CASE WHEN O.StaffCategory='1' THEN O.StaffName ELSE (CASE WHEN O.STAFFNAME='Bill House' THEN 'Tom Albright' WHEN O.StaffName='Randy Barnes' THEN 'Greg Barnes' ELSE P.STAFFNAME END) END)
) T
Inner Join tblStaff S ON T.Originator=S.StaffName
order by T.Originator""")

names = namerow.fetchall()

for name in names:
    if re.search("'", name[0]) != None:
        nick = re.sub("'", "''", name[0])
        create(nick)
    else:
        create(name[0])
