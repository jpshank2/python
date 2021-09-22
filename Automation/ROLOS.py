import pyodbc, re, os, openpyxl
import win32com.client as wc
from dotenv import load_dotenv
from openpyxl.worksheet.table import Table, TableStyleInfo

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

engineConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS'))
devopsConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))

homeroomList = engineConn.cursor()
homeroomList.execute("""SELECT StaffName, StaffEMail
FROM dbo.tblStaff
WHERE StaffName IN (SELECT CatName
FROM dbo.tblCategory
WHERE CatName <> 'Unknown' AND CatType = 'SUBDEPT')""")

homerooms = homeroomList.fetchall()

for homeroom in homerooms:
    homeroomMemberList = engineConn.cursor()
    homeroomMemberList.execute("""Select
ts.StaffIndex,
ts.StaffName AS [Employee],
ts.StaffUser AS [Staff Email],
H.CatName AS Homeroom
FROM       
tblStaff ts 
inner join tblStaffEx SE ON ts.StaffIndex = SE.StaffIndex
inner Join tblCategory H ON H.Category = SE.StaffSubDepartment AND H.CatType = 'SUBDEPT'
WHERE H.CatName = '""" + homeroom[0] + """' AND TS.StaffEnded IS NULL AND TS.StaffType <> 4""")
    homeroomMembers = homeroomMemberList.fetchall()

    homeroomListForSQL = str()

    for i in range(len(homeroomMembers)):
        if i == len(homeroomMembers) - 1:
            homeroomListForSQL += """'""" + homeroomMembers[i][1] + """'"""
        else:
            homeroomListForSQL += """'""" + homeroomMembers[i][1] + """', """

    kudosAndRolosList = devopsConn.cursor()
    kudosAndRolosList.execute("""SELECT CAST([EventDate] AS date) AS [EventDate]
      ,[EventPerson]
      ,[EventAction]
      ,[EventNotes]
      ,[EventUpdatedBy]
  FROM [DataWarehouse].[dbo].[MandM]
  WHERE OwnerType IS NULL AND EventAction IN ('KUDOS', 'UPWARD', 'DOWNWARD') AND EventPerson IN (""" + homeroomListForSQL + """) AND EventDate < '2021-06-01'
  ORDER BY EventPerson""")
    
    kudosAndRolos = kudosAndRolosList.fetchall()

    wb = openpyxl.Workbook()
    ws = wb.active

    ws.append(["Date Submitted", "Recipient", "Type", "Notes", "Submitter"])

    endCell = "E" + str(len(kudosAndRolos) + 1)

    for kudosAndRolo in kudosAndRolos:
        ws.append(list(kudosAndRolo))
    
    table = Table(displayName="M_and_M_Table", ref="A1:" + endCell)

    tableStyle = TableStyleInfo(name="TableStyleMedium4", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)

    table.tableStyleInfo = tableStyle
    path = "C:\\Users\\jeremyshank\\Desktop\\M+M Reports\\" + homeroom[0] + " Homeroom Report.xlsx"
    ws.add_table(table)
    wb.save(path)

    outlook = wc.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.To = homeroom[1]
    mail.CC = 'hrussell@bmss.com'
    mail.Subject = 'Homeroom Summary for Evals'
    mail.Attachments.Add(path)
    mail.HTMLBody = '<p>' + homeroom[0] + ',</p><p>Attached is a summary report with all KUDOS and ROLOs given for your homeroom members</p>'

    mail.Send()