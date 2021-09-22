import pyodbc, re, os
import win32com.client as wc
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

getStaff = conn.cursor()
getStaff.execute("""SELECT DISTINCT
	M.StaffName AS [Manager]
	,M.StaffEMail AS [Email]
FROM
	tblTranWIP TW
	INNER JOIN tblEngagement E ON TW.ContIndex = E.ContIndex
	INNER JOIN tblStaff P ON E.ClientPartner = P.StaffIndex
	INNER JOIN tblStaff M ON E.ClientManager = M.StaffIndex
WHERE E.ContIndex < 900000 AND TW.WIPOutstanding <> 0  AND M.StaffName <> 'No Staff Allocated'
GROUP BY
    M.StaffName, M.StaffEMail
HAVING
	SUM(CASE WHEN DATEDIFF(DAY,TW.WIPDate,GETDATE()) > 180 THEN TW.WIPOutstanding ELSE 0 END) <> 0""")

staff = getStaff.fetchall()

for i in range(0, 2):
    name = staff[i][0]

    if re.search("'", name) != None:
        name = re.sub("'", "''", name)

    getClients = conn.cursor()
    getClients.execute("""SELECT
        E.ClientCode AS [Code],
        E.ClientName AS [Client],
        P.StaffName AS [Partner],
        M.StaffName AS [Manager],
        SUM(TW.WIPOutstanding) AS [OS_WIP],
        SUM(CASE WHEN DATEDIFF(DAY,TW.WIPDate,GETDATE()) > 180 THEN TW.WIPOutstanding ELSE 0 END) AS [OVER_180]
    FROM
        tblTranWIP TW
        INNER JOIN tblEngagement E ON TW.ContIndex = E.ContIndex
        INNER JOIN tblStaff P ON E.ClientPartner = P.StaffIndex
        INNER JOIN tblStaff M ON E.ClientManager = M.StaffIndex
    WHERE E.ContIndex < 900000 AND TW.WIPOutstanding <> 0 AND M.StaffName = '""" + name + """'
    GROUP BY
        E.ClientCode,E.ClientName,P.StaffName,M.StaffName
    HAVING
        SUM(CASE WHEN DATEDIFF(DAY,TW.WIPDate,GETDATE()) > 180 THEN TW.WIPOutstanding ELSE 0 END) <> 0
    ORDER BY
        E.ClientCode""")

    clients = getClients.fetchall()

    clientsList = ""

    for j in range(0, len(clients)):
        clientsList += "<tr><td>" + str(clients[j][0]) + "</td><td>" + clients[j][1] + "</td><td>" + clients[j][2] + "</td><td>" + clients[j][3] + "</td><td>" + str(clients[j][4])[0:-2] + "</td><td>" + str(clients[j][5])[0:-2] + "</td></tr>"

    # outlook = wc.Dispatch('outlook.application')
    # mail = outlook.CreateItem(0)

    # mail.To = staff[i][1]
    # mail.Subject = 'WIP Cleanup Required'
    # mail.HTMLBody = '<p>' + staff[i][0] + ',</p><p>Below is a listing of clients that you are manager on with outstanding WIP over 180 days old. Leadership team is evaluating our clients as of 10/31 and all old WIP needs to be cleaned up prior to that date. If you have a special situation that warrants this old WIP to be carried forward, please email <a href="mailto:ar@bmss.com">AR@bmss.com</a> and Scott Garrison with an explanation of why an exception should be made. Please ensure that this WIP is cleared as of 10/31 or that an exception has been approved. Invoices to clear this WIP should be dated 10/31 to ensure they post in the October period. This clean up project needs to be completed by Friday 11/6.</p><p>Thank you!</p><table><thead><tr><th>Client Code</th><th>Client Name</th><th>Partner</th><th>Manager</th><th>WIP</th><th>Over 180</th></tr></thead><tbody>' + clientsList + '<tbody></table>'

    # mail.Send()

    # print(staff[i][0] + ": Below is a listing of clients that you are manager on that have outstanding WIP over 180 days old.  Leadership team is evaluating our clients as of 10/31 and all WIP needs to be cleaned up prior to that date.  If you have a special situation that warrants this old WIP to be carried forward, please email AR@bmss.com and Scott Garrison with an explanation of why an exception should be made.  Please ensure that this WIP is cleared or an exception approved by 10/31.  Thank you!\n" + clientsList)
