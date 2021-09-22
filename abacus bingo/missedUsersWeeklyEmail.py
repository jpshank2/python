import win32com.client as wc
import pyodbc, os
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

cwConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('CW_SERVER') + ';DATABASE=' + os.getenv('CW_DATABASE') + ';UID=' + os.getenv('CW_USER') + ';PWD=' + os.getenv('CW_PASS'))
devopsConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))

missedStaffList = devopsConn.cursor()
missedStaffList.execute("""SELECT DISTINCT BC.BingoUser, B.BingoMissed
FROM dbo.BingoCards BC
	INNER JOIN dbo.Bingo B ON B.BingoCard = BC.BingoCard AND BC.BingoCompany = 'ABIT'
WHERE B.BingoMissed <> 0
ORDER BY BC.BingoUser""")

missedStaff = missedStaffList.fetchall()
staffIndexForSQL = str()

for i in range(len(missedStaff)):
	if i == len(missedStaff) - 1:
		staffIndexForSQL += str(missedStaff[i][0])
	else:
		staffIndexForSQL += str(missedStaff[i][0]) + ', '

staffNamesList = cwConn.cursor()
staffNamesList.execute("""SELECT CONCAT(First_Name, ' ', Last_Name) FROM dbo.Member WHERE Member_RecID IN (""" + staffIndexForSQL + """) ORDER BY Member_RecID""")

staffNames = staffNamesList.fetchall()

htmlStaffNames = str()

for i in range(len(staffNames)):
    if missedStaff[i][1] == 1:
        htmlStaffNames += '<li>' + staffNames[i][0] + ' has missed the time entry cutoff 1 time</li>'
    else:
        htmlStaffNames += '<li>' + staffNames[i][0] + ' has missed the time entry cutoff ' + str(missedStaff[i][1]) + ' times</li>'

outlook = wc.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

send_account = None
for account in outlook.Session.Accounts:
    if account.DisplayName == 'jshank@abacustechnologies.com':
        send_account = account
        break

mail._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))

mail.To = 'jfinn@abacustechnologies.com'
mail.Subject = 'Bingo Missed Users'
mail.HTMLBody = '<p>The following users are out of the current bingo game:</p><ul>' + htmlStaffNames + '</ul>'

mail.Send()