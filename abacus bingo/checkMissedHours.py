import win32com.client as wc
import pyodbc, os
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

cwConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('CW_SERVER') + ';DATABASE=' + os.getenv('CW_DATABASE') + ';UID=' + os.getenv('CW_USER') + ';PWD=' + os.getenv('CW_PASS'))
devopsConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))

def email(staff):
    bingoMissedIntList = devopsConn.cursor()
    bingoMissedIntList.execute("""SELECT BingoMissed 
        FROM dbo.Bingo B 
        INNER JOIN dbo.BingoCards BC ON BC.BingoCard = B.BingoCard 
        WHERE BC.BingoCompany = 'ABIT' AND BC.BingoUser = """ + str(staff[0]))
    
    bingoMissedInt = bingoMissedIntList.fetchall()[0][0]

    outlook = wc.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    if bingoMissedInt == 1:
        subject = 'Current Bingo Game Missed First Violation'
        body = """<p>""" + staff[1] + """,</p><p>You have missed entering your time for yesterday by our cutoff and have been kicked out of this Bingo game. Keep entering your time to avoid further penalty!</p>"""
    elif bingoMissedInt == 2:
        subject = 'Current Bingo Game Missed Second Violation'
        body = """<p>""" + staff[1] + """,</p><p>You have missed entering your time for yesterday by our cutoff and have contributed $10 to the pot from you Christmas bonus. Keep entering your time to avoid further penalty!</p>"""
    elif bingoMissedInt == 3:
        subject = 'Current Bingo Game Missed Third Violation'
        body = """<p>""" + staff[1] + """,</p><p>You have missed entering your time for yesterday by our cutoff and have contributed $25 to the pot from you Christmas bonus. Keep entering your time to avoid further penalty!</p>"""
    else:
        mail.CC = 'bjackson@abacustechnologies.com'
        subject = 'Current Bingo Game Missed Fourth or More Violation'
        body = """<p>""" + staff[1] + """,</p><p>You have missed entering your time for yesterday by our cutoff more than three times. Please schedule a time to meet with Brian to talk about your time entry habits.</p>"""

    send_account = None
    for account in outlook.Session.Accounts:
        if account.DisplayName == 'jshank@abacustechnologies.com':
            send_account = account
            break
    
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))

    mail.To = staff[2]
    mail.Subject = subject
    mail.HTMLBody = body

    mail.Send()

bingoStaffList = devopsConn.cursor()
bingoStaffList.execute("""SELECT BingoUser FROM dbo.BingoCards WHERE BingoCompany = 'ABIT'""")

bingoStaff = bingoStaffList.fetchall()

staffIndexForSQL = str()

for i in range(len(bingoStaff)):
	if i == len(bingoStaff) - 1:
		staffIndexForSQL += str(bingoStaff[i][0])
	else:
		staffIndexForSQL += str(bingoStaff[i][0]) + ', '

namerow = cwConn.cursor()
namerow.execute("""SELECT M.Member_RecID, CONCAT(M.First_Name, ' ', M.Last_Name) AS [Name], M.Email_Address, SUM(COALESCE(TE.Hours_Actual, 0)) AS [Hours_Actual], M.Hours_Min FROM dbo.Time_Entry TE
INNER JOIN dbo.Member M ON M.Member_ID = TE.Member_ID
WHERE CONVERT(DATE, TE.Date_Start) = CONVERT(DATE, DATEADD(DAY, -1, GETDATE())) AND M.Member_RecID IN (""" + staffIndexForSQL + """)
GROUP BY M.Member_RecID, M.Email_Address, M.Hours_Min, M.First_Name, M.Last_Name
HAVING SUM(TE.Hours_Actual) < 7""")

names = namerow.fetchall()

for i in range(0, len(names)):
    bingoCards = devopsConn.cursor()
    bingoCards.execute("""DECLARE @card int
SET @card = (SELECT BingoCard FROM dbo.BingoCards WHERE BingoUser = """ + str(names[i][0]) + """ AND BingoCompany = 'ABIT')

DECLARE @missed int
SET @missed = (SELECT BingoMissed FROM dbo.Bingo WHERE BingoCard = @card AND BingoNumber = 0)

UPDATE dbo.Bingo
SET BingoMissed = (@missed + 1)
WHERE BingoCard = @card

UPDATE dbo.Bingo
SET BingoDate = CONVERT(DATE, GETDATE())
WHERE BingoCard = @card AND BingoNumber = 0""")
    devopsConn.commit()

    email(names[i])