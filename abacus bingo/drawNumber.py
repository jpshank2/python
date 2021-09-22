import pyodbc, os, random
import win32com.client as wc
from datetime import date
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

cwConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('CW_SERVER') + ';DATABASE=' + os.getenv('CW_DATABASE') + ';UID=' + os.getenv('CW_USER') + ';PWD=' + os.getenv('CW_PASS'))
devopsConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))

bingoNumbersList = devopsConn.cursor()
bingoNumbersList.execute("""SELECT DISTINCT BingoNumber
FROM dbo.Bingo B
INNER JOIN dbo.BingoCards BC ON BC.BingoCard = B.BingoCard
WHERE B.BingoCalled = 0 AND BingoCompany = 'ABIT'""")

bingoNumbers = bingoNumbersList.fetchall()

x = random.randint(0, len(bingoNumbers))
bingoNumber = bingoNumbers[x][0]

letter = str()

if bingoNumber > 0 and bingoNumber < 16:
    letter = 'B'
elif bingoNumber > 15 and bingoNumber < 31:
    letter = 'I'
elif bingoNumber > 30 and bingoNumber < 46:
    letter = 'N'
elif bingoNumber > 45 and bingoNumber < 61:
    letter = 'G'
else:
    letter = 'O'

updateBingoNumbers = devopsConn.cursor()
updateBingoNumbers.execute("""UPDATE dbo.Bingo
SET BingoCalled = 1, BingoDate = GETDATE()
WHERE BingoNumber = """ + str(bingoNumber) + """
AND BingoCard IN (SELECT BingoCard
    FROM dbo.BingoCards
    WHERE BingoCompany = 'ABIT')""")

updateBingoNumbers.commit()

calledBingoUsersList = devopsConn.cursor()
calledBingoUsersList.execute("""SELECT BingoUser FROM dbo.Bingo B INNER JOIN dbo.BingoCards BC ON BC.BingoCard = B.BingoCard WHERE BC.BingoCompany = 'ABIT' AND B.BingoNumber = """ + str(bingoNumber))

calledBingoUsers = calledBingoUsersList.fetchall()

for user in calledBingoUsers:
    calledBingoInfoList = cwConn.cursor()
    calledBingoInfoList.execute("""SELECT Member_RecID, CONCAT(First_Name, ' ', Last_Name) AS Name, Email_Address 
    FROM dbo.Member
    WHERE Member_RecID = """ + str(user[0]))

    calledBingoInfo = calledBingoInfoList.fetchall()

    outlook = wc.Dispatch('outlook.application')
    send_account = None
    for account in outlook.Session.Accounts:
        if account.DisplayName == 'jshank@abacustechnologies.com':
            send_account = account
            break

    mail = outlook.CreateItem(0)
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))

    mail.To = calledBingoInfo[0][2]
    mail.Subject = 'Bingo Draw ' + letter + ' ' + str(bingoNumber) + ' - ' + str(date.today())
    mail.HTMLBody = '<p>You had ' + letter + ' ' + str(bingoNumber) + ' on your Bingo Card! Remember to enter your time for yesterday by noon today!</p>'
    
    mail.Send()

notCalledBingoUsersList = devopsConn.cursor()
notCalledBingoUsersList.execute("""SELECT DISTINCT BingoUser
FROM dbo.Bingo B
INNER JOIN dbo.BingoCards BC ON BC.BingoCard = B.BingoCard
WHERE BC.BingoCompany = 'ABIT' AND BingoUser NOT IN (
	SELECT BingoUser 
	FROM dbo.Bingo B 
		INNER JOIN dbo.BingoCards BC ON BC.BingoCard = B.BingoCard 
	WHERE BC.BingoCompany = 'ABIT' AND B.BingoNumber = """ + str(bingoNumber) + """)""")

notCalledBingoUsers = notCalledBingoUsersList.fetchall()

for user in notCalledBingoUsers:
    notCalledBingoInfoList = cwConn.cursor()
    notCalledBingoInfoList.execute("""SELECT Member_RecID, CONCAT(First_Name, ' ', Last_Name) AS Name, Email_Address 
    FROM dbo.Member
    WHERE Member_RecID = """ + str(user[0]))

    notCalledBingoInfo = notCalledBingoInfoList.fetchall()

    outlook = wc.Dispatch('outlook.application')
    send_account = None
    for account in outlook.Session.Accounts:
        if account.DisplayName == 'jshank@abacustechnologies.com':
            send_account = account
            break

    mail = outlook.CreateItem(0)
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))

    mail.To = notCalledBingoInfo[0][2]
    mail.Subject = 'Bingo Draw ' + letter + ' ' + str(bingoNumber) + ' - ' + str(date.today())
    mail.HTMLBody = '<p>You did not have ' + letter + ' ' + str(bingoNumber) + ' on your Bingo Card. Remember to enter your time for yesterday by noon today and better luck tomorrow!</p>'
    
    mail.Send()
