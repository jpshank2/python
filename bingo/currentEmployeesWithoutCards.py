import pyodbc, os
import win32com.client as wc
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

engineConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')
devopsConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))

allCardsList = devopsConn.cursor()
allCardsList.execute("""SELECT * FROM dbo.BingoCards WHERE BingoUser IS NOT NULL AND BingoCompany = 'BMSS'""")

allCards = allCardsList.fetchall()

cardsForSQL = str()

for x in range(len(allCards)):
    if x == len(allCards) - 1:
        cardsForSQL += str(allCards[x][1])
    else:
        cardsForSQL += str(allCards[x][1]) + ""","""

currentEmpsWithoutCardsList = engineConn.cursor()
currentEmpsWithoutCardsList.execute("""SELECT StaffIndex,
StaffName,
StaffEMail
FROM dbo.tblStaff
WHERE StaffIndex NOT IN (""" + cardsForSQL + """, -100, 0, 387, 114, 28016, 35023, 52, 271, 35288, 34483) AND StaffEnded IS NULL AND StaffType <> 4""")

currentEmpsWithoutCards = currentEmpsWithoutCardsList.fetchall()

for employee in currentEmpsWithoutCards:
    firstAvailableCard = devopsConn.cursor()
    firstAvailableCard.execute("SELECT TOP 1 BingoCard FROM dbo.BingoCards WHERE BingoUser IS NULL AND BingoCompany IS NULL")

    cardToAssign = firstAvailableCard.fetchall()[0][0]
    print(employee[1] + " - " + str(employee[0]) + ' ' + str(cardToAssign))
    addEmpToGame = devopsConn.cursor()
    addEmpToGame.execute("UPDATE dbo.BingoCards SET BingoUser = " + str(employee[0]) + ", BingoCompany = 'BMSS' WHERE BingoCard = " + str(cardToAssign))
    addEmpToGame.commit()

    path = "C:\\Users\\jeremyshank\\Documents\\Bingo\\Cards\\" + str(cardToAssign) + '.xlsx'

    outlook = wc.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.To = employee[2]
    mail.CC = 'bshealy@bmss.com'
    mail.Subject = 'Welcome to the Bingo Game!'
    mail.Attachments.Add(path)
    mail.HTMLBody = '<p>' + employee[1] + ',</p><p>Here is your newly assigned Bingo card. If you had a card previously and were removed from the game for any reason this will be your new permanent card. Rules for the Bingo game can be found <a href="http://zeal.bmss.com/kb/time-entry-and-bingo/">here</a> on Zeal.</p>'

    mail.Send()
