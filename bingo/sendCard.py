import pyodbc, os
import win32com.client as wc
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

engineConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS'))
devopsConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))

allCardsList = devopsConn.cursor()
allCardsList.execute("""SELECT * FROM dbo.BingoCards WHERE BingoUser IS NOT NULL AND BingoCompany = 'BMSS' AND BingoUser < 78 AND BingoUser > 40 ORDER BY BingoUser""")

allCards = allCardsList.fetchall()

cardsForSQL = str()

for x in range(len(allCards)):
    if x == len(allCards) - 1:
        cardsForSQL += str(allCards[x][1])
    else:
        cardsForSQL += str(allCards[x][1]) + ""","""

currentEmpsWithCardsList = engineConn.cursor()
currentEmpsWithCardsList.execute("""SELECT StaffIndex,
StaffName,
StaffEMail
FROM dbo.tblStaff
WHERE StaffIndex IN (""" + cardsForSQL + """)
ORDER BY StaffIndex""")

currentEmpsWithCards = currentEmpsWithCardsList.fetchall()

for x in range(len(currentEmpsWithCards)):
    # print(currentEmpsWithCards[x])
    # path = "C:\\Users\\jeremyshank\\Documents\\Bingo\\Cards\\" + str(allCards[x][0]) + '.xlsx'

    outlook = wc.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = currentEmpsWithCards[x][2]
    mail.Subject = 'Two Bingo Cards'
    mail.HTMLBody = '<p>' + currentEmpsWithCards[x][1] + ',</p><p>You received 2 bingo cards. In trying to write the assigning script too quickly I made an error as assigned you a card equalling your Staff Index (' + str(currentEmpsWithCards[x][0]) + ') instead of your actual card. The second bingo card you received (' + str(allCards[x][0]) + ') is the correct card.</p>'

    mail.Send()
