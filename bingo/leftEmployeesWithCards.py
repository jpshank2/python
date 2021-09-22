import pyodbc, os
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

engineConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS'))
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

leftEmpsWithCardsList = engineConn.cursor()
leftEmpsWithCardsList.execute("""SELECT StaffIndex, 
StaffName, 
StaffEMail,
StaffEnded 
FROM dbo.tblStaff 
WHERE StaffIndex IN (""" + cardsForSQL + """) AND StaffEnded IS NOT NULL""")

leftEmpsWithCards = leftEmpsWithCardsList.fetchall()

for employee in leftEmpsWithCards:
    # removeEmpFromGame = devopsConn.cursor()
    # removeEmpFromGame.execute("""UPDATE dbo.BingoCards
    # SET BingoUser = NULL, BingoCompany = NULL
    # WHERE BingoUser = """ + str(employee[0]) + """
    
    # UPDATE dbo.Bingo
    # SET BingoMissed = 0, BingoPerfect = 0, BingoMissedTwice = 0, BingoMissedThrice = 0, BingoDate = null, BingoCalled = 0
    # WHERE BingoNumber <> 0 AND BingoCard = (SELECT BingoCard FROM dbo.BingoCards WHERE BingoUser = """ + str(employee[0]) + """)

    # UPDATE dbo.Bingo
    # SET BingoMissed = 0, BingoPerfect = 0, BingoMissedTwice = 0, BingoMissedThrice = 0, BingoDate = null, BingoCalled = 1
    # WHERE BingoNumber = 0 AND BingoCard = (SELECT BingoCard FROM dbo.BingoCards WHERE BingoUser = """ + str(employee[0]) + """)
    # """)

    # removeEmpFromGame.commit()
    print(employee)