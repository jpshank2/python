import pyodbc, os
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

engineConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS'))
devopsConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))

removeUsersList = ['Betsy Nolen', 'Abigail Waddell', 'Katherine White']

removeUsersIntList = engineConn.cursor()

staffIndexForSQL = str()

for i in range(len(removeUsersList)):
    removeUsersIntList.execute("""SELECT StaffIndex FROM dbo.tblStaff WHERE StaffName = '""" + removeUsersList[i] + """'""")
    removeUsersInt = removeUsersIntList.fetchall()
    if i == len(removeUsersList) - 1:
        staffIndexForSQL += str(removeUsersInt[0][0])
    else:
        staffIndexForSQL += str(removeUsersInt[0][0]) + ", "

removeBingo = devopsConn.cursor()

removeBingo.execute("""UPDATE dbo.BingoCards
SET BingoUser = null, BingoCompany = null
WHERE BingoCard IN (SELECT BingoCard FROM dbo.BingoCards WHERE BingoUser IN (""" + staffIndexForSQL + """))

UPDATE dbo.Bingo
SET BingoCalled = 0, BingoMissed = 0, BingoDate = NULL, BingoPerfect = 0, BingoMissedTwice = 0, BingoMissedThrice = 0
WHERE BingoCard IN (SELECT BingoCard FROM dbo.BingoCards WHERE BingoUser IN (""" + staffIndexForSQL + """)) AND BingoNumber <> 0

UPDATE dbo.Bingo
SET BingoMissed = 0, BingoDate = NULL, BingoPerfect = 0, BingoMissedTwice = 0, BingoMissedThrice = 0
WHERE BingoCard IN (SELECT BingoCard FROM dbo.BingoCards WHERE BingoUser IN (""" + staffIndexForSQL + """)) AND BingoNumber = 0
""")

removeBingo.commit()