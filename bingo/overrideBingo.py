import pyodbc, os
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

engineConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')
devopsConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))

overrideUsersList = ['Meredith Clayton', 'Jennifer Stedham', 'Ashley Caldwell', 'Auston Sullivan', 'Becky Shealy', 'Dalton Bradshaw', 'Theresa Gilley', 'Jennifer Stedham', 'Ann Marie Tanielu', 'Hannah Piper', 'Jonathan Hall']

overrideUsersIntList = engineConn.cursor()

staffIndexForSQL = str()

for i in range(len(overrideUsersList)):
    overrideUsersIntList.execute("""SELECT StaffIndex FROM dbo.tblStaff WHERE StaffName = '""" + overrideUsersList[i] + """'""")
    overrideUsersInt = overrideUsersIntList.fetchall()
    if i == len(overrideUsersList) - 1:
        staffIndexForSQL += str(overrideUsersInt[0][0])
    else:
        staffIndexForSQL += str(overrideUsersInt[0][0]) + ", "

overrideBingo = devopsConn.cursor()

overrideBingo.execute("""UPDATE dbo.Bingo
SET BingoDate = CONVERT(DATE, DATEADD(DAY, -1, BingoDate))
WHERE BingoNumber = 0 
AND BingoCard IN (SELECT BingoCard FROM dbo.BingoCards WHERE BingoUser IN (""" + staffIndexForSQL + """) AND BingoCompany = 'BMSS')

UPDATE dbo.Bingo
SET BingoMissed = BingoMissed - 1
WHERE BingoCard IN (SELECT BingoCard FROM dbo.BingoCards WHERE BingoUser IN (""" + staffIndexForSQL + """) AND BingoCompany = 'BMSS')""")

overrideBingo.commit()
