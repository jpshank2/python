import pyodbc, os
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

cwConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('CW_SERVER') + ';DATABASE=' + os.getenv('CW_DATABASE') + ';UID=' + os.getenv('CW_USER') + ';PWD=' + os.getenv('CW_PASS'))
devopsConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))

overrideUsersList = ['Jeremy Shank']

overrideUsersIntList = cwConn.cursor()

staffIndexForSQL = str()

for i in range(len(overrideUsersList)):
    overrideUsersIntList.execute("""SELECT Member_RecID, First_Name, Last_Name 
                                    FROM dbo.Member
                                    WHERE CONCAT(First_Name, ' ', Last_Name) = '""" + overrideUsersList[i] + """'""")
    overrideUsersInt = overrideUsersIntList.fetchall()
    if i == len(overrideUsersList) - 1:
        staffIndexForSQL += str(overrideUsersInt[0][0])
    else:
        staffIndexForSQL += str(overrideUsersInt[0][0]) + ", "

overrideBingo = devopsConn.cursor()

overrideBingo.execute("""UPDATE dbo.Bingo
SET BingoDate = CONVERT(DATE, DATEADD(DAY, -1, BingoDate))
WHERE BingoNumber = 0 
AND BingoCard IN (SELECT BingoCard FROM dbo.BingoCards WHERE BingoUser IN (""" + staffIndexForSQL + """) AND BingoCompany = 'ABIT')

UPDATE dbo.Bingo
SET BingoMissed = BingoMissed - 1
WHERE BingoCard IN (SELECT BingoCard FROM dbo.BingoCards WHERE BingoUser IN (""" + staffIndexForSQL + """)AND BingoCompany = 'ABIT')""")

overrideBingo.commit()
