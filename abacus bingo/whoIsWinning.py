import pyodbc, os
import win32com.client as wc
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

cwConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('CW_SERVER') + ';DATABASE=' + os.getenv('CW_DATABASE') + ';UID=' + os.getenv('CW_USER') + ';PWD=' + os.getenv('CW_PASS'))
devopsConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))

staffMaster = cwConn.cursor()
bingoMasterList = devopsConn.cursor()
bingoMasterList.execute("""SELECT BC.BingoUser, COUNT(*)
FROM dbo.BingoCards BC
	INNER JOIN dbo.Bingo B ON B.BingoCard = BC.BingoCard AND BC.BingoCompany = 'ABIT'
WHERE B.BingoMissed < 1 AND B.BingoCalled = 1
GROUP BY BC.BingoUser
ORDER BY 2 DESC""")

bingoUserList = bingoMasterList.fetchall()

for bingoUser in bingoUserList:
    staffMaster.execute("""SELECT First_Name, Last_Name
    FROM dbo.Member
    WHERE Member_RecID = """ + str(bingoUser[0]))
    staff = staffMaster.fetchall()
    print(staff[0][0] + ' ' + staff[0][1] + ' has ' + str(bingoUser[1]) + ' hits on the bingo game')
