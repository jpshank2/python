import random, pyodbc, os
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))


def createBoard(index):
    numbers = list()
    # bingoInfoList = list()

    for i in range(1, 26):
        #B column
        if i < 6:
            randomNumber = random.randint(1,15)
            while randomNumber in numbers:
                randomNumber = random.randint(1,15)
            numbers.append(randomNumber)
        #I column
        elif i < 11 and i > 5:
            randomNumber = random.randint(16,30)
            while randomNumber in numbers:
                randomNumber = random.randint(16,30)
            numbers.append(randomNumber)
        #N column
        elif i < 13 and i > 10:
            randomNumber = random.randint(31,45)
            while randomNumber in numbers:
                randomNumber = random.randint(31,45)
            numbers.append(randomNumber)
        #Free space
        elif i == 13:
            numbers.append(0)
        #N column continued
        elif i < 16 and i > 13:
            randomNumber = random.randint(31,45)
            while randomNumber in numbers:
                randomNumber = random.randint(31,45)
            numbers.append(randomNumber)
        #G column
        elif i < 21 and i > 15:
            randomNumber = random.randint(46,60)
            while randomNumber in numbers:
                randomNumber = random.randint(46,60)
            numbers.append(randomNumber)
        #O column
        else:
            randomNumber = random.randint(61,75)
            while randomNumber in numbers:
                randomNumber = random.randint(61,75)
            numbers.append(randomNumber)
    
    bingoInfoQuery = conn.cursor()

    for j in range(0, len(numbers)):
        if j == 12:
            bingoInfoQuery.execute("""INSERT INTO dbo.Bingo(BingoCard, BingoNumber, BingoPosition, BingoCalled, BingoMissed)
                                VALUES(""" + str(index) + """, """ + str(numbers[j]) + """, """ + str(j+1) + """, 1, 0)""")
        else :
            bingoInfoQuery.execute("""INSERT INTO dbo.Bingo(BingoCard, BingoNumber, BingoPosition, BingoCalled, BingoMissed)
                                VALUES(""" + str(index) + """, """ + str(numbers[j]) + """, """ + str(j+1) + """, 0, 0)""")
        conn.commit()

for i in range(21, 201):
    createBoard(i)
