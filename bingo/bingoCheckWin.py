import pyodbc, datetime, os
import win32com.client as wc
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))

bingoWinner = False
# global start
# start = 0
# global win
# win = []
# global winner
# winner = False

def emailWinner(winner):
    outlook = wc.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'bingo@bmss.com'
    mail.Subject = winner[0][2] + ' Won Bingo!'
    mail.HTMLBody = '<p>' + winner[0][2] + ' won this game of bingo!</p><br><p>We will start a new game tomorrow!</p>'

    mail.Send()

def emailNoWinner():
    outlook = wc.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'jeremyshank@bmss.com'
    mail.Subject = 'No Winner for ' + datetime.datetime.today().strftime('%m/%d/%y')
    mail.Body = 'EOM'

    mail.Send()

def cleanup():
    cleanupBingo = conn.cursor()
    cleanupBingo.execute("""UPDATE dbo.Bingo
        SET BingoPerfect = 
            CASE
                WHEN BingoMissed = 0 THEN BingoPerfect - 1
                ELSE BingoPerfect
            END,
        BingoMissedTwice = 
            CASE
                WHEN BingoMissed > 1 THEN BingoMissedTwice + 1
                ELSE BingoMissedTwice
            END,
        BingoMissedThrice = 
            CASE
                WHEN BingoMissed > 2 THEN BingoMissedThrice + 1
                ELSE BingoMissedThrice
            END""")
    conn.commit()

def checkWin(card): 
    def winColumn(card):
        start = 0
        win = []
        while start < len(card):
            i = start
            win.append(card[i][1])
            j = i + 1
            for tile in range(j, len(card)):
                if card[tile][1] - card[i][1] == 1:
                    win.append(card[tile][1])
                    i = tile
            
            if len(win) == 5:
                if win[4] - win[0] == 4:
                    return True
            else:
                win = []
                start += 1
        return False

    def winRow(card):
        start = 0
        win = []
        while start < len(card):
            i = start
            win.append(card[i][1])
            j = i + 1
            for tile in range(j, len(card)):
                if card[tile][1] - card[i][1] == 5:
                    win.append(card[tile][1])
                    i = tile
            
            if len(win) == 5:
                if win[4] - win[0] == 20:
                    return True
            else:
                win = []
                start += 1
        return False

    def winLeftDiag(card):
        start = 0
        win = []
        if card[start][1] == 1:
            win.append(card[start][1])
            i = start
            j = i + 1
            for tile in range(j, len(card)):
                if card[tile][1] - card[i][1] == 6:
                    win.append(card[tile][1])
                    i = tile
            if len(win) == 5:
                if win[0] == 1 and win[4] == 25:
                    return True
            else:
                return False
        else:
            return False
    
    def winRightDiag(card):
        start = 0
        win = []
        while start < len(card):
            if card[start][1] == 5:
                win.append(card[start][1])
                i = start
                j = i + 1
                for tile in range(j, len(card)):
                    if card[tile][1] - card[i][1] == 4:
                        win.append(card[tile][1])
                        i = tile
                if len(win) == 5:
                    if win[0] == 5 and win[4] == 21:
                        return True
                else:
                    return False
            else:
                win = []
                start += 1
        return False
    
    def check(card):
        if winColumn(card) or winRightDiag(card) or winLeftDiag(card) or winRow(card):
            return True
        else:
            return False
    
    return check(card)

getCards = conn.cursor()
getCards.execute("""SELECT DISTINCT BingoCard
                        FROM dbo.Bingo""")

cards = getCards.fetchall()

for card in cards:
    getMyCard = conn.cursor()
    getMyCard.execute("""DECLARE @missed int
                            SET @missed = (SELECT BingoMissed FROM dbo.Bingo WHERE BingoNumber = 0 AND BingoCard = """ + str(card[0]) + """)
                            
                            SELECT BingoNumber, BingoPosition, S.StaffName,
                            CASE
                                WHEN @missed = 0 THEN 0
                                ELSE 1
                            END AS BingoMissed
                            FROM dbo.Bingo B
                            INNER JOIN dbo.tblStaff S ON B.BingoCard = S.StaffBingo
                            WHERE BingoCalled = 1 AND BingoMissed < 1 AND BingoCard = """ + str(card[0]) + """ AND BingoDate > S.StaffStarted
                            ORDER BY BingoPosition""")
    myCard = getMyCard.fetchall()
    
    if len(myCard) > 0:
        if checkWin(myCard):
            #add mail function
            emailWinner(myCard)
            #cleanup()
            bingoWinner = True

# if checkWin([(1, 1, 'J', 0), (1, 2, "j", 0), (1, 5, "J", 0), (1, 10, "j", 0), (1, 13, "j", 0), (1, 17, "j", 0), (1, 21, "j", 0)]):
#     print("Winner!")
#     bingoWinner = True

if not bingoWinner:
    #add no winner mail function
    emailNoWinner()
