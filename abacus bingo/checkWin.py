import pyodbc, datetime, os
import win32com.client as wc
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))
cwConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('CW_SERVER') + ';DATABASE=' + os.getenv('CW_DATABASE') + ';UID=' + os.getenv('CW_USER') + ';PWD=' + os.getenv('CW_PASS'))

bingoWinner = False
winners = list()

def emailWinner():
    if len(winners) == 1:
        outlook = wc.Dispatch('outlook.application')
        send_account = None
        for account in outlook.Session.Accounts:
            if account.DisplayName == 'jshank@abacustechnologies.com':
                send_account = account
                break
        mail = outlook.CreateItem(0)
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))
        mail.To = 'AITS@abacusit.com'
        mail.Subject = winners[0] + ' Won Bingo!'
        mail.HTMLBody = '<p>' + winners[0] + ' won this game of bingo!</p><br><p>We will start a new game tomorrow!</p>'

        mail.Send()
    else:
        allWinners = ""
        for i in range(len(winners) - 1):
            allWinners += winners[i] + '; '
        allWinners += winners[-1] + ' '
        outlook = wc.Dispatch('outlook.application')
        send_account = None
        for account in outlook.Session.Accounts:
            if account.DisplayName == 'jshank@abacustechnologies.com':
                send_account = account
                break
        mail = outlook.CreateItem(0)
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))
        mail.To = 'AITS@abacusit.com'
        mail.Subject = allWinners + 'Won Bingo!'
        mail.HTMLBody = '<p>' + allWinners + 'won this game of bingo!</p><br><p>We will start a new game tomorrow!</p>'

        mail.Send()



def emailNoWinner():
    outlook = wc.Dispatch('outlook.application')
    send_account = None
    for account in outlook.Session.Accounts:
        if account.DisplayName == 'jshank@abacustechnologies.com':
            send_account = account
            break
    mail = outlook.CreateItem(0)
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))
    mail.To = 'jeremyshank@bmss.com'
    mail.Subject = 'No Winner for ' + datetime.datetime.today().strftime('%m/%d/%y')
    mail.Body = 'EOM'

    mail.Send()

def cleanup():
    cleanupBingo = conn.cursor()
    cleanupBingo.execute("""UPDATE dbo.Bingo
        SET BingoPerfect = 
            CASE
                WHEN BingoMissed = 0 THEN BingoPerfect + 1
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
            END
        WHERE BingoCard IN (SELECT BingoCard FROM dbo.BingoCards WHERE BingoCompany = 'ABIT')""")
    conn.commit()

    resetBingo = conn.cursor()
    resetBingo.execute("""UPDATE dbo.Bingo SET BingoCalled = 0, BingoDate = NULL, BingoMissed = 0 WHERE BingoNumber != 0 AND BingoCard IN (SELECT BingoCard FROM dbo.BingoCards WHERE BingoCompany = 'ABIT')
    UPDATE dbo.Bingo SET BingoMissed = 0, BingoDate = NULL WHERE BingoNumber = 0 AND BingoCard IN (SELECT BingoCard FROM dbo.BingoCards WHERE BingoCompany = 'ABIT')""")
    conn.commit()

def checkWin(card, staff): 
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
                if win[0] == 1 or win[0] == 6 or win[0] == 11 or win[0] == 16 or win[0] == 21:
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
                if win[0] == 1 or win[0] == 2 or win[0] == 3 or win[0] == 4 or win[0] == 5:
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
    
    def check(card, staff):
        if winColumn(card) or winRightDiag(card) or winLeftDiag(card) or winRow(card):
            winners.append(staff[0])
            return True
        else:
            return False
    
    return check(card, staff)

getCards = conn.cursor()
getCards.execute("""SELECT BC.BingoCard, BC.BingoUser, MAX(B.BingoMissed) AS BingoMissed
FROM BingoCards BC
INNER JOIN dbo.Bingo B ON B.BingoCard = BC.BingoCard
WHERE BC.BingoUser IS NOT NULL AND BC.BingoCompany = 'ABIT'
GROUP BY BC.BingoCard, BC.BingoUser
HAVING MAX(B.BingoMissed) < 1""")

cards = getCards.fetchall()

for card in cards:
    getStaffInfo = cwConn.cursor()
    getStaffInfo.execute("""SELECT CONCAT(First_Name, ' ', Last_Name) AS Name, CAST(TE_Date_Start as date) AS Date_Started
    FROM dbo.Member
    WHERE Member_RecID = """ + str(card[1]))

    staffInfo = getStaffInfo.fetchall()

    for staff in staffInfo:
        getMyNumbers = conn.cursor()
        getMyNumbers.execute("""SELECT BingoNumber, 
            BingoPosition, 
            BingoDate 
            FROM dbo.Bingo 
            WHERE BingoCard = """ + str(card[0]) + """ 
            AND BingoCalled = 1 
            AND (BingoDate > '""" + str(staff[1]) + """' OR BingoNumber = 0)""")

        myNumbers = getMyNumbers.fetchall()

        if len(myNumbers) >= 5:
            checkWin(myNumbers, staff)

if len(winners) > 0:
    print(winners)
    # emailWinner()
    # cleanup()
else:
    emailNoWinner()
