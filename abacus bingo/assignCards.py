import pyodbc, os
import win32com.client as wc
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

cwConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('CW_SERVER') + ';DATABASE=' + os.getenv('CW_DATABASE') + ';UID=' + os.getenv('CW_USER') + ';PWD=' + os.getenv('CW_PASS'))
devopsConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))

getStaffList = cwConn.cursor()
getStaffList.execute("""SELECT Member_RecID, CONCAT(First_Name, ' ', Last_Name) AS Name, Email_Address FROM dbo.Member
WHERE Time_Flag = 1 AND Inactive_Flag = 0""")

getStaff = getStaffList.fetchall()

getOpenCardsList = devopsConn.cursor()
getOpenCardsList.execute("""SELECT BingoCard FROM dbo.BingoCards WHERE BingoUser IS NULL""")

getOpenCards =  getOpenCardsList.fetchall()

for i in range(0, len(getStaff)):
    updateBingo = devopsConn.cursor()
    updateBingo.execute("""UPDATE dbo.BingoCards
    SET BingoUser = """ + str(getStaff[i][0]) + """, BingoCompany = 'ABIT'
    WHERE BingoCard = """ + str(getOpenCards[i][0]))
    devopsConn.commit()

    path = "C:\\Users\\jeremyshank\\Documents\\Bingo\\Cards\\" + str(getOpenCards[i][0]) + ".xlsx"
    outlook = wc.Dispatch('outlook.application')
    send_account = None
    for account in outlook.Session.Accounts:
        if account.DisplayName == 'jshank@abacustechnologies.com':
            send_account = account
            break

    mail = outlook.CreateItem(0)
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))

    mail.To = getStaff[i][2]
    mail.Subject = 'New Bingo Card'
    mail.Attachments.Add(path)
    mail.HTMLBody = '<p>' + getStaff[i][1] + ',</p><p>Your new Bingo card is attached. You will keep this card for each game. Our first game will start <strong style="background-color:rgb(255, 255, 49)">on Monday</strong>.</p><p>Rules for the Bingo Game:<ul><li>Each day make sure you put your hours (a minimum of 8 hours) for the previous day by 12pm noon</li><li>On Mondays make sure you <strong>submit your timesheet</strong> by 12pm noon</li><li>You will receive a bingo number each day by 10:30am</li><li>The system is designed to keep track of your bingo card if you choose not to</li><li>What happens if your time is not entered on time?</li><ul><li>First violation - you are out of the current game and ineligible for prizes</li><li>Second violation - $10 will be taken from your Christmas bonus and put into a pot</li><li>Third violation - $25 will be taken from your Christmas bonus and put into a pot (you have now contributed $35 to the pot)</li><li>Fourth violation - you will need to schedule a time to meet with Brian to talk about your time entry habits and how to change them</li></ul><li>For each perfect game (no violations) your name will be entered into a drawing to win the pot at the end of the year</li><li>Once a new game starts everyone starts back at zero violations</li><li>You can still take EOFO and be in the game as we check for a submitted timesheet on Mondays rather than 8 hours for the previous working day</li><li><strong>If you are on PTO you are still expected to have your time and timesheets in on time - there is no penalty for putting in time early</strong></li><li>When someone wins they will receive a $100 gift card to Amazon</li><li>If two people win they will each receive a $100 gift card to Amazon</li><li>If more than two people win they will each receive a $50 gift card to Amazon</li></ul></p><p>Good luck!</p>'
    
    mail.Send()
