#! python3

import win32com.client as wc
import datetime, openpyxl

path = r'C:\Users\jeremyshank\Documents\BMSS Assets\Reports\Expiring Password.xlsx'
today = datetime.datetime.now()
year = int(datetime.datetime.now().strftime('%Y'))
month = int(datetime.datetime.now().strftime('%m'))
day = int(datetime.datetime.now().strftime('%d'))

if month < 8 and month % 2 != 0:
    length = 31
elif month < 8 and month % 2 == 0:
    length = 30
elif month > 8 and month % 2 != 0:
    length = 30
elif month > 8 and month % 2 == 0:
    length = 31
elif month == 2 and year % 4 != 0:
    length = 28
elif month == 2 and year % 4 == 0:
    length = 29
else:
    length = 31

if (day + 8) > length:
    endDay = ((day + 8) - length)
    month += 1
    endDate = datetime.datetime(year, month, endDay)
else:
    endDay = day + 8
    endDate = datetime.datetime(year, month, endDay)

def refresh():
    print("Refreshing")
    xlapp = wc.DispatchEx("Excel.Application")
    wb = xlapp.Workbooks.Open(path)
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()
    xlapp.DisplayAlerts = False
    wb.Close(SaveChanges=True)
    xlapp.Quit()

def sendMail(name, email, date):
    print("Sending mail")
    outlook = wc.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.To = email + '@bmss.com'
    mail.CC = 'support@bmss.com'
    mail.Subject = 'Password Expiring Soon'
    mail.HTMLBody = '<p>' + name + ',</p><p>&emsp;Your password will be expiring on ' + date + """. Please submit a support ticket in Zeal with the <b>Expiring Password</b> option so we can call you to change it. Please do not send your current or desired password in the support ticket. You can find a list of all the services this will affect <a href="http://zeal.bmss.com/kb/single-sign-on/">here</a>.<br><br>Thanks!</p>"""

    mail.Send()

def getExp():
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb['Sheet2']
    print("checking")
    for row in ws.iter_rows(min_row=2):
        if row[4].value >= today and row[4].value <= endDate:
            name = row[0].value
            email = row[2].value
            date = datetime.datetime.strftime(row[4].value, '%x')
            sendMail(name, email, date)
        elif (row[4].value > endDate):
            break

def main():
    refresh()
    getExp()

main()
