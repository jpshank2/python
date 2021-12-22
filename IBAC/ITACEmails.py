import win32com.client as wc

xl = wc.Dispatch('Excel.Application')
wb = xl.Workbooks.Open(r'C:\Users\jeremyshank\Desktop\ITAC Event Registration.xlsx')
readData = wb.Worksheets('Form Responses 1')
outlook = wc.Dispatch('outlook.application')

for x in range(2, readData.UsedRange.Rows.Count):
    mail = outlook.CreateItem(0)
    mail.To = readData.Cells(x, 2).Value
    mail.Subject = 'ITAC Event Registration'
    mail.HTMLBody = '<p> Hey ' + readData.Cells(x, 3).Value + ',</p><p>This is Jeremy Shank with the Innovate Birmingham Alumni Council. You signed up as interested in our event partnering with ITAC for a round of mock interviews and resume help. We have selected next Monday, November 22nd and next Tuesday, November 23rd from 11 AM to 1 PM as our time for this event. Are you available for either of those dates? If you would let me know if you can make it to either day or both days we will assign you a time to meet with one of their recruiters!</p><p>Thank you!</p>'
    mail.Send()

wb.Close(SaveChanges=False)
xl.Quit()
