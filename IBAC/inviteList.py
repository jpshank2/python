import win32com.client as wc

xl = wc.Dispatch('Excel.Application')
wb = xl.Workbooks.Open(r'C:\Users\jeremyshank\Documents\Alumni List.xlsx')
readData = wb.Worksheets('Alumni Contact')

for x in range(1, readData.UsedRange.Rows.Count):
    if readData.Cells(x, 1).Value is None or readData.Cells(x, 2).Value is None or readData.Cells(x, 3).Value is None:
        continue
    else:
        print(readData.Cells(x, 1).Value + ' ' + readData.Cells(x, 2).Value + ' <' + readData.Cells(x, 3).Value + '>')

wb.Close(SaveChanges=False)
xl.Quit()
