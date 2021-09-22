import os, openpyxl, re
import win32com.client as wc

def refresh():
    path = r"c:\users\jeremyshank\documents\bmss assets\reports\quickbooks clients.xlsx"
    print("Refreshing")
    xlapp = wc.DispatchEx("Excel.Application")
    wb = xlapp.Workbooks.Open(path)
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()
    xlapp.DisplayAlerts = False
    wb.Close(SaveChanges=True)
    xlapp.Quit()


def makeFolder(rowNum):
    #path = "Q:"
    #wb = openpyxl.load_workbook("C:\\users\\jeremyshank\\documents\\BMSS Assets\\Reports\\quickbooks clients.xlsx", data_only=True)
    wb = openpyxl.load_workbook(r"C:\Users\jeremyshank\Desktop\Current Employee Scans Folders.xlsx", data_only=True)
    ws = wb['Sheet2']
    #cell = ws['F' + str(rowNum)]
    #Used for employee scan folders
    cell = ws['C' + str(rowNum)]
    if cell.value != None:
        print(rowNum)
        value1 = re.sub("/", "-", cell.value)
        cValue = re.sub(":", "", value1)
        #os.mkdir(path + "\\" + cValue)
        os.mkdir(cValue)
        rowNum += 1
        makeFolder(rowNum)

rowNum = 2

#refresh()
makeFolder(rowNum)
