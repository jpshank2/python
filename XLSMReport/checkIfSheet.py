import openpyxl, os

def checkDon(dirName):
    listOfItems = os.listdir(dirName)
    for item in listOfItems:
        fullPath = os.path.join(dirName, item)
        wb = openpyxl.load_workbook(fullPath, read_only=True)
        if "Sheet1" in wb.sheetnames:
            print(item)

checkDon(r'C:\Users\jeremyshank\Desktop\No Macro')