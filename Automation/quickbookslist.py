import os, re, openpyxl

def getQDrive(dirName):
    listOfItems = os.listdir(dirName)
    # badFolders = list()
    # count = 1
    for item in listOfItems:
        fullPath = os.path.join(dirName, item)
        if os.path.isdir(fullPath):
            getQDrive(fullPath)
        else:
            if re.search("TVRG 7.2.2020", item) != None:
                #badFolders.append(item)
                print(os.path.join(dirName, item))

getQDrive(r"C:\Users\jeremyshank\OneDrive - BMSS (1)")