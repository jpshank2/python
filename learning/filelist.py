#! python3

import os, csv, zipfile

def getListOfFiles(dirName):
    # create a list of file and sub directories
    # names in the given directory
    listOfFile = os.listdir(dirName)
    allFiles = list()

    # Iterate over all the entries
    for entry in listOfFile:
        # Create full path
        fullPath = os.path.join(dirName, entry)
        # If entry is a directory then get the list of files in this directory
        if os.path.isdir(fullPath):
            allFiles = allFiles + getListOfFiles(fullPath)
        elif (fullPath[-3:] == "zip"):
            with zipfile.ZipFile(fullPath, "r") as zip_ref:
                zip_ref.extractall(os.path.dirname(os.path.dirname( __file__ )))
            allFiles = allFiles + getListOfFiles(fullPath)
        else:
            allFiles.append(fullPath)

    return allFiles


def main():
    dirName = 'C:\\Users\\zkimes\\Desktop\\y'
    # Get the list of all files in directory tree at given path
    listOfFiles = getListOfFiles(dirName)

    # # Print the files
    # for elem in listOfFiles:
    #     print(elem)

    # print("****************")

    # Get the list of all files in directory tree at given path
    listOfFiles = list()
    
    for (dirpath, dirnames, filenames) in os.walk(dirName):
        listOfFiles += [os.path.join(dirpath, file) for file in filenames]
    
    # Print the files
    with open("C:\\users\\zkimes\\desktop\\filenames.csv", "a") as csv_file:
        for elem in listOfFiles:
            writer = csv.writer(csv_file)
            writer.writerow([elem])
# if __name__ == '__main__':
#     main()
main()
