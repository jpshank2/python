import os, zipfile

dir_name = 'T:\\M-Files Importer\\Zak test 12.12'
extension = ".zip"

os.chdir(dir_name) # change directory from working dir to dir with files

for item in os.listdir(dir_name): # loop through items in dir
        print("good so far")
        if item.endswith(extension): # check for ".zip" extension
                print(item)
                file_name = os.path.abspath(item) # get full path of files
                print("have file name " + file_name)
                zip_ref = zipfile.ZipFile(file_name, 'r') # create zipfile object
                print("zip file object")
                zip_ref.extractall(dir_name) # extract file to dir
                print("extracted")
                zip_ref.close() # close file
                print("closed")
                os.remove(file_name) # delete zipped file
                print("good here")