import pyodbc, re, openpyxl

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

getJobNames = conn.cursor()

getJobNames.execute("""SELECT DISTINCT Job_Name
FROM dbo.tblJob_Header
WHERE Job_Period_End > '2011'
ORDER BY Job_Name""")

jobNames = getJobNames.fetchall()

for job in jobNames:
    jobLength = 0
    timesCounted = 0
    iterate = 1
    if re.search("'", job[0]) != None:
        job[0] = re.sub("'", "''", job[0])
    checkAllNamesLength = conn.cursor()
    checkAllNamesLength.execute("""SELECT Job_Name FROM dbo.tblJob_Header WHERE Job_Name = '""" + job[0] + "'")
    jobNamesLength = checkAllNamesLength.fetchall()
    for name in jobNamesLength:
        if jobLength != len(name[0]):
            timesCounted += 1
            if timesCounted > 1:
                checkJobNames = conn.cursor()
                checkJobNames.execute("""SELECT Job_Idx, E.ContIndex, E.ClientName, Job_Name
                    FROM dbo.tblJob_Header JH
                    INNER JOIN dbo.tblEngagement E ON E.ContIndex = JH.ContIndex
                    WHERE Job_Name = '""" + name[0] + """'""")
                jobsByJobName = checkJobNames.fetchall()

                for clientJob in jobsByJobName:
                    print(clientJob[3] + "-" + clientJob[2] + ": " +str(len(clientJob[3])))
                    # wb = openpyxl.Workbook()
                    # ws = wb.create_sheet('Duplicates')
                    # ws.cell(row=iterate, column=1, value=clientJob[0])
                    # ws.cell(row=iterate, column=2, value=clientJob[1])
                    # ws.cell(row=iterate, column=3, value=clientJob[2])
                    # ws.cell(row=iterate, column=4, value=clientJob[3])
                    # ws.cell(row=iterate, column=5, value=len(clientJob[3]))
                    # iterate += 1
                
            jobLength = len(name[0])
