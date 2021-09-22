import pyodbc, os
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS'))

getJobNames = conn.cursor()

getJobNames.execute("""SELECT Job_Name
FROM dbo.tblJob_Header 
WHERE ContIndex IN (select ContIndex from dbo.tblEngagement where ClientCode = '31657-000') 
AND Job_Name = '2021 Business Personal Property Tax - Sky High LLC'
ORDER BY Job_Name""")

jobNames = getJobNames.fetchall()

for job in jobNames:
    x = job[0].split()
    print(x)
    print(len(job[0]))