import pyodbc, os
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))


getStaff = conn.cursor()
getStaff.execute("""SELECT [StaffIndex]
      ,[StaffName]
  FROM [DataWarehouse].[dbo].[tblStaff]
  WHERE StaffEnded IS NULL AND StaffName NOT IN ('Administration', 'E-File Pool', 'Cindy Cpa')
  ORDER BY StaffIndex""")

staff = getStaff.fetchall()

for i in range(0, len(staff)):
    updateBingo = conn.cursor()
    updateBingo.execute("""UPDATE dbo.tblStaff
                            SET StaffBingo = """ + str(i+1) + """
                            WHERE StaffIndex = """ + str(staff[i][0]))
    conn.commit()