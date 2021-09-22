from dotenv import load_dotenv
import os, pyodbc

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')
engine = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

sqlTest = engine.cursor()

sqlTest.execute('SELECT * FROM dbo.tblStaff')

result = sqlTest.fetchall()

print(result)