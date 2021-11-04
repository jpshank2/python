import re, pyodbc, os
import pandas as pd
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

bmss = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

abit = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('CW_SERVER') + ';DATABASE=' + os.getenv('CW_DATABASE') + ';UID=' + os.getenv('CW_USER') + ';PWD=' + os.getenv('CW_PASS'))

pbsDF = pd.read_excel(r'C:\Users\jeremyshank\Documents\BMSS Assets\Reports\PBS Clients.xlsx', 'Clients')

path = r'C:\Users\jeremyshank\Desktop\cfo.xlsx'

df = pd.read_excel(path, 'Table1')

companies = pd.DataFrame()

for x in range(df.shape[0]):
    row = df.iloc[x]
    bmssDF = pd.read_sql("""SELECT ClientName
	,ClientStatus
	,P.StaffName AS [Partner]
	,M.StaffName AS [Manager]
FROM dbo.tblEngagement E
	INNER JOIN dbo.tblStaff P ON P.StaffIndex = E.ClientPartner
	INNER JOIN dbo.tblStaff M ON M.StaffIndex = E.ClientManager
WHERE ClientName LIKE '%""" + re.sub("'", "''", row['Company']) + """%'
ORDER BY ClientStatus""", bmss)
    BMSSbreakdown = bmssDF['ClientStatus'].value_counts(normalize=True) * 100
    
    if bmssDF.shape[0] == 1:
        bmssS = bmssDF.iloc[0]['ClientStatus']
        bmssP = bmssDF.iloc[0]['Partner']
        bmssM = bmssDF.iloc[0]['Manager']
    elif bmssDF.shape[0] > 1:
        bmssS = 'PROBABLE' if BMSSbreakdown['ACTIVE'] > 50.0 else 'UNLIKELY'
        bmssP = bmssDF.iloc[0]['Partner']
        bmssM = bmssDF.iloc[0]['Manager']
    else:
        bmssS = None
        bmssP = None
        bmssM = None
    
    abitDF = pd.read_sql("""SELECT C.Company_Name AS [Client]
	,CS.Description AS [Status]
	,CASE 
		WHEN C.Company_RecID IN (SELECT Company_RecID FROM dbo.AGR_Header WHERE AGR_Date_End IS NULL OR AGR_Date_End > GETDATE()) THEN 'Yes'
		WHEN C.Company_RecID IN (SELECT Company_RecID FROM dbo.AGR_Header WHERE AGR_Date_End < GETDATE()) THEN 'Had'
		ELSE 'No'
	END AS [Agreement]
FROM dbo.Company C
	INNER JOIN dbo.Company_Status CS ON CS.Company_Status_RecID = C.Company_Status_RecID
	INNER JOIN dbo.Company_Company_Type CCT ON CCT.Company_RecID = C.Company_RecID AND CCT.Company_Type_RecID IN (1, 37, 39)
WHERE C.Company_Name LIKE '%""" + re.sub("'", "''", row['Company']) + """%'""", abit)
    ABITbreakdown = abitDF['Status'].value_counts(normalize=True) * 100

    if abitDF.shape[0] == 1:
        abitS = abitDF.iloc[0]['Status']
        abitA = abitDF.iloc[0]['Agreement']
    elif abitDF.shape[0] > 1:
        abitS = 'Probable' if ABITbreakdown['Active'] > 50.0 else 'Unlikely'
        abitA = abitDF.iloc[0]['Agreement']
    else:
        abitS = None
        abitA = None
    
    pbsClients = pbsDF[pbsDF['Client Name'].str.contains(row['Company'])]
    PBSbreakdown = pbsClients['Status'].value_counts(normalize=True) * 100

    if pbsClients.shape[0] == 1:
        pbsS = pbsClients.iloc[0]['Status']
        pbsM = pbsClients.iloc[0]['CSR']
    elif pbsClients.shape[0] > 1:
        pbsS = 'Probable' if PBSbreakdown['Active'] > 50.0 else 'Unlikely'
        pbsM = pbsClients.iloc[0]['CSR']
    else:
        pbsS = None
        pbsM = None
    
    companies = companies.append({'company': row['Company'], 'cfo': row['CFO'], 'abitStatus': abitS, 'abitAgreement': abitA, 'bmssStatus': bmssS, 'bmssPartner': bmssP, 'bmssManager': bmssM, 'pbsStatus': pbsS, 'pbsManager': pbsM}, ignore_index=True)


writer = pd.ExcelWriter(r'C:\users\jeremyshank\desktop\result.xlsx')
companies.to_excel(writer)
writer.save()

