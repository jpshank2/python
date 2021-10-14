from dotenv import load_dotenv
import os, pyodbc, openpyxl

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')
engine = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

getContIndices = engine.cursor()

getContIndices.execute("""select
	E.Contindex AS [ContIndex],
	E.ClientCode as [Code],
	E.ClientName as [Client],
	(CASE WHEN BG.ContName IS Null THEN 'NONE' ELSE BG.ContName END) AS [BillingGroup],
	E.ClientOffice as [Office],
	Ent.OwnerName as [Entity],
	I.CatName as [Industry],
	E.ClientStatus as [Status],
	OWN.StaffIndex as [Originator],
	E.ClientPartner as [Client Partner],
	E.ClientManager as [Client Manager],
	convert(date,E.ClientCreated) AS [Created],
	(CASE WHEN E.ClientStatus = 'ACTIVE' THEN NULL ELSE convert(date,LOSS.GainLossDate) END) AS [Lost],
	CON.ContCountry AS [Country],
	CON.ContAddress AS [Address],
	CON.ContTownCity AS [City],
	CON.ContCounty AS [State],
	CON.ContPostCode AS [ZipCode],
	CON.ContEmail AS [Email]	
from tblEngagement E
	Left Join tblCategory I on E.ClientIndustry = I.Category and I.CatType = 'INDUSTRY'
	LEFT JOIN tblContacts BG ON E.ClientHold = BG.ContIndex
	left join tblcontacts CON ON E.ContIndex=CON.ContIndex
	inner Join tblClientOrigination OWN on E.Contindex=OWN.ContIndex
	Inner join tblstaff O on OWN.StaffIndex=O.StaffIndex
	Left Join tblOwnerType Ent on E.ClientOwnership = Ent.OwnerIndex
	Left join tblClientGainLoss LOSS ON E.ContIndex=LOSS.ContIndex and Loss.ClientGainLoss='LOSS'
order by E.ContIndex
""")

ContIndices = getContIndices.fetchall()

wb = openpyxl.load_workbook(r'C:\Users\jeremyshank\Desktop\Fake_Companies.xlsx')

ws = wb['Master']

for i in range(len(ContIndices)):
    print(ContIndices[i][0])
    ws.cell(row=i+2, column=2, value=ContIndices[i][0])

wb.save(r'C:\Users\jeremyshank\Desktop\Companies.xlsx')