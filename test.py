#
# this is a playground to quickly test ideas
#

from dotenv import load_dotenv
import pandas as pd
import json, requests, os, pyodbc

load_dotenv(os.path.dirname(__file__) + '\\.env')

bmss = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

planningJobs = pd.read_sql("""SELECT TOP 10 JH.Job_Idx
	  ,JS.ServIndex
      ,JH.ContIndex
      ,JH.Job_Name
      ,JH.Job_Dept
      ,JH.Job_Partner
	  ,JH.Job_Manager
      ,JH.Job_InCharge
      ,JH.Job_Class
      ,JH.Job_Status
      ,JH.Job_Frequency
      ,JH.Job_Recurring
      ,JH.Job_Period_Start
      ,JH.Job_Period_End
      ,JH.Job_Finish_Deadline
      ,JH.Job_Budget_Hours
      ,JH.Job_Budget_Value
      ,JH.Job_Code
      ,JH.Job_Value
      ,JH.Job_ETCRequired
      ,JH.Job_PreviousJob
      ,JH.Job_NextJob
      ,JH.Job_LimitTimeEntry
      ,JH.Job_Complexity
      ,JH.Job_MasterFile
      ,JH.Job_Office
      ,JH.Job_WorkStatus
      ,JH.Job_CurrentStaff
      ,JH.Job_StaffingType
      ,JH.Job_Biller
      ,JH.PercentComplete
  FROM tblJob_Header JH
  inner join tblJob_Serv JS on JH.Job_Idx=JS.Job_Idx
  WHERE Job_Template=101 and Job_NextJob=0 and Job_Period_End > '12/31/2020'
  ORDER BY JH.Job_Idx
  --OFFSET 3 ROWS
  --FETCH NEXT 1 ROW ONLY
""", bmss)

for x in range(len(planningJobs)):
    if str(planningJobs.iloc[x]['Job_Biller']) == 'nan':
        print('no biller')
    print(str(planningJobs.iloc[x]['ContIndex']) + ' - Biller is: ' + str(planningJobs.iloc[x]['Job_Biller']))
