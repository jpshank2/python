
import win32com.client as wc
import pyodbc, re, shutil, os
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

connDevOps = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS'))

def email(record):
    name = re.sub("'", "''", record[1])
    outlook = wc.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.To = record[0]
    mail.Subject = 'Incomplete Timesheets'
    mail.HTMLBody = '<p>' + name + ',</p><p>&emsp;Your target hours are ' + str(record[3]) + ' and you have only released ' + str(record[5]) + """. Please enter and release all time for yesterday before 2pm today.<br><br>Thanks!</p>"""

    mail.Send()

namerow = conn.cursor()
namerow.execute("""select		p.StaffEMail
			,p.StaffName [Staff Name]
      ,p.StaffIndex
			,convert(varchar(12)
			,dateadd(day,-1,getdate()), 101)  [Entry Date]
			,coalesce(tH.WPHours,0) [Target Hours]
			,coalesce(TDU.DurationUnits,0) [Un-Released Hours]
			,coalesce(TDR.DurationUnits,0) [Released Hours] 
			
From tblStaff P
			left join (select h.staffindex,H.StartDate, datediff(day, h.startdate,  dateadd(day,-1,getdate()) ) DateDelta
			,case datediff(day, h.startdate,  dateadd(day,-1,getdate()) )
when 0  then WorkDay1Hours             
when 1  then WorkDay2Hours                
when 2  then WorkDay3Hours            
when 3  then WorkDay4Hours            
when 4  then WorkDay5Hours          
when 5  then WorkDay6Hours            
when 6  then WorkDay7Hours            
		else 0 end WPHours           
FROM tblTSO_Header H           
			where H.StartDate =   dateadd(week, datediff(week, 7 , dateadd(day,-2,getdate()))+1, 0)) TH on TH.Staffindex = p.Staffindex    
			left join (select h.staffindex, sum(coalesce(d.DurationUnits, 0)) DurationUnits           
FROM tblTSO_Header H            
			left join [tblTSO_Details] D on D.hEADERiNDEX = h.hEADERiNDEX              
WHERE d.Status ='ACTIVE' and d.EntryDate = convert(varchar(12), dateadd(day,-1,getdate()), 101)    
GROUP BY h.staffindex ) TDU  on TDU.Staffindex = p.Staffindex          
			left join (select h.staffindex, sum(coalesce(d.DurationUnits, 0)) DurationUnits               
FROM tblTSO_Header H            
			left join [tblTSO_Details] D on D.hEADERiNDEX = h.hEADERiNDEX              
WHERE d.Status <>'ACTIVE' and d.EntryDate = convert(varchar(12), dateadd(day,-1,getdate()), 101)     
GROUP BY h.staffindex              ) TDR  on TDR.Staffindex = p.Staffindex           
WHERE p.staffTimesheets=1 and p.Staffindex >0 and p.stafftype=1 and p.staffended is null and  isnull(TDR.DurationUnits,0)  < 4 and tH.WPHours >= 4
  UNION ALL
SELECT  p.StaffEMail ,p.StaffName [Staff Name], p.StaffIndex, convert(varchar(12), dateadd(day,-1,getdate()), 101)  [Entry Date],    
 coalesce(tH.WPHours,0) [Target Hours], coalesce(TDU.DurationUnits,0) [Un-Released Hours],  coalesce(TDR.DurationUnits,0) [Released Hours] from  tblStaff P             
 left join (           select h.staffindex,       H.StartDate,      datediff(day, h.startdate,  dateadd(day,-1,getdate()) ) DateDelta,            
 case datediff(day, h.startdate,  dateadd(day,-1,getdate()) )             when 0  then WorkDay1Hours             
 when 1  then WorkDay2Hours                
 when 2  then WorkDay3Hours            
 when 3  then WorkDay4Hours            
 when 4  then WorkDay5Hours          
  when 5  then WorkDay6Hours            
  when 6  then WorkDay7Hours            
  else 0 end WPHours           
  from tblTSO_Header H           
  where H.StartDate =   dateadd(week, datediff(week, 7 , dateadd(day,-2,getdate()))+1, 0)          ) TH  
  on TH.Staffindex = p.Staffindex    
  left join (           select h.staffindex, sum(coalesce(d.DurationUnits, 0)) DurationUnits           
  from tblTSO_Header H            
  left join [tblTSO_Details] D on D.hEADERiNDEX = h.hEADERiNDEX              
  where d.Status ='ACTIVE' 
  and d.EntryDate = convert(varchar(12), dateadd(day,-1,getdate()), 101)          
  group by h.staffindex       ) TDU  on TDU.Staffindex = p.Staffindex          
  left join (           select h.staffindex, sum(coalesce(d.DurationUnits, 0)) DurationUnits               
  from tblTSO_Header H            
  left join [tblTSO_Details] D on D.hEADERiNDEX = h.hEADERiNDEX              
  where d.Status <>'ACTIVE' and d.EntryDate = convert(varchar(12), dateadd(day,-1,getdate()), 101)     
  group by h.staffindex              ) TDR  on TDR.Staffindex = p.Staffindex           
  where p.staffTimesheets=1 
  and p.Staffindex =3 
  and p.stafftype=1 
  and p.staffended is null 
  and  isnull(TDR.DurationUnits,0)  < 4  
  and tH.WPHours >= 4""")

names = namerow.fetchall()

for name in names:
    #email(name)
    bingoCards = connDevOps.cursor()
    bingoCards.execute("""DECLARE @card int
SET @card = (SELECT StaffBingo FROM dbo.tblStaff WHERE StaffIndex = """ + str(name[2]) + """)

UPDATE dbo.Bingo
SET BingoMissed = 1
WHERE BingoCard = @card AND CONVERT(DATE, BingoDate) = CONVERT(DATE, getdate())""")
    connDevOps.commit()