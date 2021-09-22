Set objFSO = CreateObject("Scripting.FileSystemObject")

src_file = objFSO.GetAbsolutePathName("C:\AutomatedReports\M and M Weekly.xlsx")

Dim oExcel
Set oExcel = CreateObject("Excel.Application")

Dim oBook
Set oBook = oExcel.Workbooks.Open(src_file)
oExcel.DisplayAlerts = False
oExcel.AskToUpdateLinks = False
oExcel.AlertBeforeOverwriting = False


oBook.RefreshAll

wscript.sleep 100*100

oBook.RefreshAll

wscript.sleep 100*100

oBook.Save

oBook.Close (False)
oExcel.Quit

'Dim outlook, email
Set outlook = CreateObject("Outlook.Application")
Set email = outlook.CreateItem(0)

with email
	.to = "hrussell@bmss.com"
	.bcc = "jeremyshank@bmss.com"
	.subject = "Homeroom Leader Check Ins"
	.Attachments.Add src_file
	.HTMLBody = "See attached spreadsheet..."
	.Send
wscript.sleep 100*100
End with


wscript.quit