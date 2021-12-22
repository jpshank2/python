import pyodbc, os
import win32com.client as wc
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

bmss = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

getStaffList = bmss.cursor()
getStaffList.execute("""SELECT P.StaffName AS [Partner], P.StaffEMail, COUNT(*) AS [Clients]
FROM dbo.tblEngagement E
	INNER JOIN dbo.tblStaff P ON P.StaffIndex = E.ClientPartner
	INNER JOIN dbo.tblStaff M ON M.StaffIndex = E.ClientManager
	INNER JOIN dbo.tblCategory I ON I.Category = E.ClientIndustry AND I.CatType = 'INDUSTRY'
	--INNER JOIN dbo.tblTranDebtor AR ON AR.ContIndex = E.ContIndex AND AR.DebtTranType IN ('3', '6')
WHERE E.ContIndex IN (
	SELECT DISTINCT ContIndex
	FROM dbo.tblTranWIP
	WHERE WIPService = 'ACCTG')
AND E.ClientStatus <> 'LOST'
AND (I.CatName LIKE '%Real%Estate%' OR I.CatName LIKE '%construction%')
--AND AR.DebtTranDate > '7/1/2020'
--AND P.StaffName <> 'Scott Garrison'
--AND E.ClientOffice = 'BHM'
GROUP BY P.StaffName, P.StaffEMail
--HAVING SUM(AR.DebtTranAmount) > 5000
ORDER BY 3 DESC""")

getStaff = getStaffList.fetchall()


outlook = wc.Dispatch('outlook.application')

for x in range(len(getStaff)):
    send_account = None
    for account in outlook.Session.Accounts:
        if account.DisplayName == 'jshank@abacustechnologies.com':
            send_account = account
            break

    mail = outlook.CreateItem(0)
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))

    mail.To = getStaff[x][1]
    mail.CC = 'cneal@abacustechnologies.com'
    mail.Subject = 'Business Intelligence for Real Estate and Construction'
    mail.HTMLBody = r"""<head>
    <meta http-equiv=Content-Type content="text/html; charset=windows-1252">
    <meta name=ProgId content=Word.Document>
    <meta name=Generator content="Microsoft Word 15">
    <meta name=Originator content="Microsoft Word 15">
    <link rel=File-List href="C:\Users\jeremyshank\AppData\Roaming\Microsoft\Signatures\abacus_files/filelist.xml">
    <link rel=Edit-Time-Data href="C:\Users\jeremyshank\AppData\Roaming\Microsoft\Signatures\abacus_files/editdata.mso">
    <!--[if !mso]>
    <style>
    v\:* {behavior:url(#default#VML);}
    o\:* {behavior:url(#default#VML);}
    w\:* {behavior:url(#default#VML);}
    .shape {behavior:url(#default#VML);}
    </style>
    <![endif]--><!--[if gte mso 9]><xml>
    <o:OfficeDocumentSettings>
    <o:AllowPNG/>
    </o:OfficeDocumentSettings>
    </xml><![endif]-->
    <link rel=themeData href="C:\Users\jeremyshank\AppData\Roaming\Microsoft\Signatures\abacus_files/themedata.thmx">
    <link rel=colorSchemeMapping href="C:\Users\jeremyshank\AppData\Roaming\Microsoft\Signatures\abacus_files/colorschememapping.xml">
    <!--[if gte mso 9]><xml>
    <w:WordDocument>
    <w:View>Normal</w:View>
    <w:Zoom>0</w:Zoom>
    <w:TrackMoves/>
    <w:TrackFormatting/>
    <w:PunctuationKerning/>
    <w:ValidateAgainstSchemas/>
    <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>
    <w:IgnoreMixedContent>false</w:IgnoreMixedContent>
    <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>
    <w:DoNotPromoteQF/>
    <w:LidThemeOther>EN-US</w:LidThemeOther>
    <w:LidThemeAsian>X-NONE</w:LidThemeAsian>
    <w:LidThemeComplexScript>X-NONE</w:LidThemeComplexScript>
    <w:DoNotShadeFormData/>
    <w:Compatibility>
    <w:BreakWrappedTables/>
    <w:SnapToGridInCell/>
    <w:WrapTextWithPunct/>
    <w:UseAsianBreakRules/>
    <w:DontGrowAutofit/>
    <w:SplitPgBreakAndParaMark/>
    <w:EnableOpenTypeKerning/>
    <w:DontFlipMirrorIndents/>
    <w:OverrideTableStyleHps/>
    <w:UseFELayout/>
    </w:Compatibility>
    <m:mathPr>
    <m:mathFont m:val="Cambria Math"/>
    <m:brkBin m:val="before"/>
    <m:brkBinSub m:val="&#45;-"/>
    <m:smallFrac m:val="off"/>
    <m:dispDef/>
    <m:lMargin m:val="0"/>
    <m:rMargin m:val="0"/>
    <m:defJc m:val="centerGroup"/>
    <m:wrapIndent m:val="1440"/>
    <m:intLim m:val="subSup"/>
    <m:naryLim m:val="undOvr"/>
    </m:mathPr></w:WordDocument>
    </xml><![endif]--><!--[if gte mso 9]><xml>
    <w:LatentStyles DefLockedState="false" DefUnhideWhenUsed="false"
    DefSemiHidden="false" DefQFormat="false" DefPriority="99"
    LatentStyleCount="376">
    <w:LsdException Locked="false" Priority="0" QFormat="true" Name="Normal"/>
    <w:LsdException Locked="false" Priority="9" QFormat="true" Name="heading 1"/>
    <w:LsdException Locked="false" Priority="9" SemiHidden="true"
    UnhideWhenUsed="true" QFormat="true" Name="heading 2"/>
    <w:LsdException Locked="false" Priority="9" SemiHidden="true"
    UnhideWhenUsed="true" QFormat="true" Name="heading 3"/>
    <w:LsdException Locked="false" Priority="9" SemiHidden="true"
    UnhideWhenUsed="true" QFormat="true" Name="heading 4"/>
    <w:LsdException Locked="false" Priority="9" SemiHidden="true"
    UnhideWhenUsed="true" QFormat="true" Name="heading 5"/>
    <w:LsdException Locked="false" Priority="9" SemiHidden="true"
    UnhideWhenUsed="true" QFormat="true" Name="heading 6"/>
    <w:LsdException Locked="false" Priority="9" SemiHidden="true"
    UnhideWhenUsed="true" QFormat="true" Name="heading 7"/>
    <w:LsdException Locked="false" Priority="9" SemiHidden="true"
    UnhideWhenUsed="true" QFormat="true" Name="heading 8"/>
    <w:LsdException Locked="false" Priority="9" SemiHidden="true"
    UnhideWhenUsed="true" QFormat="true" Name="heading 9"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="index 1"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="index 2"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="index 3"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="index 4"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="index 5"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="index 6"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="index 7"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="index 8"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="index 9"/>
    <w:LsdException Locked="false" Priority="39" SemiHidden="true"
    UnhideWhenUsed="true" Name="toc 1"/>
    <w:LsdException Locked="false" Priority="39" SemiHidden="true"
    UnhideWhenUsed="true" Name="toc 2"/>
    <w:LsdException Locked="false" Priority="39" SemiHidden="true"
    UnhideWhenUsed="true" Name="toc 3"/>
    <w:LsdException Locked="false" Priority="39" SemiHidden="true"
    UnhideWhenUsed="true" Name="toc 4"/>
    <w:LsdException Locked="false" Priority="39" SemiHidden="true"
    UnhideWhenUsed="true" Name="toc 5"/>
    <w:LsdException Locked="false" Priority="39" SemiHidden="true"
    UnhideWhenUsed="true" Name="toc 6"/>
    <w:LsdException Locked="false" Priority="39" SemiHidden="true"
    UnhideWhenUsed="true" Name="toc 7"/>
    <w:LsdException Locked="false" Priority="39" SemiHidden="true"
    UnhideWhenUsed="true" Name="toc 8"/>
    <w:LsdException Locked="false" Priority="39" SemiHidden="true"
    UnhideWhenUsed="true" Name="toc 9"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Normal Indent"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="footnote text"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="annotation text"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="header"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="footer"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="index heading"/>
    <w:LsdException Locked="false" Priority="35" SemiHidden="true"
    UnhideWhenUsed="true" QFormat="true" Name="caption"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="table of figures"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="envelope address"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="envelope return"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="footnote reference"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="annotation reference"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="line number"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="page number"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="endnote reference"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="endnote text"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="table of authorities"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="macro"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="toa heading"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List Bullet"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List Number"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List 2"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List 3"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List 4"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List 5"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List Bullet 2"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List Bullet 3"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List Bullet 4"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List Bullet 5"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List Number 2"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List Number 3"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List Number 4"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List Number 5"/>
    <w:LsdException Locked="false" Priority="10" QFormat="true" Name="Title"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Closing"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Signature"/>
    <w:LsdException Locked="false" Priority="1" SemiHidden="true"
    UnhideWhenUsed="true" Name="Default Paragraph Font"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Body Text"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Body Text Indent"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List Continue"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List Continue 2"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List Continue 3"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List Continue 4"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="List Continue 5"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Message Header"/>
    <w:LsdException Locked="false" Priority="11" QFormat="true" Name="Subtitle"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Salutation"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Date"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Body Text First Indent"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Body Text First Indent 2"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Note Heading"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Body Text 2"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Body Text 3"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Body Text Indent 2"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Body Text Indent 3"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Block Text"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Hyperlink"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="FollowedHyperlink"/>
    <w:LsdException Locked="false" Priority="22" QFormat="true" Name="Strong"/>
    <w:LsdException Locked="false" Priority="20" QFormat="true" Name="Emphasis"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Document Map"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Plain Text"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="E-mail Signature"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="HTML Top of Form"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="HTML Bottom of Form"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Normal (Web)"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="HTML Acronym"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="HTML Address"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="HTML Cite"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="HTML Code"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="HTML Definition"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="HTML Keyboard"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="HTML Preformatted"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="HTML Sample"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="HTML Typewriter"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="HTML Variable"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Normal Table"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="annotation subject"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="No List"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Outline List 1"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Outline List 2"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Outline List 3"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Simple 1"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Simple 2"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Simple 3"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Classic 1"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Classic 2"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Classic 3"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Classic 4"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Colorful 1"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Colorful 2"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Colorful 3"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Columns 1"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Columns 2"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Columns 3"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Columns 4"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Columns 5"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Grid 1"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Grid 2"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Grid 3"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Grid 4"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Grid 5"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Grid 6"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Grid 7"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Grid 8"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table List 1"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table List 2"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table List 3"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table List 4"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table List 5"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table List 6"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table List 7"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table List 8"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table 3D effects 1"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table 3D effects 2"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table 3D effects 3"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Contemporary"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Elegant"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Professional"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Subtle 1"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Subtle 2"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Web 1"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Web 2"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Web 3"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Balloon Text"/>
    <w:LsdException Locked="false" Priority="39" Name="Table Grid"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Table Theme"/>
    <w:LsdException Locked="false" SemiHidden="true" Name="Placeholder Text"/>
    <w:LsdException Locked="false" Priority="1" QFormat="true" Name="No Spacing"/>
    <w:LsdException Locked="false" Priority="60" Name="Light Shading"/>
    <w:LsdException Locked="false" Priority="61" Name="Light List"/>
    <w:LsdException Locked="false" Priority="62" Name="Light Grid"/>
    <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1"/>
    <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2"/>
    <w:LsdException Locked="false" Priority="65" Name="Medium List 1"/>
    <w:LsdException Locked="false" Priority="66" Name="Medium List 2"/>
    <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1"/>
    <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2"/>
    <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3"/>
    <w:LsdException Locked="false" Priority="70" Name="Dark List"/>
    <w:LsdException Locked="false" Priority="71" Name="Colorful Shading"/>
    <w:LsdException Locked="false" Priority="72" Name="Colorful List"/>
    <w:LsdException Locked="false" Priority="73" Name="Colorful Grid"/>
    <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 1"/>
    <w:LsdException Locked="false" Priority="61" Name="Light List Accent 1"/>
    <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 1"/>
    <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 1"/>
    <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 1"/>
    <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 1"/>
    <w:LsdException Locked="false" SemiHidden="true" Name="Revision"/>
    <w:LsdException Locked="false" Priority="34" QFormat="true"
    Name="List Paragraph"/>
    <w:LsdException Locked="false" Priority="29" QFormat="true" Name="Quote"/>
    <w:LsdException Locked="false" Priority="30" QFormat="true"
    Name="Intense Quote"/>
    <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 1"/>
    <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 1"/>
    <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 1"/>
    <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 1"/>
    <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 1"/>
    <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 1"/>
    <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 1"/>
    <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 1"/>
    <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 2"/>
    <w:LsdException Locked="false" Priority="61" Name="Light List Accent 2"/>
    <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 2"/>
    <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 2"/>
    <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 2"/>
    <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 2"/>
    <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 2"/>
    <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 2"/>
    <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 2"/>
    <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 2"/>
    <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 2"/>
    <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 2"/>
    <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 2"/>
    <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 2"/>
    <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 3"/>
    <w:LsdException Locked="false" Priority="61" Name="Light List Accent 3"/>
    <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 3"/>
    <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 3"/>
    <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 3"/>
    <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 3"/>
    <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 3"/>
    <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 3"/>
    <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 3"/>
    <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 3"/>
    <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 3"/>
    <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 3"/>
    <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 3"/>
    <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 3"/>
    <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 4"/>
    <w:LsdException Locked="false" Priority="61" Name="Light List Accent 4"/>
    <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 4"/>
    <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 4"/>
    <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 4"/>
    <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 4"/>
    <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 4"/>
    <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 4"/>
    <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 4"/>
    <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 4"/>
    <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 4"/>
    <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 4"/>
    <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 4"/>
    <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 4"/>
    <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 5"/>
    <w:LsdException Locked="false" Priority="61" Name="Light List Accent 5"/>
    <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 5"/>
    <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 5"/>
    <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 5"/>
    <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 5"/>
    <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 5"/>
    <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 5"/>
    <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 5"/>
    <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 5"/>
    <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 5"/>
    <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 5"/>
    <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 5"/>
    <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 5"/>
    <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 6"/>
    <w:LsdException Locked="false" Priority="61" Name="Light List Accent 6"/>
    <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 6"/>
    <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 6"/>
    <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 6"/>
    <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 6"/>
    <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 6"/>
    <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 6"/>
    <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 6"/>
    <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 6"/>
    <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 6"/>
    <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 6"/>
    <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 6"/>
    <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 6"/>
    <w:LsdException Locked="false" Priority="19" QFormat="true"
    Name="Subtle Emphasis"/>
    <w:LsdException Locked="false" Priority="21" QFormat="true"
    Name="Intense Emphasis"/>
    <w:LsdException Locked="false" Priority="31" QFormat="true"
    Name="Subtle Reference"/>
    <w:LsdException Locked="false" Priority="32" QFormat="true"
    Name="Intense Reference"/>
    <w:LsdException Locked="false" Priority="33" QFormat="true" Name="Book Title"/>
    <w:LsdException Locked="false" Priority="37" SemiHidden="true"
    UnhideWhenUsed="true" Name="Bibliography"/>
    <w:LsdException Locked="false" Priority="39" SemiHidden="true"
    UnhideWhenUsed="true" QFormat="true" Name="TOC Heading"/>
    <w:LsdException Locked="false" Priority="41" Name="Plain Table 1"/>
    <w:LsdException Locked="false" Priority="42" Name="Plain Table 2"/>
    <w:LsdException Locked="false" Priority="43" Name="Plain Table 3"/>
    <w:LsdException Locked="false" Priority="44" Name="Plain Table 4"/>
    <w:LsdException Locked="false" Priority="45" Name="Plain Table 5"/>
    <w:LsdException Locked="false" Priority="40" Name="Grid Table Light"/>
    <w:LsdException Locked="false" Priority="46" Name="Grid Table 1 Light"/>
    <w:LsdException Locked="false" Priority="47" Name="Grid Table 2"/>
    <w:LsdException Locked="false" Priority="48" Name="Grid Table 3"/>
    <w:LsdException Locked="false" Priority="49" Name="Grid Table 4"/>
    <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark"/>
    <w:LsdException Locked="false" Priority="51" Name="Grid Table 6 Colorful"/>
    <w:LsdException Locked="false" Priority="52" Name="Grid Table 7 Colorful"/>
    <w:LsdException Locked="false" Priority="46"
    Name="Grid Table 1 Light Accent 1"/>
    <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 1"/>
    <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 1"/>
    <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 1"/>
    <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 1"/>
    <w:LsdException Locked="false" Priority="51"
    Name="Grid Table 6 Colorful Accent 1"/>
    <w:LsdException Locked="false" Priority="52"
    Name="Grid Table 7 Colorful Accent 1"/>
    <w:LsdException Locked="false" Priority="46"
    Name="Grid Table 1 Light Accent 2"/>
    <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 2"/>
    <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 2"/>
    <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 2"/>
    <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 2"/>
    <w:LsdException Locked="false" Priority="51"
    Name="Grid Table 6 Colorful Accent 2"/>
    <w:LsdException Locked="false" Priority="52"
    Name="Grid Table 7 Colorful Accent 2"/>
    <w:LsdException Locked="false" Priority="46"
    Name="Grid Table 1 Light Accent 3"/>
    <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 3"/>
    <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 3"/>
    <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 3"/>
    <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 3"/>
    <w:LsdException Locked="false" Priority="51"
    Name="Grid Table 6 Colorful Accent 3"/>
    <w:LsdException Locked="false" Priority="52"
    Name="Grid Table 7 Colorful Accent 3"/>
    <w:LsdException Locked="false" Priority="46"
    Name="Grid Table 1 Light Accent 4"/>
    <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 4"/>
    <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 4"/>
    <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 4"/>
    <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 4"/>
    <w:LsdException Locked="false" Priority="51"
    Name="Grid Table 6 Colorful Accent 4"/>
    <w:LsdException Locked="false" Priority="52"
    Name="Grid Table 7 Colorful Accent 4"/>
    <w:LsdException Locked="false" Priority="46"
    Name="Grid Table 1 Light Accent 5"/>
    <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 5"/>
    <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 5"/>
    <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 5"/>
    <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 5"/>
    <w:LsdException Locked="false" Priority="51"
    Name="Grid Table 6 Colorful Accent 5"/>
    <w:LsdException Locked="false" Priority="52"
    Name="Grid Table 7 Colorful Accent 5"/>
    <w:LsdException Locked="false" Priority="46"
    Name="Grid Table 1 Light Accent 6"/>
    <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 6"/>
    <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 6"/>
    <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 6"/>
    <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 6"/>
    <w:LsdException Locked="false" Priority="51"
    Name="Grid Table 6 Colorful Accent 6"/>
    <w:LsdException Locked="false" Priority="52"
    Name="Grid Table 7 Colorful Accent 6"/>
    <w:LsdException Locked="false" Priority="46" Name="List Table 1 Light"/>
    <w:LsdException Locked="false" Priority="47" Name="List Table 2"/>
    <w:LsdException Locked="false" Priority="48" Name="List Table 3"/>
    <w:LsdException Locked="false" Priority="49" Name="List Table 4"/>
    <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark"/>
    <w:LsdException Locked="false" Priority="51" Name="List Table 6 Colorful"/>
    <w:LsdException Locked="false" Priority="52" Name="List Table 7 Colorful"/>
    <w:LsdException Locked="false" Priority="46"
    Name="List Table 1 Light Accent 1"/>
    <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 1"/>
    <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 1"/>
    <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 1"/>
    <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 1"/>
    <w:LsdException Locked="false" Priority="51"
    Name="List Table 6 Colorful Accent 1"/>
    <w:LsdException Locked="false" Priority="52"
    Name="List Table 7 Colorful Accent 1"/>
    <w:LsdException Locked="false" Priority="46"
    Name="List Table 1 Light Accent 2"/>
    <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 2"/>
    <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 2"/>
    <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 2"/>
    <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 2"/>
    <w:LsdException Locked="false" Priority="51"
    Name="List Table 6 Colorful Accent 2"/>
    <w:LsdException Locked="false" Priority="52"
    Name="List Table 7 Colorful Accent 2"/>
    <w:LsdException Locked="false" Priority="46"
    Name="List Table 1 Light Accent 3"/>
    <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 3"/>
    <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 3"/>
    <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 3"/>
    <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 3"/>
    <w:LsdException Locked="false" Priority="51"
    Name="List Table 6 Colorful Accent 3"/>
    <w:LsdException Locked="false" Priority="52"
    Name="List Table 7 Colorful Accent 3"/>
    <w:LsdException Locked="false" Priority="46"
    Name="List Table 1 Light Accent 4"/>
    <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 4"/>
    <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 4"/>
    <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 4"/>
    <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 4"/>
    <w:LsdException Locked="false" Priority="51"
    Name="List Table 6 Colorful Accent 4"/>
    <w:LsdException Locked="false" Priority="52"
    Name="List Table 7 Colorful Accent 4"/>
    <w:LsdException Locked="false" Priority="46"
    Name="List Table 1 Light Accent 5"/>
    <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 5"/>
    <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 5"/>
    <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 5"/>
    <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 5"/>
    <w:LsdException Locked="false" Priority="51"
    Name="List Table 6 Colorful Accent 5"/>
    <w:LsdException Locked="false" Priority="52"
    Name="List Table 7 Colorful Accent 5"/>
    <w:LsdException Locked="false" Priority="46"
    Name="List Table 1 Light Accent 6"/>
    <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 6"/>
    <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 6"/>
    <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 6"/>
    <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 6"/>
    <w:LsdException Locked="false" Priority="51"
    Name="List Table 6 Colorful Accent 6"/>
    <w:LsdException Locked="false" Priority="52"
    Name="List Table 7 Colorful Accent 6"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Mention"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Smart Hyperlink"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Hashtag"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Unresolved Mention"/>
    <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
    Name="Smart Link"/>
    </w:LatentStyles>
    </xml><![endif]-->
    <style>
    <!--
    /* Font Definitions */
    @font-face
        {font-family:"Cambria Math";
        panose-1:2 4 5 3 5 4 6 3 2 4;
        mso-font-charset:0;
        mso-generic-font-family:roman;
        mso-font-pitch:variable;
        mso-font-signature:3 0 0 0 1 0;}
    @font-face
        {font-family:Calibri;
        panose-1:2 15 5 2 2 2 4 3 2 4;
        mso-font-charset:0;
        mso-generic-font-family:swiss;
        mso-font-pitch:variable;
        mso-font-signature:-469750017 -1073732485 9 0 511 0;}
    @font-face
        {font-family:"Franklin Gothic Book";
        panose-1:2 11 5 3 2 1 2 2 2 4;
        mso-font-charset:0;
        mso-generic-font-family:swiss;
        mso-font-pitch:variable;
        mso-font-signature:647 0 0 0 159 0;}
    /* Style Definitions */
    p.MsoNormal, li.MsoNormal, div.MsoNormal
        {mso-style-unhide:no;
        mso-style-qformat:yes;
        mso-style-parent:"";
        margin:0in;
        mso-pagination:widow-orphan;
        font-size:11.0pt;
        font-family:"Calibri",sans-serif;
        mso-ascii-font-family:Calibri;
        mso-ascii-theme-font:minor-latin;
        mso-fareast-font-family:"Times New Roman";
        mso-fareast-theme-font:minor-fareast;
        mso-hansi-font-family:Calibri;
        mso-hansi-theme-font:minor-latin;
        mso-bidi-font-family:"Times New Roman";
        mso-bidi-theme-font:minor-bidi;}
    a:link, span.MsoHyperlink
        {mso-style-noshow:yes;
        mso-style-priority:99;
        mso-style-parent:"";
        color:#0563C1;
        text-decoration:underline;
        text-underline:single;}
    a:visited, span.MsoHyperlinkFollowed
        {mso-style-noshow:yes;
        mso-style-priority:99;
        color:#954F72;
        mso-themecolor:followedhyperlink;
        text-decoration:underline;
        text-underline:single;}
    .MsoChpDefault
        {mso-style-type:export-only;
        mso-default-props:yes;
        font-size:11.0pt;
        mso-ansi-font-size:11.0pt;
        mso-bidi-font-size:11.0pt;
        mso-ascii-font-family:Calibri;
        mso-ascii-theme-font:minor-latin;
        mso-fareast-font-family:"Times New Roman";
        mso-fareast-theme-font:minor-fareast;
        mso-hansi-font-family:Calibri;
        mso-hansi-theme-font:minor-latin;
        mso-bidi-font-family:"Times New Roman";
        mso-bidi-theme-font:minor-bidi;}
    @page WordSection1
        {size:8.5in 11.0in;
        margin:1.0in 1.0in 1.0in 1.0in;
        mso-header-margin:.5in;
        mso-footer-margin:.5in;
        mso-paper-source:0;}
    div.WordSection1
        {page:WordSection1;}
    -->
    </style>
    <!--[if gte mso 10]>
    <style>
    /* Style Definitions */
    table.MsoNormalTable
        {mso-style-name:"Table Normal";
        mso-tstyle-rowband-size:0;
        mso-tstyle-colband-size:0;
        mso-style-noshow:yes;
        mso-style-priority:99;
        mso-style-parent:"";
        mso-padding-alt:0in 5.4pt 0in 5.4pt;
        mso-para-margin:0in;
        mso-pagination:widow-orphan;
        font-size:11.0pt;
        font-family:"Calibri",sans-serif;
        mso-ascii-font-family:Calibri;
        mso-ascii-theme-font:minor-latin;
        mso-hansi-font-family:Calibri;
        mso-hansi-theme-font:minor-latin;
        mso-bidi-font-family:"Times New Roman";
        mso-bidi-theme-font:minor-bidi;}
    </style>
    <![endif]-->
    </head>

    <body lang=EN-US link="#0563C1" vlink="#954F72" style='tab-interval:.5in;
    word-wrap:break-word'>
    <div>
    <p>""" + getStaff[x][0] + r""",</p>
    <p>Today the Birmingham Business Journal released <a href="https://www.bizjournals.com/birmingham/news/2021/12/13/why-construction-and-real-estate-companies.html?ana=e_ae_native&j=26029992&senddate=2021-12-13">this article</a> entitled 'Why construction and real estate companies need dashboards.' The main points of the article are:<ul><li>Companies need timely reporting in an easily digestible format in a post-pandemic construction and real estate landscape</li><li>Dashboards provide immediate and accurate status updates on key performance indicators</li><li>Dashboards can implement historical data, real time metrics, and forecasting all in one place for better executive decision-making</li><li>Dashboards can aggregate data from multiple sources to easily manage finding and retaining top talent in an overly competitve market</li></ul></p>
    <p>I know you have """ + ('a couple' if getStaff[x][2] < 5 else 'quite a few') + r""" clients in this industry. In addition to a wealth of experience in dashboards, your <em>in-house</em> Business Intelligence team can provide dynamic automation to further streamline your clients' workforce. In fact, one of our very first engagements was automating 1095 forms for a BMSS construction client who had over 700 employees. Our process saved them hundreds of hours by eliminating the need to manually input the required codes for each employee for each month of the year. With Success Season right around the corner, do you have a few minutes in the next two weeks to chat with us about how we can help you provide your clients peace of mind with Business Intelligence?</p>
    <p>Thanks for getting back with us and we're look forward to working together on this,
    </div>
    <div class=WordSection1>

    <p class=MsoNormal><!--[if gte vml 1]><v:shapetype id="_x0000_t75" coordsize="21600,21600"
    o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f"
    stroked="f">
    <v:stroke joinstyle="miter"/>
    <v:formulas>
    <v:f eqn="if lineDrawn pixelLineWidth 0"/>
    <v:f eqn="sum @0 1 0"/>
    <v:f eqn="sum 0 0 @1"/>
    <v:f eqn="prod @2 1 2"/>
    <v:f eqn="prod @3 21600 pixelWidth"/>
    <v:f eqn="prod @3 21600 pixelHeight"/>
    <v:f eqn="sum @0 0 1"/>
    <v:f eqn="prod @6 1 2"/>
    <v:f eqn="prod @7 21600 pixelWidth"/>
    <v:f eqn="sum @8 21600 0"/>
    <v:f eqn="prod @7 21600 pixelHeight"/>
    <v:f eqn="sum @10 21600 0"/>
    </v:formulas>
    <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
    <o:lock v:ext="edit" aspectratio="t"/>
    </v:shapetype><v:shape id="_x0000_i1025" type="#_x0000_t75" alt="" style='width:132pt;
    height:49.5pt'>
    <v:imagedata src="C:\Users\jeremyshank\AppData\Roaming\Microsoft\Signatures\abacus_files/image001.jpg" o:href="cid:image004.jpg@01D6D3A0.A0905550"/>
    </v:shape><![endif]--><![if !vml]><img width=176 height=66
    src="C:\Users\jeremyshank\AppData\Roaming\Microsoft\Signatures\abacus_files/image001.jpg" style='height:.687in;width:1.833in' v:shapes="_x0000_i1025"><![endif]><span
    style='font-family:"Franklin Gothic Book",sans-serif'><o:p></o:p></span></p>

    <p class=MsoNormal><b>Jeremy Shank, LSM, LSPO | Scrum Master / Developer<o:p></o:p></b></p>

    <p class=MsoNormal><a href="mailto:jshank@abacustechnologies.com">jshank@abacustechnologies.com</a><u><span
    style='color:#0563C1'> <o:p></o:p></span></u></p>

    <p class=MsoNormal>p)205-443-5926 m)205-243-1212<o:p></o:p></p>

    <p class=MsoNormal><o:p>&nbsp;</o:p></p>

    <p class=MsoNormal>1121 Riverchase Office Rd, Birmingham, AL 35244<o:p></o:p></p>

    <p class=MsoNormal><a
    href="https://www.abacustechnologies.com/">www.abacustechnologies.com</a><u><span
    style='color:#0563C1'><o:p></o:p></span></u></p>

    <p class=MsoNormal><a
    href="https://www.linkedin.com/in/jpshank2/"><span
    style='color:windowtext;text-decoration:none;text-underline:none'><!--[if gte vml 1]><v:shape
    id="_x0000_i1026" type="#_x0000_t75" alt="download (2)" style='width:18.75pt;
    height:18.75pt'>
    <v:imagedata src="C:\Users\jeremyshank\AppData\Roaming\Microsoft\Signatures\abacus_files/image002.jpg" o:href="cid:image005.jpg@01D6D3A0.A0905550"/>
    </v:shape><![endif]--><![if !vml]><img border=0 width=25 height=25
    src="C:\Users\jeremyshank\AppData\Roaming\Microsoft\Signatures\abacus_files/image002.jpg" style='height:.263in;width:.263in'
    alt="download (2)" v:shapes="_x0000_i1026"><![endif]></span></a><a
    href="https://nam10.safelinks.protection.outlook.com/?url=https%3A%2F%2Ftwitter.com%2FAbacusTechAL&amp;data=04%7C01%7Ckmoore%40abacustechnologies.com%7C5ab946914545486085c508d8a1d39163%7Cfef92eeb36924b348da5822186e8ad26%7C0%7C0%7C637437277335835546%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C1000&amp;sdata=mcqhSqQlYdfhJ6taqYiGObdSN9JLLHX%2FmtIQ%2F7r47%2Fg%3D&amp;reserved=0"><span
    style='text-decoration:none;text-underline:none'><!--[if gte vml 1]><v:shape
    id="_x0000_i1027" type="#_x0000_t75" alt="download (1)" style='width:18.75pt;
    height:18.75pt'>
    <v:imagedata src="C:\Users\jeremyshank\AppData\Roaming\Microsoft\Signatures\abacus_files/image003.jpg" o:href="cid:image007.jpg@01D6D3A0.A0905550"/>
    </v:shape><![endif]--><![if !vml]><img border=0 width=25 height=25
    src="C:\Users\jeremyshank\AppData\Roaming\Microsoft\Signatures\abacus_files/image003.jpg" style='height:.263in;width:.263in'
    alt="download (1)" v:shapes="_x0000_i1027"><![endif]></span></a><a
    href="https://nam10.safelinks.protection.outlook.com/?url=https%3A%2F%2Fwww.facebook.com%2FAbacusTechnologies1%2F&amp;data=04%7C01%7Ckmoore%40abacustechnologies.com%7C5ab946914545486085c508d8a1d39163%7Cfef92eeb36924b348da5822186e8ad26%7C0%7C0%7C637437277335835546%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C1000&amp;sdata=aZpvP70zC1%2FMB%2BchcKFPf7GrKjh3xnc8T0JiW241h88%3D&amp;reserved=0"><span
    style='text-decoration:none;text-underline:none'><!--[if gte vml 1]><v:shape
    id="_x0000_i1028" type="#_x0000_t75" alt="download" style='width:18.75pt;
    height:18.75pt'>
    <v:imagedata src="C:\Users\jeremyshank\AppData\Roaming\Microsoft\Signatures\abacus_files/image004.jpg" o:href="cid:image008.jpg@01D6D3A0.A0905550"/>
    </v:shape><![endif]--><![if !vml]><img border=0 width=25 height=25
    src="C:\Users\jeremyshank\AppData\Roaming\Microsoft\Signatures\abacus_files/image004.jpg" style='height:.263in;width:.263in'
    alt=download v:shapes="_x0000_i1028"><![endif]></span></a><a
    href="https://nam10.safelinks.protection.outlook.com/?url=https%3A%2F%2Fwww.instagram.com%2Fabacus_technologies%2F&amp;data=04%7C01%7Ckmoore%40abacustechnologies.com%7C5ab946914545486085c508d8a1d39163%7Cfef92eeb36924b348da5822186e8ad26%7C0%7C0%7C637437277335845545%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C1000&amp;sdata=RtBykrKMWVMTjY1Ygs8Zxp%2BHns2IzACO8K0nYVnyLVI%3D&amp;reserved=0"><span
    style='text-decoration:none;text-underline:none'><!--[if gte vml 1]><v:shape
    id="_x0000_i1029" type="#_x0000_t75" alt="cropped-IG_Glyph_Fill" style='width:21pt;
    height:21pt'>
    <v:imagedata src="C:\Users\jeremyshank\AppData\Roaming\Microsoft\Signatures\abacus_files/image005.png" o:href="cid:image009.png@01D6D3A0.A0905550"/>
    </v:shape><![endif]--><![if !vml]><img border=0 width=28 height=28
    src="C:\Users\jeremyshank\AppData\Roaming\Microsoft\Signatures\abacus_files/image005.png" style='height:.291in;width:.291in'
    alt="cropped-IG_Glyph_Fill" v:shapes="_x0000_i1029"><![endif]></span></a><a
    href="https://nam10.safelinks.protection.outlook.com/?url=https%3A%2F%2Fwww.youtube.com%2Fchannel%2FUCqNQSCg5HWqmAISadcIFfOQ&amp;data=04%7C01%7Ckmoore%40abacustechnologies.com%7C5ab946914545486085c508d8a1d39163%7Cfef92eeb36924b348da5822186e8ad26%7C0%7C0%7C637437277335845545%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C1000&amp;sdata=U0hwLMz5SHDeJOZ1TXeoPqBhjOPzrUIYRANDqTprTuQ%3D&amp;reserved=0"><span
    style='text-decoration:none;text-underline:none'><!--[if gte vml 1]><v:shape
    id="_x0000_i1030" type="#_x0000_t75" alt="2000px-YouTube_social_white_square_2017"
    style='width:26.25pt;height:19.5pt'>
    <v:imagedata src="C:\Users\jeremyshank\AppData\Roaming\Microsoft\Signatures\abacus_files/image006.png" o:href="cid:image010.png@01D6D3A0.A0905550"/>
    </v:shape><![endif]--><![if !vml]><img border=0 width=35 height=26
    src="C:\Users\jeremyshank\AppData\Roaming\Microsoft\Signatures\abacus_files/image006.png" style='height:.27in;width:.368in'
    alt="2000px-YouTube_social_white_square_2017" v:shapes="_x0000_i1030"><![endif]></span></a><a
    href="https://github.com/jpshank2"><span
    style='color:windowtext;text-decoration:none;text-underline:none'><!--[if gte vml 1]><v:shape
    id="_x0000_i1026" type="#_x0000_t75" alt="download (2)" style='width:18.75pt;
    height:18.75pt'>
    <v:imagedata src="C:\Users\jeremyshank\AppData\Roaming\Microsoft\Signatures\abacus_files/image007.png" o:href="cid:image005.jpg@01D6D3A0.A0905550"/>
    </v:shape><![endif]--><![if !vml]><img border=0 width=25 height=25
    src="C:\Users\jeremyshank\AppData\Roaming\Microsoft\Signatures\abacus_files/image007.png" style='height:.263in;width:.263in'
    alt="download (2)" v:shapes="_x0000_i1026"><![endif]></span></a><u><span
    style='color:#0563C1'><o:p></o:p></span></u></p>

    <p class=MsoNormal><o:p>&nbsp;</o:p></p>

    </div>

    </body>"""

    mail.Send()
