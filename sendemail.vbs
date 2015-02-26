'Arguments
Set objArgs = WScript.Arguments
Arg1 = objArgs(0)
Set objArgs = Nothing

'Email Invoked Y
Set wshShell = CreateObject( "WScript.Shell" )
Set wshSystemEnv = wshShell.Environment( "SYSTEM" )
'Set the environment variable
wshSystemEnv( "SeleniumEmailSentYN" ) = "Y"
ClientName = wshSystemEnv( "ClientName" )
TestCaseIDsList = wshSystemEnv( "TestCaseIDsList" )
FileNameCounter = wshSystemEnv( "FileNameCounter" )
FileName = wshSystemEnv( "FileName" )
Total = cInt(wshSystemEnv( "Total" ))
Pass = cInt(wshSystemEnv( "Pass" ))
Fail = cInt(wshSystemEnv( "Fail" ))
BatchTotal = cInt(wshSystemEnv( "BatchTotal" ))
BatchPass = cInt(wshSystemEnv( "BatchPass" ))
BatchFail = cInt(wshSystemEnv( "BatchFail" ))

If ClientName = Empty Then
	ClientName = "NoClientAssigned"
End If

If FileName = Empty Then
	FileName = ClientName & "_" & funGetTimeStamp()
	wshSystemEnv( "FileName" ) = FileName
	FileName = FileName & "__" & FileNameCounter
Else
	FileName = FileName & "__" & FileNameCounter
End If

If Arg1 = "c" then
	'MsgBox "Inside Wait"
	WScript.Sleep 30000 ' For TestNG to create reports
	FileName = FileName & "_Final"
	FileNameCounter = FileNameCounter & "_Final"
Else
End If

set fso=CreateObject("Scripting.FileSystemObject")
WorkingDirectory = fso.GetParentFolderName(Wscript.ScriptFullName)
'######################### Results Excel file Reading and Sending Email #################################
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(WorkingDirectory&"\TestSetup\TestSetup.xls")
Set objSheet = objExcel.ActiveWorkbook.Worksheets("TestSetup")
ResultFile = objSheet.Cells(5,2).Value
objExcel.Quit

Dim TestCaseName()
Dim DataSetNo()
Dim TestStatus()
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(WorkingDirectory&"\Result\"& ResultFile &"_Result.xls")
Set objSheet = objExcel.ActiveWorkbook.Worksheets("TestScript")
intCol = 17
Count=0
ReDim TestCaseName(1)
ReDim DataSetNo(1)
Do While objSheet.Cells(1, intCol).Value <> ""
If objSheet.Cells(3, intCol).Value <> "SKIP_ALL" Then
TestCaseName(Count) = objSheet.Cells(1, intCol).Value
DataSetNo(Count) = objSheet.Cells(2, intCol).Value
Count = Count + 1
ReDim Preserve TestCaseName(Count+1)
ReDim Preserve DataSetNo(Count+1)
End If
intcol = intcol + 1
Loop
Set objSheet = objExcel.ActiveWorkbook.Worksheets("Results")
Flag = 1
intRow =2
TestPassCount = 0
TestFailCount = 0
DataLength = UBound(DataSetNo)
ReDim TestStatus(1)
for i = 0 to DataLength - 2	
	Do Until objSheet.Cells(intRow, 1).Value = "" 
	'If CInt(objSheet.Cells(intRow, 1).Value) = CInt(DataSetNo(i)) and objSheet.Cells(intRow, 11).Value = "Fail" Then
	If objSheet.Cells(intRow, 11).Value = "Fail" Then
		TestFailCount = TestFailCount + 1
		TestStatus(i) = "<td style='border:1px solid black;'bgcolor=#ff0000>FAIL</td>"
		Flag = 0
		intRow = intRow +1
		Exit Do
	Else
		intRow = intRow +1
	End If
	
	'If CInt(objSheet.Cells(intRow, 1).Value) <> CInt(DataSetNo(i)) then
	'Exit Do
	'End If	
	Loop
	
	If Flag = 1 then 
		TestStatus(i) = "<td style='border:1px solid black;'bgcolor=#008000>PASS</td>"
		TestPassCount = TestPassCount + 1
	End If
	ReDim Preserve TestStatus(i+2)
Next

objExcel.Quit
Set objExcel = nothing 

'######################### End of Results Excel file Reading and Sending Email ###########################
'MsgBox WorkingDirectory

Set objEmailMessage = CreateObject("CDO.Message")
objEmailMessage.From = "skhan@primetgi.com"


'*****Be careful with notifications********************************************************


objEmailMessage.To = "skhan@primetgi.com"

objEmailMessage.Cc = "skhan@primetgi.com; skhan@primetgi.com"

'objEmailMessage.Cc = "skhan@primetgi.com"

'*****Be careful with notifications***********************************************************


objEmailMessage.Subject = ClientName & " - TestResults"

sMessage = sMessage & "Below are the executed test cases ID's<BR>" & VBCR 
sMessage = sMessage & "All TC Executed  " & "Total:"& Count & "   " & "Pass:" & TestPassCount & "   " & "Fail:" & TestFailCount & "<BR>" & VBCR
sMessage = sMessage & VBCR & VBCR
sMessage = sMessage & "<P>" & VBCR
'sMessage = sMessage & "This Batch:" & FileNameCounter & "   " & "Total:"& BatchTotal & "   " & "Pass:" & BatchPass & "   " & "Fail:" & BatchFail
sMessage = sMessage & "<table style='border:1px solid black;border-collapse:collapse;'><tr><th style='border:1px solid black;'>S.no</th><th style='border:1px solid black;'> Test Case Name</th><th style='border:1px solid black;'>Status</th><th style='border:1px solid black;'>Comments</th></tr>"& VBCR

for i =0 to UBound(TestCaseName) - 2
sMessage = sMessage & "<tr><td style='border:1px solid black;'>"&i+1&"</td><td style='border:1px solid black;'>" & TestCaseName(i) & "</td>" & TestStatus(i) & "</td><td style='border:1px solid black;'></td></tr>" & VBCR
next
sMessage = sMessage & "</table>" & VBCR
If Arg1 = "e" then
	sMessage =  VBCR & VBCR & sMessage & "Test not executed, some thing gone wrong!" & VBCR & VBCR
Else
End If
sMessage = sMessage & VBCR & VBCR
sMessage = sMessage & "<p>" & VBCR
sMessage = sMessage & "Attachments contain TestNG report in HTML, Test log file and Failed test case screen shots.<br>" & VBCR
sMessage = sMessage & "All the test cases are broken into groups, due to some technical issues.<br>" & VBCR
sMessage = sMessage & "You will receive multiple email notification for the same client.<br>" & VBCR
sMessage = sMessage & "Executed System: " & IPAddress() & "   " & Now() & vbLf
sMessage = sMessage & "<p>" & VBCR
sMessage = sMessage & "Regards,<br>" & vbLf
sMessage = sMessage & "QA Team<br>" & VBCR

objEmailMessage.HTMLBody = sMessage

objEmailMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
objEmailMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objEmailMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmailMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
objEmailMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
objEmailMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 20
objEmailMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "venkatkandula09@gmail.com"
objEmailMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "9494084346"

'Attach emailable report.html
'objStartFolder = WorkingDirectory & "\Test_Reports\" 
'ShowSubfolders fso.GetFolder(objStartFolder) ' Call sub 
'Attach Screenshots
Set folder = fso.GetFolder(WorkingDirectory & "\Screenshots\")
Set files = folder.Files
For each folderIdx In files
	objEmailMessage.AddAttachment(folderIdx)
Next
objEmailMessage.AddAttachment WorkingDirectory&"\Result\"& ResultFile &"_Result.xls"
objEmailMessage.Configuration.Fields.Update

On Error Resume Next
objEmailMessage.Send      
         
If  err.number = 0 Then
   On Error Goto 0
   Set oNotepad = fso.createtextfile(WorkingDirectory & "\Screenshots\" & "EmailSentSuccessfully" & funGetTimeStamp()& ".txt")
   oNotepad.writeline("Email Sent Successfully")
   oNotepad.writeline funGetTimeStamp()
   oNotepad.Close
Else
   errNum = err.number
   errDesc = err.Description
   On Error Goto 0
   Set oNotepad = fso.createtextfile(WorkingDirectory & "\Screenshots\" & "EmailNOTSent" & funGetTimeStamp()& ".txt")
   oNotepad.writeline("EmailNOTSent")
   oNotepad.writeline("ErrorNumber:" & errNum)
   oNotepad.writeline("ErrorDesc:" & errDesc)
   oNotepad.writeline funGetTimeStamp()
   oNotepad.Close
End If

Set objEmailMessage = nothing
Erase TestCaseName
Erase DataSetNo
Erase TestStatus

'***********************************************************************

MainFolder = "D:\QEdge_Selenium\eclipse_java\eclipse_scripts_RC\MobileTesting"
ModuleName = ClientName

TestResultsFolder = WorkingDirectory&"\Test_Reports\"
ScreenshotsFolder = WorkingDirectory&"\ScreenShots\"

BaseFolder = MainFolder&ModuleName&"\"

'Check base folder exist, else create new
If  Not fso.FolderExists(MainFolder) Then
   fso.CreateFolder (MainFolder)
End If

If  Not fso.FolderExists(MainFolder&ModuleName&"\") Then
   fso.CreateFolder (MainFolder&ModuleName&"\")
End If

'Check if folders are empty
set folder1 = fso.getFolder(TestResultsFolder)
Set folder1Sub = folder1.SubFolders  

set folder2 = fso.getFolder(ScreenshotsFolder)

'new folder name with time stamp
'NewFolder = BaseFolder & funGetTimeStamp() & "__" & FileNameCounter
NewFolder = BaseFolder & FileName 

'Count>0 there are files inside the folders
If (folder1Sub.Count>0 or folder2.files.Count>0) then
	fso.CreateFolder (NewFolder)
End if

If (folder1Sub.Count>0) then
	Set objShell = CreateObject("Shell.Application")
	Set objFolder = objShell.NameSpace(NewFolder) 
	objFolder.CopyHere TestResultsFolder 
	Set objFolder = Nothing
	Set objShell = Nothing
End if

If (folder2.files.Count>0) then
	Set objShell = CreateObject("Shell.Application")
	Set objFolder = objShell.NameSpace(NewFolder) 
	objFolder.CopyHere ScreenshotsFolder 
	Set objFolder = Nothing
	Set objShell = Nothing
End if

For Each f1 in folder1Sub
	f1.delete
Next

fso.DeleteFile(ScreenshotsFolder & "*.*")

wshSystemEnv( "TestCaseIDsList" ) = ""

wshSystemEnv( "BatchTotal" ) = 0
wshSystemEnv( "BatchPass" ) = 0
wshSystemEnv( "BatchFail" ) = 0

'MsgBox "Email Sent"

set fso = Nothing
Set wshSystemEnv = Nothing
Set wshShell     = Nothing



'**************************************************************************************
Sub ShowSubFolders(Folder)
    For Each Subfolder in Folder.SubFolders
        'Wscript.Echo Subfolder.Path
        Set objFolder = fso.GetFolder(Subfolder.Path)
        Set colFiles = objFolder.Files
        For Each objFile in colFiles
			'Wscript.Echo objFile.Name
			if objFile.Name = "emailable-report.html" then
				objEmailMessage.AddAttachment Subfolder.Path & "\" & objFile.Name
			End if
        Next
        'Wscript.Echo
        ShowSubFolders Subfolder
    Next
End Sub 
'***************************************************************************************
Function funGetTimeStamp()
		sDateTIme = Now()
		
		iDate = Datepart("d",sDateTime)
		iLen = Len(iDate)
		If iLen = 1 Then
				iDate = "0" & iDate
		End If
		
		sMonth=  mid(MonthName(Datepart("m",sDateTime)),1,3)
		
		iYear = Datepart("yyyy",sDateTime)
		
		iHour = Datepart("h",sDateTime)
		iLen = Len(iHour)
		If iLen = 1 Then
				iHour = "0" & iHour
		End If
		
		iMinute = Datepart("n",sDateTime)
		iLen = Len(iMinute)
		If iLen = 1 Then
				iMinute = "0" & iMinute
		End If
		
		iSec = Datepart("s",sDateTime)
		iLen = Len(iSec)
		If iLen = 1 Then
				iSec = "0" & iSec
		End If
		
		
		funGetTimeStamp =  sMonth & "_" &  iDate & "_" & iYear & "_" & iHour & "_" & iMinute & "_" & iSec 
	 
End Function
'*************************************************************************************
Function IPAddress()
				strQuery = "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE MACAddress > ''"
				Set goWMIService = GetObject( "winmgmts://./root/CIMV2" )
				Set colItems = goWMIService.ExecQuery( strQuery, "WQL", 48 )
				For Each objItem In colItems
					If IsArray( objItem.IPAddress ) Then
						If UBound( objItem.IPAddress ) = 0 Then
							strIP = objItem.IPAddress(0)
						Else
							strIP = Join( objItem.IPAddress, "," )
						End If
					End If
				Next
				Set colItems      = Nothing
                IPAddress = "IP_" & strIP
End Function



