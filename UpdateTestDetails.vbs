'Arguments
Set objArgs = WScript.Arguments
Arg1 = objArgs(0)
Arg2 = objArgs(1)
Set objArgs = Nothing

'WScript.Sleep 3000 

Set wshShell = CreateObject( "WScript.Shell" )
Set wshSystemEnv = wshShell.Environment( "SYSTEM" )
FileName = wshSystemEnv( "FileName" )
FileNameCounter = wshSystemEnv( "FileNameCounter" )
ClientName = wshSystemEnv( "ClientName" )
SendNotificationsForTestCases = wshSystemEnv( "SendNotificationsForTestCases" )
TestCaseCounter = wshSystemEnv( "TestCaseCounter" )
TestCaseIDsList = wshSystemEnv( "TestCaseIDsList" )
Total = cInt(wshSystemEnv( "Total" ))
Pass = cInt(wshSystemEnv( "Pass" ))
Fail = cInt(wshSystemEnv( "Fail" ))
BatchTotal = cInt(wshSystemEnv( "BatchTotal" ))
BatchPass = cInt(wshSystemEnv( "BatchPass" ))
BatchFail = cInt(wshSystemEnv( "BatchFail" ))

Total = Total + 1
BatchTotal = BatchTotal + 1
If Instr(1,Arg1,"PASS",1)>0 Then
	Pass = Pass + 1
	BatchPass = BatchPass + 1
Else
	Fail = Fail + 1 
	BatchFail = BatchFail + 1
End If

wshSystemEnv( "Total" ) = Total
wshSystemEnv( "Pass" ) = Pass
wshSystemEnv( "Fail" ) = Fail

wshSystemEnv( "BatchTotal" ) = BatchTotal
wshSystemEnv( "BatchPass" ) = BatchPass
wshSystemEnv( "BatchFail" ) = BatchFail

'Msgbox "SendNotificationsForTestCases:" & SendNotificationsForTestCases & "  " & "TestCaseCounter:"& TestCaseCounter & "  " & "TestCaseIDsList" & TestCaseIDsList & " FileNameCounter:" & FileNameCounter

set fso=CreateObject("Scripting.FileSystemObject")
WorkingDirectory = fso.GetParentFolderName(Wscript.ScriptFullName)
'MsgBox WorkingDirectory

ScreenshotsFolder = WorkingDirectory&"\ScreenShots\"
If  Not fso.FolderExists(ScreenshotsFolder) Then
   fso.CreateFolder (ScreenshotsFolder)
End If

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

FileFound = 0
Set folder = fso.GetFolder(ScreenshotsFolder)
Set files = folder.Files
For each folderIdx In files
	'MsgBox folderIdx.Name
	'MsgBox FileName&".txt"
	If Ucase(folderIdx.Name) = Ucase(FileName&".txt") Then
			FileFound = 1
			'MsgBox "FileExist"
	End If
Next

If FileFound = 0 Then
	Set oNotepad = fso.createtextfile(ScreenshotsFolder & FileName & ".txt")
Else
	Const ForAppending = 8
	Set oFS=createobject ("scripting.filesystemobject")
	Set oNotepad=oFS.OpenTextFile (ScreenshotsFolder & FileName & ".txt", ForAppending, True)
End If

oNotepad.writeline("##################################################################")
oNotepad.writeline(Arg1 & "-----" &funGetTimeStamp() )
oNotepad.writeline("##################################################################")
ArrString = Split(Arg2,"#")
for each line in ArrString
   	oNotepad.writeline(line)
next
oNotepad.writeline
oNotepad.writeline "All TC Executed  " & "Total:"& Total & "   " & "Pass:" & Pass & "   " & "Fail:" & Fail
oNotepad.writeline
oNotepad.writeline "This Batch:" & FileNameCounter & "   " & "Total:"& BatchTotal & "   " & "Pass:" & BatchPass & "   " & "Fail:" & BatchFail
oNotepad.writeline
oNotepad.Close
Set oNotepad = Nothing

If (cint(SendNotificationsForTestCases) = 0 or SendNotificationsForTestCases = Empty) then
	'MsgBox "Then"
	TestCaseIDsList = TestCaseIDsList & "---" & Arg1
	wshSystemEnv( "TestCaseIDsList" ) = TestCaseIDsList
	IntTestCaseCounter = cint(TestCaseCounter) + 1
	wshSystemEnv( "TestCaseCounter" ) = IntTestCaseCounter
Else
	'MsgBox "Else"
	TestCaseIDsList = TestCaseIDsList & "---" & Arg1
	wshSystemEnv( "TestCaseIDsList" ) = TestCaseIDsList
	IntTestCaseCounter = cint(TestCaseCounter) + 1
	wshSystemEnv( "TestCaseCounter" ) = IntTestCaseCounter
	If IntTestCaseCounter >= cint(SendNotificationsForTestCases) then
		'MsgBox "Else1"
		WorkingDirectory = fso.GetParentFolderName(Wscript.ScriptFullName)
		strFileName =  WorkingDirectory & "\sendemail.vbs" & " " & "i"
		'MsgBox strFileName
		wshShell.Run "wscript " & strFileName, 1, True
		WScript.Sleep 5000
		wshSystemEnv( "TestCaseCounter" ) = "0"		
		wshSystemEnv( "TestCaseIDsList" ) = ""
		FileNameCounter = cint(FileNameCounter) + 1	
		wshSystemEnv( "FileNameCounter" ) = FileNameCounter	
		'Msgbox "SendNotificationsForTestCases:" & SendNotificationsForTestCases & "  " & "TestCaseCounter:0"& "TestCaseIDsList--" & " FileNameCounter:" & FileNameCounter
	Else
		'Msgbox "SendNotificationsForTestCases:" & SendNotificationsForTestCases & "  " & "TestCaseCounter:"& IntTestCaseCounter & "  " & "TestCaseIDsList" & TestCaseIDsList & "  FileNameCounter:" & FileNameCounter
	End If
	
End If

'MsgBox"Updated"
Set wshSystemEnv = Nothing
set fso = Nothing

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
 