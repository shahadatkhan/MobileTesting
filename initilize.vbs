'Author Shahadat Khan - skhan@primetgi.com

ClientName = "RedBus" 		'Change this for each project implementation
SendNotificationsForTestCases = "10" 'Send emails after executing so many test cases

Set wshShell = CreateObject( "WScript.Shell" )
Set wshSystemEnv = wshShell.Environment( "SYSTEM" )
'Set the environment variable
wshSystemEnv( "SeleniumEmailSentYN" ) = "N"
wshSystemEnv( "ClientName" ) = ClientName
FileName = ClientName & "_" & funGetTimeStamp()
wshSystemEnv( "FileName" ) = FileName
wshSystemEnv( "FileNameCounter" ) = "0"
wshSystemEnv( "SendNotificationsForTestCases" ) = SendNotificationsForTestCases
wshSystemEnv( "TestCaseCounter" ) = "0"
wshSystemEnv( "TestCaseIDsList" ) = ""
wshSystemEnv( "Total" ) = "0"
wshSystemEnv( "Pass" ) = "0"
wshSystemEnv( "Fail" ) = "0"
wshSystemEnv( "BatchTotal" ) = "0"
wshSystemEnv( "BatchPass" ) = "0"
wshSystemEnv( "BatchFail" ) = "0"
Set wshSystemEnv = Nothing
Set wshShell     = Nothing
'******************************************************************************************
'Used to empty test results and screenshots
set fso=CreateObject("Scripting.FileSystemObject")
WorkingDirectory = fso.GetParentFolderName(Wscript.ScriptFullName)

MainFolder = "D:\\QEdge_Selenium\\eclipse_java\\eclipse_scripts_RC\\MobileTesting"
ModuleName = ClientName

TestResultsFolder = WorkingDirectory&"\Test_Reports\"
ScreenshotsFolder = WorkingDirectory&"\ScreenShots\"

'Don't change beyond this code **************************************************************

BaseFolder = MainFolder&ModuleName&"\"

'Check backup folder exist, else create new
If  Not fso.FolderExists(MainFolder) Then
   fso.CreateFolder (MainFolder)
End If
If  Not fso.FolderExists(MainFolder&ModuleName&"\") Then
   fso.CreateFolder (MainFolder&ModuleName&"\")
End If
'Check screenshots, test reports folder exist, else create new
If  Not fso.FolderExists(TestResultsFolder) Then
   fso.CreateFolder (TestResultsFolder)
End If
If  Not fso.FolderExists(ScreenshotsFolder) Then
   fso.CreateFolder (ScreenshotsFolder)
End If


'Check if folders are empty
set folder1 = fso.getFolder(TestResultsFolder)
Set folder1Sub = folder1.SubFolders  

set folder2 = fso.getFolder(ScreenshotsFolder)

'new folder name with time stamp
NewFolder = BaseFolder & FileName & "__OldData"

'Count>0 there are files inside the folders
'Msgbox folder1Sub.Count
'msgbox folder2.files.Count
If (folder1Sub.Count>1 or folder2.files.Count>0) then
	fso.CreateFolder (NewFolder)
	'Msgbox "New folder created"
End if

If (folder1Sub.Count>1) then
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


'MsgBox"Completed"

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

