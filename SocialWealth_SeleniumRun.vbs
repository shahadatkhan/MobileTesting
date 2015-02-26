'Schedule this script in the windows scheduler, it will automatically invoke ANT from cmd

Minutes = 200           'Minutes. Force stop after this time

set fso=CreateObject("Scripting.FileSystemObject")
WorkingDirectory = fso.GetParentFolderName(Wscript.ScriptFullName)
set fso=Nothing

Set objShell = WScript.CreateObject("WScript.Shell") 
objShell.CurrentDirectory = WorkingDirectory


'************************************************************************************
'Email reset N
Set wshShell = CreateObject( "WScript.Shell" )
Set wshSystemEnv = wshShell.Environment( "SYSTEM" )
'Set the environment variable
wshSystemEnv( "SeleniumEmailSentYN" ) = "N"
Set wshSystemEnv = Nothing
Set wshShell     = Nothing

objShell.Run "ant run,1,true" 
WScript.Sleep 7000

'Wait till the process completed *****
Start1 = Now()
End1 = DateAdd("n",Minutes,Start1)
set svc=getobject("winmgmts:root\cimv2")
sQuery="select * from win32_process where name='cmd.exe'"
set cproc=svc.execquery(sQuery)
iniproc=cproc.count    'it can be more than 1
Do While iniproc > 0
    wscript.sleep 20000
    set svc=getobject("winmgmts:root\cimv2")
    sQuery="select * from win32_process where name='cmd.exe'"
    set cproc=svc.execquery(sQuery)
    iniproc=cproc.count
    If DateDiff("n",Now(),End1) <= 0 then
    	KillCmdProcess()
    End if
Loop
set cproc=nothing
set svc=nothing
WScript.Sleep 30000

'Check Email sent
Set wshShell = CreateObject( "WScript.Shell" )
Set wshSystemEnv = wshShell.Environment( "SYSTEM" )
EmailSentYN =  wshSystemEnv( "SeleniumEmailSentYN" )
Set wshSystemEnv = Nothing
Set wshShell     = Nothing

' Send email notification
If EmailSentYN = "N" then
  objShell.Run "sendemail.vbs e" 
  WScript.Sleep 30000
End If
'If EmailSentYN = "Y" then
  'objShell.Run "sendemail.vbs c" 
  'WScript.Sleep 30000
'End If
'***************************************************************************************

Set objShell = Nothing


Function KillCmdProcess()
	strComputer = "."
	
	strProcessToKill = "java.exe" 
	SET objWMIService = GETOBJECT("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" _ 
		& strComputer & "\root\cimv2") 

	SET colProcess = objWMIService.ExecQuery _
		("Select * from Win32_Process Where Name = '" & strProcessToKill & "'")

	count = 0
	FOR EACH objProcess in colProcess
		objProcess.Terminate()
		count = count + 1
	NEXT 
	
	strProcessToKill = "cmd.exe" 
	SET objWMIService = GETOBJECT("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" _ 
		& strComputer & "\root\cimv2") 

	SET colProcess = objWMIService.ExecQuery _
		("Select * from Win32_Process Where Name = '" & strProcessToKill & "'")

	count = 0
	FOR EACH objProcess in colProcess
		objProcess.Terminate()
		count = count + 1
	NEXT 	
	
	strProcessToKill = "iexplore.exe" 
	SET objWMIService = GETOBJECT("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" _ 
		& strComputer & "\root\cimv2") 

	SET colProcess = objWMIService.ExecQuery _
		("Select * from Win32_Process Where Name = '" & strProcessToKill & "'")

	count = 0
	FOR EACH objProcess in colProcess
		objProcess.Terminate()
		count = count + 1
	NEXT 	
	
	SET objWMIService = Nothing
	SET colProcess = Nothing
	wscript.sleep 5000
	
	
	
End Function


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