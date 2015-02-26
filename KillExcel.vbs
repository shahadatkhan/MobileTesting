
Set objShell = Nothing

call KillCmdProcess()

Function KillCmdProcess()
	strComputer = "."
	strProcessToKill = "Excel.exe" 

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
	
End Function