' ---------------------------------------------------------------------
'  AC - MS GmbH
'  Description:		used to terminate application's process.
' ---------------------------------------------------------------------
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colProcessList = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_Process WHERE Name = 'EPWD.exe'")
For Each objProcess in colProcessList
    objProcess.Terminate()
Next


