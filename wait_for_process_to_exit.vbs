Dim CommandLine

CommandLine = "C:\Program Files\Internet Explorer\iexplore.exe"

IsProcessRunning(CommandLine)

Function IsProcessRunning(strCommandLine)
Dim objWMIService, objProcess, colProcess, Linie, strComputer, strList
strComputer = "." 
IsProcessRunning = False
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
Set colProcess = objWMIService.ExecQuery _
      ("Select * from Win32_Process")

      For Each objProcess in colProcess
      if (objProcess.CommandLine <> "") Then
            Line = objProcess.CommandLine
            if (InStr(Line,strCommandLine)) Then
                  IsProcessRunning = True
				  Wscript.Echo objProcess.Name
            end if

      end if
      Next
	  while IsProcessRunning = True
		IsProcessRunning(strCommandLine)
	  Wend

Set objWMIService = Nothing
Set colProcess = Nothing

End Function