' --------------------------------------------------------------------------
'  Copyright (C) 2013
'  Description: Used to add active sertup registries for selfhealing
' --------------------------------------------------------------------------
Dim strProductCode, CmdLine
Set WshShell = CreateObject("WScript.Shell")

strProductCode = Session.Property("ProductCode")
DateTime = Session.Property("Date") & Session.Property("Time")
CmdLine = "msiexec.exe /fus " & strProductCode

WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Active Setup\Installed Components\" & strProductCode & DateTime & "\StubPath", CmdLine, "REG_SZ"
WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Active Setup\Installed Components\" & strProductCode & DateTime & "\IsInstalled", 1, "REG_DWORD"
