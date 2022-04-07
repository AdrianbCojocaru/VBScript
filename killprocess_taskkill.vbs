' --------------------------------------------------------------------------
'  File:         killprocess1.vbs
'  Purpose:  used to terminate application's process 
'  Date:          28,June,2017
'  Description:
'  Usage: example usage:
'  wscript.exe "killprocess.vbs"
' --------------------------------------------------------------------------
Dim oSH 
Dim shellCommand, returnVal

Set oSH = CreateObject("WScript.Shell")

shellCommand = "cmd.exe /c taskkill /im " & Chr(34) & "GoogleUpdate.exe" & Chr(34) & " /f"
returnVal = oSH.Run (shellCommand, 0, true)

set oSH = nothing