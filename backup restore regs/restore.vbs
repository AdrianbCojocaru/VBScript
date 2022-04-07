' --------------------------------------------------------------------------
'  File:         restore.vbs
'  Purpose:  	used to remove the last backslash 
'  Date:         07, June, 2017
'  Description: Used to restore application's extensions and progIDs
'  wscript.exe "restore.vbs"
' ----------------------------------------------------------------------------
On error resume next
const HKEY_CLASSE_ROOT = &H80000000
path1 = session.property("CustomActionData")
path = path1 & "backup\" 
Set oSH = CreateObject("WScript.Shell")
  
osh.Run "regedit /s " & """" & path & "dwfx.reg" & """", 0, True

osh.Run "cmd /c rmdir /s /q " & """" & path & """", 0, True