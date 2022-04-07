' --------------------------------------------------------------------------
'  File:         backup.vbs
'  Purpose:  	used to remove the last backslash 
'  Date:         07, June, 2017
'  Description: Used to backup application's extensions and progIDs
'  wscript.exe "backup.vbs"
' ----------------------------------------------------------------------------
On error resume next
const HKEY_CLASSES_ROOT = &H80000000
path1 = session.property("CustomActionData")
path = path1 & "backup\" 
Set osh = CreateObject("WScript.Shell")

osh.run "cmd /c mkdir " & """" & path & """", 0, True
 
cmd = "reg export " & """" & "HKCR\.dwfx" & """" & " " & chr(34) & path & "dwfx.reg" & chr(34)
osh.Run cmd, 0, True