' --------------------------------------------------------------------------
'  File:         SetCurrDir.vbs
'  Purpose:  	 used to get the current directory 
'  Date:         06,June,2014
'  Description:
'  Usage: example usage:
'  wscript.exe "SetCurrDir.vbs"
' ----------------------------------------------------------------------------
Dim WshShell, strCurrDir
Set WshShell = CreateObject("WScript.Shell")
Session.Property("CURRDIR") = WshShell.CurrentDirectory
Set WshShell = Nothing