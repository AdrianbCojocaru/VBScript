' --------------------------------------------------------------------------
'  File:         restoreext.vbs
'  Purpose:  used to restore two extension at uninstall
'  Date:          19 July 2017
'  Description:
'  Usage: example usage:
'  wscript.exe "restoreext.vbs"
' ----------------------------------------------------------------------------
quote = """"

Set WshShell = CreateObject("WScript.Shell")
set fso = createObject("scripting.FileSystemObject")
inst = session.property("CustomActionData")
inst1 = Left(inst, (Len(inst) - 1))


if fso.fileexists(inst & "shtml.reg") then
	cmd = "reg import" & " " & quote & inst & "\shtml.reg" & quote
	WshShell.Run cmd, 0, True 

	fso.deletefile inst & "\shtml.reg"
end if

fso.deletefolder inst1