' --------------------------------------------------------------------------
'  File:         backupext.vbs
'  Purpose:  used to backup two extension
'  Date:          19 July 2017
'  Description:
'  Usage: example usage:
'  wscript.exe "backupext.vbs"
' ----------------------------------------------------------------------------
quote = """" 
Dim objFSO, objFolder, objShell, strDirectory
Set WshShell = CreateObject("WScript.Shell")
set fso = createObject("scripting.FileSystemObject")
Set objFSO = CreateObject("Scripting.FileSystemObject")

strDirectory = session.property("CustomActionData")


If Not objFSO.FolderExists(strDirectory) Then
   objFSO.CreateFolder(strDirectory)
End If

                                            
if not fso.fileexists(strDirectory & "shtml.reg") then
	cmd = "reg export " & """" & "HKCR\.shtml" & """" & " " & quote & strDirectory & "\shtml.reg" & quote
	WshShell.Run cmd, 0, True
end if
