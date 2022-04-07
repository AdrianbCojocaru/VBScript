Dim destination, objFSO, sh, file1, file2, file3, file4, file5, file6
Const OverwriteExisting = True

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("Wscript.shell")

currentDirectory = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))

file1 = CurrentDirectory & "SlpUserConfig.xml"
file2 = CurrentDirectory & "slpformatsmru0.dat"
file3 = CurrentDirectory & "slplabelsmru0.dat"
file4 = CurrentDirectory & "slplabelsmru1.dat"
file5 = CurrentDirectory & "slplabeltypesmru0.dat"
file6 = CurrentDirectory & "slplabeltypesmru1.dat"

destination = sh.expandenvironmentstrings("%AppData%")
destination = destination & "\Smart Label Printer\"

If Not objFSO.FolderExists(destination) Then 
	objFSO.CreateFolder destination
End If  
objFSO.CopyFile file1, destination, OverwriteExisting
objFSO.CopyFile file2, destination, OverwriteExisting
objFSO.CopyFile file3, destination, OverwriteExisting
objFSO.CopyFile file4, destination, OverwriteExisting
objFSO.CopyFile file5, destination, OverwriteExisting
objFSO.CopyFile file6, destination, OverwriteExisting
