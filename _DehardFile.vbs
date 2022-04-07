' --------------------------------------------------------------
'  Author: 			MS GmbH
'  Description:		Used to solve hardcoded paths.
' --------------------------------------------------------------

'On Error Resume Next

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim strFilePath, strToReplace, strNewValue, FileName, FolderName, myPattern, replacestring
Set objFSO2 = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

PropertyLine = Session.Property("CustomActionData")
ArrayPropertyLine = Split (PropertyLine, ";")

vDirSL = ArrayPropertyLine(0)
'vDirSL =  Session.Property("CustomActionData")
idir = left(vDirSL,len(vDirSL)-1)

SADD = ArrayPropertyLine(1)
SPORT = ArrayPropertyLine(2)
'replacestring1 = replace(replacestring, "\", "\\")

myPattern1 = "C:\Program Files (x86)\Amana Consulting\SmartNotesClient"
myPattern2 = "127.0.0.1"
myPattern3 = "8888"
'myPatternHome = "http://intranet/"

ReplaceInFile myPattern1, idir, idir & "\SmartNotes.exe.config"
ReplaceInFile myPattern2, SADD, idir & "\SmartNotes.exe.config"
ReplaceInFile myPattern3, SPORT, idir & "\SmartNotes.exe.config"



Function ReplaceInFile(strToReplace, strNewValue, strFilePath)

	Dim objFSO, objFile, strText
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	if objFSO.FileExists(strFilePath) Then
		Set objFile = objFSO.OpenTextFile(strFilePath, ForReading, True)
			strText = objFile.ReadAll
		objFile.Close
		Set objFile = Nothing
		
		strText = Replace(strText, strToReplace, strNewValue, 1, -1, 0)
	
		Set objFile = objFSO.CreateTextFile(strFilePath, True)
		objFile.Write strText
		objFile.Close
		Set objFile = Nothing
	
	End If

	Set objFSO = Nothing

End Function