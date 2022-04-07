' -------------------------------------------------------------------------- 
' File: Xemac5.8.vbs 
' Purpose: used to dehardcode the application's files 
' Date: 5,December,2016 
' Description: 
' Usage: example usage: 
' wscript.exe "Xemac5.8.vbs"
' -------------------------------------------------------------------------- 


Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim strFilePath, strToReplace, strNewValue, FileName, myPattern, replacestring, ScriptPath, ExeName, returnVal
Set objFSO2 = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

ScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")

ExeName = """" & ScriptPath & "Xemac5.8.exe" & """"
FileName = ScriptPath & "Xemac5.8.lax"
replacestring = WshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\"
replacestring = Replace(replacestring,"\", "\\", 1, -1, 0)
replacestring = "lax.command.line.args=./ " & """" & replacestring & "xemac-write" & """" & " xemac.cfg cc.cfg"

myPattern = "lax.command.line.args="

ReplaceInFile myPattern, replacestring, FileName
returnVal = WshShell.Run (ExeName, 1, false)


Function ReplaceInFile(strToReplace, strNewValue, strFilePath)

	Dim objFSO, objFile, strText, PropertyLine
	strText=""
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	if objFSO.FileExists(strFilePath) Then
		Set objFile = objFSO.OpenTextFile(strFilePath, ForReading, True)
			do until objFile.AtEndOfStream
				currline=objFile.ReadLine
				If InStr(1, currline, strToReplace, 1) then
					currline = strNewValue
				End if
				strText = strText & currline & vbCRLF
			loop
		objFile.Close
		Set objFile = Nothing

		Set objFile = objFSO.CreateTextFile(strFilePath, True)
		objFile.Write strText
		objFile.Close
		Set objFile = Nothing

	End If

	Set objFSO = Nothing

End Function 