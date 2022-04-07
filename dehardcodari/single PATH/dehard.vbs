' --------------------------------------------------------------------------
'  File:         dehard.vbs
'  Purpose:  used for hardcoded files 
'  Date:          25,September,2013
'  Description:
'  Usage: example usage:
'  wscript.exe "dehard.vbs"
' ----------------------------------------------------------------------------
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim strFilePath, strToReplace, strNewValue, strLocFis
strLocFis = Session.Property("CustomActionData")


Function Ascii(strToReplace, strNewValue, strFilePath)

	Dim objFSO, objFile, strText, re
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


strToReplace = "C:\Program Files\IBM\Client Access"
strNewValue = Left(strLocFis, (Len(strLocFis) - 1))
Ascii strToReplace, strNewValue, strLocFis & "Emulator\Private\AS400PRDA.ws"
Ascii strToReplace, strNewValue, strLocFis & "Emulator\Private\AS400PRDB.ws"
