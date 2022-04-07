' --------------------------------------------------------------------------
'  File:         dehard.vbs
'  Purpose:  used for hardcoded files 
'  Date:          11,September,2013
'  Description:
'  Usage: example usage:
'  wscript.exe "dehard.vbs"
' ----------------------------------------------------------------------------
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim strFilePath, strToReplace, strNewValue, strArgs, arrArgs
Dim CONPORT, idb
strArgs = Session.Property("CustomActionData")
arrArgs = Split(strArgs, ";", -1, 1)
strLocFis = arrArgs(0) 'installdir
CONPORT = arrArgs(1) 	'"52199"

strNewValue = Left(strLocFis, (Len(strLocFis) - 1))
idb =  Replace(strNewValue, "\", "/")


Function Ascii(strToReplace, strNewValue, strFilePath)

	Dim objFSO, objStream, strText, re
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	if objFSO.FileExists(strFilePath) Then
		Set objStream = CreateObject("ADODB.Stream")
		objStream.Open
		objStream.CharSet = "utf-8"
		objStream.LoadFromFile(strFilePath)
		strText = objStream.ReadText()
		strText = Replace(strText, strToReplace, strNewValue, 1, -1, 0)
		Set objStream = Nothing
		
		Set objStream = CreateObject("ADODB.Stream")
		objStream.CharSet = "utf-8"
		objStream.Open
		objStream.WriteText strText
		objStream.SaveToFile strFilePath, 2
		Set objStream = Nothing
	End If

	Set objFSO = Nothing

End Function

strToReplace1 = "C:\Program Files\Huawei\hedex"
strToReplace2 = "C:/Program Files/Huawei/hedex"

Ascii strToReplace1, strNewValue, strLocFis & "conf\setup.xml"			' "C:\Program Files(x86)\Huawei\hedex"
Ascii strToReplace2, idb, strLocFis & "conf\setup.xml"					' "C:/Program Files(x86)/Huawei/hedex"
strToReplace = "52199"
Ascii strToReplace, CONPORT, strLocFis & "server\conf\server.xml"
Ascii strToReplace, CONPORT, strLocFis & "conf\setup.xml"
