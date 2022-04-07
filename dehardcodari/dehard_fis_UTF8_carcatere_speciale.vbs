' --------------------------------------------------------------------------
'  File:         dehard.vbs
'  Purpose:  used for hardcoded files 
'  Date:          22,August,2013
'  Description:
'  Usage: example usage:
'  wscript.exe "dehard.vbs"
' ----------------------------------------------------------------------------
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim strFilePath, strToReplace, strNewValue, strArgs, arrArgs
Dim CONPORT, idb
'strArgs = Session.Property("CustomActionData")
'arrArgs = Split(strArgs, ";", -1, 1)

strLocFis = "E:\Aplicatii\Videotron\38447 - HedEx\Huawei\hedex\" 'installdir '"E:\Aplicatii\Videotron\38447 - HedEx\Huawei\hedex"
strNewValue = Left(strLocFis, (Len(strLocFis) - 1))
idb =  Replace(strNewValue, "\", "/")					'"E:/Aplicatii/Videotron/38447 - HedEx/Huawei\hedex"
CONPORT = "52199"



Function Ascii(strToReplace, strNewValue, strFilePath)

	Dim objFSO, objFile, strText, re
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	  dim objStream
  ' ADODB stream object used
	set objStream = WScript.CreateObject("ADODB.Stream")
	if objFSO.FileExists(strFilePath) Then

'.		Set objFile = objFSO.OpenTextFile(strFilePath, ForReading, True)
'			strText = objFile.ReadAll
'		objFile.Close
'		Set objFile = Nothing
		  objStream.Open
		  objStream.CharSet = "utf-8"
		  objStream.LoadFromFile(strFilePath)
		  strText = objStream.ReadText()
			strText = Replace(strText, strToReplace, strNewValue, 1, -1, 0)
		set objStream = Nothing
'		Set objFile = objFSO.CreateTextFile(strFilePath, True)
'		objFile.Write strText
'		objFile.Close
'		Set objFile = Nothing
		set objStream = WScript.CreateObject("ADODB.Stream")
		objStream.CharSet = "utf-8"
		objStream.Open
		objStream.WriteText strText
		objStream.SaveToFile strFilePath, 2
		set objStream = Nothing
		msgbox "end if"
	End If

	Set objFSO = Nothing

End Function

strToReplace1 = "C:\Program Files\Huawei\hedex"
strToReplace2 = "C:/Program Files/Huawei/hedex"

Ascii strToReplace1, strNewValue, strLocFis & "conf\setup.xml"			' "C:\Program Files(x86)\Huawei\hedex"
Ascii strToReplace2, idb, strLocFis & "conf\setup.xml"					' "C:/Program Files(x86)/Huawei/hedex"
strToReplace = "52199"
Ascii strToReplace, CONPORT, strLocFis & "\server\conf\server.xml"
Ascii strToReplace, CONPORT, strLocFis & "\conf\setup.xml"
