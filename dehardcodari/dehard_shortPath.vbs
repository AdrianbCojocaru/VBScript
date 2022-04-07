' --------------------------------------------------------------------------
'  File:         dehard.vbs
'  Purpose:  used for hardcoded files 
'  Date:          18,July,2017
'  Description:
'  Usage: example usage:
'  wscript.exe "dehard.vbs"
' ----------------------------------------------------------------------------
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim strFilePath, strToReplace, strLocFis, myShortPath, objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

strArgs = Session.Property("CustomActionData")
arrArgs = Split(strArgs, ";", -1, 1)

strLocFis 	= arrArgs(0) 'installdir
CP_SERVER	= arrArgs(1) 'AS410IEKI11
CP_DATABASE	= arrArgs(2) 'Astellas
CP_USERNAME	= arrArgs(3) 'Astellas
CP_PASSWORD	= arrArgs(4) '123

strToReplace = "C:\PROGRA~2\Infineer\ChipNet3"

Set myShortPath = objFSO.GetFolder(strLocFis)
strNewShortPath = myShortPath.ShortPath

Ascii strToReplace, strNewShortPath, strLocFis & "Card Production\CardProduction.cfg"
Ascii strToReplace, strNewShortPath, strLocFis & "Card Production\CardProduction_old.cfg"


Ascii "AS410IEKI11", CP_SERVER, strLocFis & "Card Production\CardProduction.cfg"
Ascii "AS410IEKI11", CP_SERVER, strLocFis & "Card Production\CardProduction_OLD.cfg"
Ascii "AS410IEKI11", CP_SERVER, strLocFis & "shell.cfg"

Ascii "123", CP_PASSWORD, strLocFis & "Card Production\CardProduction.cfg"
Ascii "123", CP_PASSWORD, strLocFis & "Card Production\CardProduction_OLD.cfg"
Ascii "123", CP_PASSWORD, strLocFis & "shell.cfg"

Ascii "CP_DATABASE=Astellas", "CP_DATABASE=" & CP_DATABASE, strLocFis & "Card Production\CardProduction.cfg"
Ascii "CP_DATABASE=Astellas", "CP_DATABASE=" & CP_DATABASE, strLocFis & "Card Production\CardProduction_OLD.cfg"
Ascii "Initial Catalog=Astellas", "Initial Catalog=" & CP_DATABASE, strLocFis & "shell.cfg"

Ascii "CP_USERNAME=Astellas", "CP_USERNAME=" & CP_USERNAME, strLocFis & "Card Production\CardProduction.cfg"
Ascii "CP_USERNAME=Astellas", "CP_USERNAME=" & CP_USERNAME, strLocFis & "Card Production\CardProduction_OLD.cfg"
Ascii "User ID=Astellas", "User ID=" & CP_USERNAME, strLocFis & "shell.cfg"
 

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