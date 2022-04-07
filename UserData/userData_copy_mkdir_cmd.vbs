' --------------------------------------------------------------------------

'  Usage: wscript.exe UserData.vbs"

' ----------------------------------------------------------------------------
Set objShell = CreateObject("Wscript.shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")


PersonalFolder = objShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Personal")

CopyFolder session.property("IBM1"), PersonalFolder & "\IBM\"

sub CopyFolder (Source, Destination)
	Set objShell = CreateObject("WScript.Shell")
	If Right(Source, 1) ="\" then Source = Left(Source, len(Source)-1)
	shellCommand = "cmd.exe /c xcopy " &  """" & Source & """" & " " & """" & Destination & """" & " /E /F /R /H /I /Y"
	objShell.Run shellCommand, 0, true
End sub

Function RegRead(strArhitecture, strRegRoot, strRegKeyPath, strRegValName, strType)
	
    strComputer = "."
    Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet") 
strArhitecture = lcase (strArhitecture)
Select Case strArhitecture
   Case "32bit" 
    objCtx.Add "__ProviderArchitecture", 32	
   Case "64bit" 
    objCtx.Add "__ProviderArchitecture", 64 
End Select
    Set objLocator = CreateObject("Wbemscripting.SWbemLocator") 
    Set objServices = objLocator.ConnectServer("","root\default","","",,,,objCtx) 
    Set objReg = objServices.Get("StdRegProv")
 
statusCode = objReg.GetStringValue(&H80000001, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal", strValue)

Select Case strRegRoot
   Case "HKEY_CLASSES_ROOT" hexRegRoot = &H80000000  
   Case "HKEY_CURRENT_USER" hexRegRoot = &H80000001  
   Case "HKEY_LOCAL_MACHINE"  hexRegRoot = &H80000002 
   Case "HKEY_USERS" hexRegRoot = &H80000003  
   Case "HKEY_CURRENT_CONFIG" hexRegRoot = &H80000005  
   Case Else hexRegRoot = "not set"
End Select


	 Select Case strType
            Case "REG_BINARY"
               statusCode = objReg.GetBinaryValue(hexRegRoot, strRegKeyPath, strRegValName, strValue) 
            Case "REG_SZ" 
               statusCode = objReg.GetStringValue(hexRegRoot, strRegKeyPath, strRegValName, strValue)
            Case "REG_EXPAND_SZ"
               statusCode = objReg.GetExpandedStringValue(hexRegRoot, strRegKeyPath, strRegValName, strValue)
            Case "REG_MULTI_SZ"
               statusCode = objReg.GetMultiStringValue(hexRegRoot, strRegKeyPath, strRegValName, strValue)
            Case "REG_DWORD"
               statusCode = objReg.GetDWORDValue(hexRegRoot, strRegKeyPath, strRegValName, strValue)
            Case "REG_QWORD"
               statusCode = objReg.GetQWORDValue(hexRegRoot, strRegKeyPath, strRegValName, strValue)
         End Select 


If  statusCode=0 then
WriteToLog strLogFile, "Read value for " & strArhitecture & " reg : " & strRegRoot & "\" & strRegKeyPath & " | "& strRegValName & " = " & strValue
RegRead = strValue
else
WriteToLog strLogFile, "Read value for " & strArhitecture & " reg : " & strRegRoot & "\" & strRegKeyPath & " | "& strRegValName & " failed. Reg Not Found"
RegRead = "Not Found"
end if
	
End Function
