' --------------------------------------------------------------------------

'  This script is used to copy applciation information to each user profile at log-on
'  Usage: wscript.exe UserData.vbs"

' ----------------------------------------------------------------------------
Set objShell = CreateObject("Wscript.shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

ScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
ScriptPath = Left(ScriptPath, len(ScriptPath)-1)
LocalLow = RegRead("32Bit", "HKEY_CURRENT_USER", "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{A520A1A4-1780-4FF6-BD18-167343C5AF16}", "REG_SZ")
CopyFile ScriptPath & "\deployment.properties", LocalLow & "\Sun\Java\Deployment\"

RegWrite "32Bit", "HKEY_CURRENT_USER", "Software\AppDataLow\Software\JavaSoft\DeploymentProperties", "deployment.expiration.decision.suppression.11.66.2", "true", "REG_SZ"
RegWrite "32Bit", "HKEY_CURRENT_USER", "Software\AppDataLow\Software\JavaSoft\DeploymentProperties", "deployment.expiration.decision.11.66.2", "later", "REG_SZ"

sub CopyFile (Source, Destination)	
	If not objFSO.FolderExists(Destination) Then
		mkdirCommand = "cmd.exe /c mkdir " &  """" & Destination & """"
		objShell.Run mkdirCommand, 0, true
	end if
		shellCommand = "cmd.exe /c copy " &  """" & Source & """" & " " & """" & Destination & """" & " /Y"
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
	
	RegRead = strValue
	else
	
	RegRead = "Not Found"
	end if
		
End Function

Function RegWrite(strArhitecture, strRegRoot, strRegKeyPath, strRegValName, strValue, strType)
		
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
	 

	Select Case strRegRoot
	   Case "HKEY_CLASSES_ROOT" hexRegRoot = &H80000000  
	   Case "HKEY_CURRENT_USER" hexRegRoot = &H80000001  
	   Case "HKEY_LOCAL_MACHINE"  hexRegRoot = &H80000002 
	   Case "HKEY_USERS" hexRegRoot = &H80000003  
	   Case "HKEY_CURRENT_CONFIG" hexRegRoot = &H80000005  
	   Case Else hexRegRoot = "not set"
	End Select

	' First create the key if it does not already exist
		 If Not keyExists(strArhitecture, hexRegRoot, strRegKeyPath) Then
			 objReg.CreateKey hexRegRoot, strRegKeyPath
		  End If

		 Select Case strType
				Case "REG_BINARY"
			iValues = Array(strValue)
				   statusCode = objReg.SetBinaryValue(hexRegRoot, strRegKeyPath, strRegValName, iValues) 
				Case "REG_SZ" 
				   statusCode = objReg.SetStringValue(hexRegRoot, strRegKeyPath, strRegValName, strValue)
				Case "REG_EXPAND_SZ"
				   statusCode = objReg.SetExpandedStringValue(hexRegRoot, strRegKeyPath, strRegValName, strValue)
				Case "REG_MULTI_SZ"
				   statusCode = objReg.SetMultiStringValue(hexRegRoot, strRegKeyPath, strRegValName, strValue)
				Case "REG_DWORD"
				   statusCode = objReg.SetDWORDValue(hexRegRoot, strRegKeyPath, strRegValName, strValue)
				Case "REG_QWORD"
				   statusCode = objReg.SetQWORDValue(hexRegRoot, strRegKeyPath, strRegValName, strValue)
			 End Select 
	RegWrite = statusCode	
End Function

Function keyExists(strArhitecture, hexRegRoot, strRegKeyPath)
      ' Determine if a registry key exists
      '     First we need to determine if the key already exists, and if not, create it

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

      blnSubKeyExists = False
      strRegKeyPath1 = Left(strRegKeyPath, InStrRev(strRegKeyPath, "\") - 1)  ' Get the portion of the path before the last "\"
      strSubKeyToMatch = Right(strRegKeyPath, Len(strRegKeyPath) - InStrRev(strRegKeyPath, "\")) ' Get the subkey to find

      objReg.EnumKey hexRegRoot, strRegKeyPath1, arrSubKeys  ' Enumerate the subkeys to see if this subkey exists
      If IsArray(arrSubKeys) Then

         For Each strSubKey In arrSubKeys
            If strSubKey = strSubKeytoMatch Then 
            	  blnSubKeyExists = True

            End If
         Next 'strSubKey
      End If
      keyExists = blnSubKeyExists
End Function