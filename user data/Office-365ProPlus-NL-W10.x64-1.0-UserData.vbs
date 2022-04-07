' --------------------------------------------------------------------------
'  File:         _SetCUregs.vbs.vbs
'  Purpose: 	 used to write user registries
'  Date:         02-Aug-2017
'  Description:  used to write user registries
'  Usage: example usage: wscript.exe "_SetCUregs.vbs"
' ----------------------------------------------------------------------------

RegWrite "32Bit", "HKEY_CURRENT_USER", "Software\Microsoft\Office\16.0\Common\General", 			"ShownFileFmtPrompt", 				"1", 								"REG_DWORD"
RegWrite "32Bit", "HKEY_CURRENT_USER", "Software\Microsoft\Office\16.0\Common\LanguageResources", 	"DisableProofingToolsAdvertisement","1",								"REG_DWORD"
RegWrite "32Bit", "HKEY_CURRENT_USER", "Software\Microsoft\Office\16.0\Common\WhatsNew",		 	"Microsoft Excel_ContentId", 		"PowerQueryManagSharedQueries",		"REG_SZ"
RegWrite "32Bit", "HKEY_CURRENT_USER", "Software\Microsoft\Office\16.0\Common\WhatsNew", 			"Microsoft Outlook_ContentId", 		"MoveToArchiveFolderBrowseGroups",	"REG_SZ"
RegWrite "32Bit", "HKEY_CURRENT_USER", "Software\Microsoft\Office\16.0\Outlook\Options\Calendar", 	"DisableWeather", 					"1",								"REG_DWORD"
RegWrite "32Bit", "HKEY_CURRENT_USER", "Software\Microsoft\Office\16.0\Outlook\Options\Mail", 		"JunkMailImportLists",				"1",								"REG_DWORD"
RegWrite "32Bit", "HKEY_CURRENT_USER", "Software\Microsoft\Office\16.0\OneNote\Options\Other", 		"NavigationBarExpColState", 		"0",								"REG_DWORD"
RegWrite "32Bit", "HKEY_CURRENT_USER", "Software\Microsoft\Office\16.0\OneNote\Options", 			"DisableCSCInfobar", 				"1",								"REG_DWORD"
RegWrite "32Bit", "HKEY_CURRENT_USER", "Software\Microsoft\Office\16.0\Registration", 				"AcceptAllEulas", 					"1",								"REG_DWORD"
RegWrite "32Bit", "HKEY_CURRENT_USER", "Software\Microsoft\Office\16.0\Outlook\Preferences", 		"DelegateSentItemsStyle", 			"1",								"REG_DWORD"
RegWrite "32Bit", "HKEY_CURRENT_USER", "Software\Microsoft\Office\16.0\Outlook\Autodiscover", 		"ExcludeHttpsRootDomain", 			"1",								"REG_DWORD"
RegWrite "32Bit", "HKEY_CURRENT_USER", "Software\Microsoft\Office\16.0\Outlook\Autodiscover", 		"ExcludeSCPLookup", 				"1",								"REG_DWORD"
RegWrite "32Bit", "HKEY_CURRENT_USER", "Software\Microsoft\Office\16.0\Common\WhatsNew", 			"SuppressForAutomation", 			"1",								"REG_DWORD"
RegWrite "32Bit", "HKEY_CURRENT_USER", "Software\Microsoft\Office\16.0\Common\General", 			"ShownFirstRunOptin", 				"1",								"REG_DWORD"
RegWrite "32Bit", "HKEY_CURRENT_USER", "Software\Microsoft\Office\16.0\Lync",			 			"IsBasicTutorialSeenByUser",		"1",								"REG_DWORD"

set oSH = CreateObject("WScript.Shell")
ComputerName = oSH.ExpandEnvironmentStrings("%COMPUTERNAME%")
RegWrite "32Bit", "HKEY_CURRENT_USER", "SOFTWARE\Microsoft\Office\16.0\Registration\EANDISW10\" & ComputerName & "{90160000-000F-0000-0000-0000000FF1CE}\O365ProPlusRetail\EULA", "16", "",	"REG_SZ"


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