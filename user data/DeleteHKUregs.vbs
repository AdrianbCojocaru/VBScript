' --------------------------------------------------------------------------
'  File:         DeleteHKUregs.vbs
'  Purpose: 	 used to write user registries
'  Date:         14-Aug-2017
'  Description:  used to write user registries
'  Usage: example usage: wscript.exe "DeleteHKUregs.vbs"
' ----------------------------------------------------------------------------

on error resume next
LOCALIZED_SENDTOMM_IMAGE = Session.Property("LOCALIZED_SENDTOMM_IMAGE")
LOCALIZED_SENDTOMM_LINK = Session.Property("LOCALIZED_SENDTOMM_LINK")
LOCALIZED_SENDTOMM_PAGE = Session.Property("LOCALIZED_SENDTOMM_PAGE")
LOCALIZED_SENDTOMM_TEXT = Session.Property("LOCALIZED_SENDTOMM_TEXT")

dim strComputer, fso
dim regkeycontents1, counter, temp, InitialString, sh, objSubkey
dim source, strKeyPath, strSubPath, objRegistry, arrSubkeys, strValueName, keyval, keyname
Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
Set fso = CreateObject( "Scripting.FileSystemObject")
set sh = createobject("wscript.shell")
Set objRegistry=GetObject( "winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
objRegistry.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys, keyval, keyname

For Each objSubkey In arrSubkeys
	strValueName = "ProfileImagePath"
	strSubPath = strKeyPath & "\" & objSubkey
	If len(objSubkey) > 8 then	
		RegKeyDelete "32Bit", "HKEY_USERS", objSubkey & "\Software\Microsoft\Internet Explorer\MenuExt\" & LOCALIZED_SENDTOMM_IMAGE
		RegKeyDelete "32Bit", "HKEY_USERS", objSubkey & "\Software\Microsoft\Internet Explorer\MenuExt\" & LOCALIZED_SENDTOMM_LINK
		RegKeyDelete "32Bit", "HKEY_USERS", objSubkey & "\Software\Microsoft\Internet Explorer\MenuExt\" & LOCALIZED_SENDTOMM_PAGE
		RegKeyDelete "32Bit", "HKEY_USERS", objSubkey & "\Software\Microsoft\Internet Explorer\MenuExt\" & LOCALIZED_SENDTOMM_TEXT			
	End if
Next


Function RegKeyDelete(strArhitecture, strRegRoot, strRegKeyPath)
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

	If keyExists(strArhitecture, hexRegRoot, strRegKeyPath) Then
		objReg.EnumKey hexRegRoot, strRegKeyPath, arrSubkeys
		If IsArray(arrSubkeys) Then
			For Each strSubkey In arrSubkeys
				statusCode = objReg.DeleteKey(hexRegRoot, strRegKeyPath & "\" & strSubkey)
			Next
		End If
		statusCode = objReg.DeleteKey(hexRegRoot, strRegKeyPath)
		RegKeyDelete = statusCode

	End If
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