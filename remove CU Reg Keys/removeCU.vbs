     
   
'==========================================================================================================================
'Standard Global Objects and variables
'==========================================================================================================================

On Error Resume Next
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_CLASSES_ROOT	= &H80000000

Dim strComputer, objReg, oShell, fso, objArgs
Dim var_hivareg, strKeypath, arrSubkeys, var_value, var_hivareg1, var_value1

var_hivareg = "HKEY_USERS"
var_key = "Software\Ordbogen.com"
var_key1 = "Printers\DevModePerUser"
var_key2 = "Software\DevModes2"
var_value = "LMab1err"
var_value1 = "LMADImon"
var_value2 = "Lexmark Universal Fax"
var_hivareg1 = "HKEY_CURRENT_USER"

'==========================================================================================================================

strComputer = "."
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv") 

DeleteRegKeyAndSubKeys var_hivareg1, var_key
msgbox var_hivareg1 & "\" & var_key
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
objRegistry.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys
Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
Set WShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set WNetwork = CreateObject("WScript.Network") 
strUserName = WNetwork.username 

For Each objSubkey In arrSubkeys
	strValueName = "ProfileImagePath"
	strSubPath = strKeyPath & "\" & objSubkey
	objRegistry.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath,strValueName,strValue
	correct_in strValue & "\ntuser.dat"
Next

If strUserName = "SYSTEM" Then
	string1 = "S-1-5-21"
	string2 = "Classes"
	objReg.EnumKey HKEY_USERS, "", arrsubkeys
	For Each Subkey In arrSubKeys
		If InStr(1,Subkey,string1,1) And (InStr(1,Subkey,string2,1) = "0") Then 
			clean Subkey
			msgbox "S-1-5-21"
		End If
	Next
End If

'==================================================================================================================
'Functions 
'==================================================================================================================

sub correct_in(dat_path)
		If FSO.FileExists(dat_path) Then
			WShell.Run "REG LOAD HKU\datastream\ " & """" & dat_path & """", 0, True
			clean "datastream"
			WShell.Run "REG UNLOAD HKU\datastream\", 0, True
		End If
End Sub

'==========================================================================================================================

Sub clean(id)
		DeleteRegKeyAndSubKeys var_hivareg, id & "\" & var_key
End Sub

'==========================================================================================================================

Function DeleteRegKeyAndSubKeys(strRegTree, strKeyPath)
    
	Dim arrSubKeys
	arrSubKeys = null
	
	Select Case strRegTree
		 Case "HKEY_CLASSES_ROOT"	  hTree = HKEY_CLASSES_ROOT
		 Case "HKEY_CURRENT_USER"	  hTree = HKEY_CURRENT_USER
		 Case "HKEY_LOCAL_MACHINE"	  hTree = HKEY_LOCAL_MACHINE
		 Case "HKEY_USERS"	          hTree = HKEY_USERS
		 Case "HKEY_CURRENT_CONFIG"	  hTree = HKEY_CURRENT_CONFIG
	End Select

	objRegistry.EnumKey hTree, strKeyPath, arrSubKeys 

	If IsArray(arrSubKeys) Then 
		For Each strSubKey In arrSubKeys 
			DeleteRegKeyAndSubKeys strRegTree, strKeyPath & "\" & strSubKey 
		Next 
	End If 

	objRegistry.DeleteKey hTree, strKeyPath 
msgbox "apel"
End Function

'==========================================================================================================================
