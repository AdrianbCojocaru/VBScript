' --------------------------------------------------------------------------
'  File:         _DeleteAddinReg
'  Purpose:  used to remove the excel add-in
'  Date:          March 20, 2017
'  Description:
'  Usage: example usage:
'  wscript.exe "_DeleteAddinReg"
' ----------------------------------------------------------------------------

on error resume next
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
strComputer = "."

Set WShell = CreateObject( "WScript.Shell")
Set FSO = CreateObject( "Scripting.FileSystemObject")
Set WNetwork = CreateObject( "WScript.Network")

'----------------------------------------------------------------------------------------------------------------
strArgs = Session.Property("CustomActionData")
arrArgs = Split(strArgs, ";", -1, 1) 

excel_path = arrArgs(0) & "Excel.exe"
addin_path = arrArgs(1) & "rendite.xlam"

version = fso.GetFileVersion(excel_path)
version = mid(version, 1, instr(instr(1, version, ".") + 1, version, ".") - 1)


Set objRegistry=GetObject( "winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
objRegistry.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys

Set objReg = GetObject( "winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

'----------------------------------------------------------------------------------------------------------------

For Each objSubkey In arrSubkeys
	strValueName = "ProfileImagePath"
	strSubPath = strKeyPath & "\" & objSubkey
	objRegistry.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath,strValueName,strValue
	correct_in strValue
Next


strUserName = WNetwork.username 

If UCase(strUserName) = getUserNameFromSID("S-1-5-18") Then
   string1 = "S-1-5-21"
   string2 = "Classes"
   objReg.EnumKey HKEY_USERS, "", arrsubkeys
   For Each Subkey In arrSubKeys
      If InStr(1,Subkey,string1,1)And (InStr(1,Subkey,string2,1) = "0") Then
       ID = Subkey
       Call SystemRun(ID)
      End If
   Next
  
Else
     Call UserSettings
End If

'----------------------------------------------------------------------------------------------------------------

Sub SystemRun(ID)
	ADDINS_KEY = "HKEY_USERS\" & ID & "\Software\Microsoft\Office\" & version & "\Excel\Options\OPEN"
	clean addin_path, ADDINS_KEY 
	ADDINS_KEY1 = "HKEY_USERS\" & ID & "\Software\Microsoft\Office\" & version & "\Excel\Options\MsoTbCust"
	clean addin_path, ADDINS_KEY1
	

End Sub

'----------------------------------------------------------------------------------------------------------------

Sub UserSettings
	ADDINS_KEY = "HKCU\Software\Microsoft\Office\" & version & "\Excel\Options\OPEN"
	clean addin_path, ADDINS_KEY 
	ADDINS_KEY1 = "HKCU\Software\Microsoft\Office\" & version & "\Excel\Options\MsoTbCust"
	clean addin_path, ADDINS_KEY1
End Sub


'----------------------------------------------------------------------------------------------------------------

Sub correct_in(dat_path)
	If FSO.FileExists(dat_path & "\ntuser.dat") Then
		WShell.Run "cmd /c REG LOAD HKLM\Addins_Excel\ """ & dat_path & "\NTUSER.DAT""", 2, True
		ADDINS_KEY = "HKLM\Addins_Excel\Software\Microsoft\Office\" & version & "\Excel\Options\OPEN"
		clean addin_path, ADDINS_KEY
		WShell.Run "cmd /c REG UNLOAD HKLM\Addins_Excel\", 2, True
	End If
End Sub

'----------------------------------------------------------------------------------------------------------------
Sub clean(add_in_name, KEY)
	Dim WshShell
	Dim TableExcelAdd
	Dim NumberExcelAddins
	On Error Resume Next
		Set WshShell = CreateObject("WScript.Shell")
		EXCEL_ADDINS_PATH=KEY
		NumberExcelAddins = 1
		ReDim TableExcelAdd(NumberExcelAddins)

		TableExcelAdd(1) = """" & add_in_name & """"


	Dim ii
	Dim t
	Dim t2
	Dim LatestNumber

	ii = 100 'assuming maximum 100 addins
	Dim RetVal
	LatestNumber = 0
	Dim KeyName


	t = ii

	On Error Resume Next

For t = ii To 0 Step -1
    If (t = 0) Then
        KeyName = EXCEL_ADDINS_PATH
    Else
        KeyName = EXCEL_ADDINS_PATH + Trim(CStr(t))
    End If
    RetVal = ""
    RetVal = WshShell.RegRead(KeyName)

    If (RetVal <> "") Then
         For j = 1 To NumberExcelAddins
            If (InStr(1, lcase(TableExcelAdd(j)), lcase(RetVal))) Then
                WshShell.RegDelete KeyName
            End If
        Next
    End If
Next

'REORDERING remaining Excel AddIns
Dim TableRegPaths()

t2 = 0
For t = ii To 0 Step -1
    If (t = 0) Then
        KeyName = EXCEL_ADDINS_PATH
    Else
        KeyName = EXCEL_ADDINS_PATH + Trim(CStr(t))
    End If
    RetVal = ""
    RetVal = WshShell.RegRead(KeyName)

    If (RetVal <> "") Then
    t2 = t2 + 1
        ReDim Preserve TableRegPaths(t2)
    TableRegPaths(t2) = RetVal
    WshShell.RegDelete KeyName
    End If

Next

For t = t2 To 1 Step -1
    If (t = t2) Then
        WshShell.RegWrite EXCEL_ADDINS_PATH, TableRegPaths(t)
    Else
        WshShell.RegWrite EXCEL_ADDINS_PATH + Trim(CStr(t)),TableRegPaths(t)
    End If
Next
End Sub
'----------------------------------------------------------------------------------------------------------------
Function getUserNameFromSID(SID)
	Err.Clear
	server = "." 
	Set objWMIService = GetObject("winmgmts:\\" & server & "\root\cimv2") 
	Set objAccount = objWMIService.Get("Win32_SID.SID='" & SID & "'") 
	
	strUser = objAccount.AccountName 
	If Err.Number <> 0 Then 
		getUserNameFromSID = Err.Description 
		Err.Clear 
	Else 
		getUserNameFromSID = UCase(strUser) 
	End If 
End Function 