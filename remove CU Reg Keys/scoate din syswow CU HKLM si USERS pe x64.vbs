'------------------------------------------------------------------------------
' Date			:	31-10-2013
'------------------------------------------------------------------------------



	SETUPGUIHEADER = "TM1 Perspectives" 					  'window headed
	SETUPGUIINFO1 = "Please be patient..."						  'Firs line of mesage
	SETUPGUIINFO2 = " "								  'Second line of mesage
	SETUPGUIINFO3 = " "								  'third line of mesage	
	SETUPGUIINFO4 = "TM1 Perspectives is uninstalling now..."		  'forth line of mesage	
	SETUPGUIINFO5 = ""	  	  		  'forth line of mesage
 	SETUPGUITIMEOUT = 1800      			 				  'window timeout in seconds 


'
'------------------------------------------------------------------------------

' Variable Declarations -------------------------------------------------------
'
	Dim WshShell
	Dim oInstaller
	Dim fso
'
'------------------------------------------------------------------------------

' Core variable definitions ---------------------------------------------------
'
	Set WshShell   = CreateObject("WScript.Shell")
	Set oInstaller = CreateObject("WindowsInstaller.Installer")
	Set fso        = CreateObject("Scripting.FileSystemObject")

	
'
'------------------------------------------------------------------------------

' Run the Uninstall of IBMCogno-TM1Perspectives-9.5.2-EN-1.0-------------------------------------------------------------
'

	sInstallCommand = "msiexec.exe /x {833BA879-3A8F-4624-A209-770C15AAC541} /qn /l*v " & """" & WshShell.expandenvironmentstrings("%windir%") & "\Logs\IBMCogno-TM1Perspectives-9.5.2-EN-1.0_hotfix_UnInstall.log"""

	returnCode = WshShell.Run(sInstallCommand, 1, TRUE)

	if returnCode <> 0 AND returnCode <> 3010 AND returnCode <> 1605 then
		WScript.Quit returnCode
	end if
	
	'-----------------------------------------------------------------------------------------------	
	Const HKEY_CURRENT_USER = &H80000001
	Const HKEY_LOCAL_MACHINE = &H80000002
	Const HKEY_USERS = &H80000003
	Const HKEY_CURRENT_CONFIG = &H80000005
	Const HKEY_CLASSES_ROOT	= &H80000000
	strComputer = "."
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
			 strComputer & "\root\default:StdRegProv")
		
	oReg.GetStringValue HKEY_LOCAL_MACHINE, "Software\Applix\Install\9.0", "InstallDir", INSTALLDIR
	
'-----------------------------------------------------------------------------------------------	
	Dim WMI: Set WMI = GetObject("winmgmts:{impersonationlevel=impersonate}")
	DIM Coll: Set Coll = WMI.ExecQuery("select Name,VariableValue from Win32_Environment where Name='PROCESSOR_ARCHITECTURE'")
	DIM EnvVar, ProcessorArchitecture
	For Each EnvVar In Coll
      ProcessorArchitecture = EnvVar.VariableValue
	next

	If InStr(ProcessorArchitecture,"64") <> 0 Then
		
		if (fso.fileExists(INSTALLDIR & "bin\PDFCamp\unpdfx64.exe")) then	
			sInstallCommand = """" & INSTALLDIR & "bin\PDFCamp\unpdfx64.exe"""
			returnCode = WshShell.Run(sInstallCommand, 1, TRUE)

			if returnCode <> 0 AND returnCode <> 3010 then
				WScript.Quit returnCode
			end if
		end if
		
		
		WshShell.Run "cmd /c rmdir /q /s " & Chr(34) & INSTALLDIR & "bin\PDFCamp" & Chr(34),2,true
	end if
	
'-----------------------------------------------------------------------------------------------	
	sRegisterCommand = "regsvr32 -u -s """ & INSTALLDIR & "bin\tm1prc.dll"" " & """" & INSTALLDIR & "bin\tm1xl.ocx"" " & """" & INSTALLDIR & "bin\tm1dasrv.dll"""
	returnCode = WshShell.Run(sRegisterCommand, 1, TRUE)
'-----------------------------------------------------------------------------------------------	
	
	sInstallCommand = "msiexec.exe /x {AD063608-666F-4B6F-B66E-204661EE9CB2} /qn /l*v " & """" & WshShell.expandenvironmentstrings("%windir%") & "\Logs\IBMCogno-TM1Perspectives-9.5.2-EN-1.0_UnInstall.log"""

	returnCode = WshShell.Run(sInstallCommand, 1, TRUE)

	if returnCode <> 0 AND returnCode <> 3010 AND returnCode <> 1605 then
		WScript.Quit returnCode
	end if


'------------------------------------------------------------------------------
' Post Install Tasks ----------------------------------------------------------
'

'delete project registries
	oReg.DeleteKey HKEY_LOCAL_MACHINE, "Software\OMV\IBM Cognos-TM1 Perspectives-9.5.2 FP1"
	INSTALLDIR1 = INSTALLDIR
	WshShell.Run "cmd /c rmdir /q /s " & Chr(34) & INSTALLDIR1 & Chr(34),2,true
	
	if (right(INSTALLDIR1,1) = "\") then
		INSTALLDIR1 = left(INSTALLDIR1, len(INSTALLDIR1)-1)
	end if
	
	TestArray = Split(INSTALLDIR1 , "\")
	For i = LBound(TestArray) to UBound(TestArray) -1
		PARENT = PARENT & TestArray(i) &  "\"
	Next
	
	If fso.FolderExists(PARENT) Then
		Set fld = fso.GetFolder(PARENT)
		If fld.Files.Count + fld.SubFolders.Count = 0 Then
			WshShell.Run "cmd /c rmdir /q /s " & """" & PARENT & """",2,true
		End If
	End If

	Programs = WshShell.SpecialFolders("AllUsersPrograms")
	WshShell.Run "cmd /c rmdir /q /s " & Chr(34) & Programs & "\IBM Cognos\TM1" & Chr(34),2,true
	wscript.sleep 5000
	PARENT = Programs & "\IBM Cognos"
	If fso.FolderExists(PARENT) Then
		Set fld = fso.GetFolder(PARENT)
		If fld.Files.Count + fld.SubFolders.Count = 0 Then
			WshShell.Run "cmd /c rmdir /q /s " & """" & PARENT & """",2,true
		End If
	End If
	
'------------------------------------------------------------------------------	
'remove active setup registries

strProductCode = "IBMCogno-TM1Perspectives-9.5.2-EN-1.0"
strKeyPath2 = "SOFTWARE\Microsoft\Active Setup\Installed Components"
oReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath2, arrSubKeys2

For Each Subkey2 in arrSubKeys2
	If Instr(1, Subkey2, strProductCode) > 0 Then
		oReg.DeleteKey HKEY_LOCAL_MACHINE, strKeyPath2 & "\" & Subkey2
	End If
Next


'--------------------------------------------------------------------------------
' remove excel addin

Set WNetwork = CreateObject( "WScript.Network")

oReg.GetStringValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe", "Path", excel_path
excel_path = excel_path & "excel.exe"
addin_path = INSTALLDIR & "bin\tm1p.xla"

version = fso.GetFileVersion(excel_path)
version = mid(version, 1, instr(instr(1, version, ".") + 1, version, ".") - 1)



Set oReg=GetObject( "winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
oReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys

Set oReg = GetObject( "winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

'----------------------------------------------------------------------------------------------------------------

For Each objSubkey In arrSubkeys
	strValueName = "ProfileImagePath"
	strSubPath = strKeyPath & "\" & objSubkey
	oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath,strValueName,strValue
	correct_in strValue
Next


strUserName = WNetwork.username 

If strUserName = "SYSTEM" Then
   string1 = "S-1-5-21"
   string2 = "Classes"
   oReg.EnumKey HKEY_USERS, "", arrsubkeys
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
		WshShell.Run "cmd /c REG LOAD HKLM\Addins_Excel\ """ & dat_path & "\NTUSER.DAT""", 2, True
		ADDINS_KEY = "HKLM\Addins_Excel\Software\Microsoft\Office\" & version & "\Excel\Options\OPEN"
		clean addin_path, ADDINS_KEY
		WshShell.Run "cmd /c REG UNLOAD HKLM\Addins_Excel\", 2, True
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
'delete user registries
On Error Resume Next
var_hivareg = "HKEY_USERS"
var_hivareg1 = "HKEY_CURRENT_USER"
var_key = "Software\Microsoft\Office\" & version & "\Excel\Add-in Manager"
var_value = addin_path

DeleteRegValueIfExists var_hivareg1, var_key, var_value

strProductCode = "IBMCogno-TM1Perspectives-9.5.2-EN-1.0"
strKeyPath1 = "SOFTWARE\Microsoft\Active Setup\Installed Components"

oReg.EnumKey HKEY_CURRENT_USER, strKeyPath1, arrSubKeys1

For Each Subkey1 in arrSubKeys1
	If Instr(1, Subkey1, strProductCode) > 0 Then
	DeleteRegKeyAndSubKeys var_hivareg1, strKeyPath1 & "\" & Subkey1
	End If
Next

If InStr(ProcessorArchitecture,"64") <> 0 Then
	strKeyPath11 = "SOFTWARE\Wow6432Node\Microsoft\Active Setup\Installed Components"
	oReg.EnumKey HKEY_CURRENT_USER, strKeyPath11, arrSubKeys11
	For Each Subkey11 in arrSubKeys11
		If Instr(1, Subkey11, strProductCode) > 0 Then
			DeleteRegKeyAndSubKeys var_hivareg1, strKeyPath11 & "\" & Subkey11
		End If
	Next
End If

Set oReg=GetObject( "winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
oReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys
strUserName = WNetwork.username 
Set oReg = GetObject( "winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

For Each objSubkey In arrSubkeys
	strValueName = "ProfileImagePath"
	strSubPath = strKeyPath & "\" & objSubkey
	oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath,strValueName,strValue
	correct_in1 strValue & "\ntuser.dat"
Next

If strUserName = "SYSTEM" Then
	string1 = "S-1-5-21"
	string2 = "Classes"
	oReg.EnumKey HKEY_USERS, "", arrsubkeys
	For Each Subkey In arrSubKeys
		If InStr(1,Subkey,string1,1) And (InStr(1,Subkey,string2,1) = "0") Then 
			clean1 Subkey
		End If
	Next
End If

'==================================================================================================================
'Functions 
'==================================================================================================================

sub correct_in1(dat_path)
		If FSO.FileExists(dat_path) Then
			WshShell.Run "REG LOAD HKU\datastream\ " & """" & dat_path & """", 0, True
			clean1 "datastream"
			WshShell.Run "REG UNLOAD HKU\datastream\", 0, True
		End If
End Sub

'==========================================================================================================================

Sub clean1(id)
	DeleteRegValueIfExists var_hivareg, id & "\" & var_key & "\", var_value
	
	strProductCode = "IBMCogno-TM1Perspectives-9.5.2-EN-1.0"
	strKeyPath1 = "SOFTWARE\Microsoft\Active Setup\Installed Components"

	oReg.EnumKey HKEY_USERS, id & "\" & strKeyPath1, arrSubKeys1
	
	For Each Subkey1 in arrSubKeys1
		If Instr(1, Subkey1, strProductCode) > 0 Then
			DeleteRegKeyAndSubKeys var_hivareg, id & "\" & strKeyPath1 & "\" & Subkey1
		End If
	Next
	If InStr(ProcessorArchitecture,"64") <> 0 Then
		strKeyPath11 = "SOFTWARE\Wow6432Node\Microsoft\Active Setup\Installed Components"
		oReg.EnumKey HKEY_USERS, id & "\" & strKeyPath11, arrSubKeys11
		For Each Subkey11 in arrSubKeys11
			If Instr(1, Subkey11, strProductCode) > 0 Then
				DeleteRegKeyAndSubKeys var_hivareg, id & "\" & strKeyPath11 & "\" & Subkey11
			End If
		Next
	End If
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

	oReg.EnumKey hTree, strKeyPath, arrSubKeys 

	If IsArray(arrSubKeys) Then 
		For Each strSubKey In arrSubKeys 
			DeleteRegKeyAndSubKeys strRegTree, strKeyPath & "\" & strSubKey 
		Next 
	End If 

	oReg.DeleteKey hTree, strKeyPath 

End Function

'==========================================================================================================================

Function DeleteRegValueIfExists(strRegHive, strKeyPath, strRegValue)
    
	Select Case strRegHive
		Case "HKEY_CLASSES_ROOT"	  rHive = HKEY_CLASSES_ROOT
		Case "HKEY_CURRENT_USER"	  rHive = HKEY_CURRENT_USER
		Case "HKEY_LOCAL_MACHINE"	  rHive = HKEY_LOCAL_MACHINE
		Case "HKEY_USERS"	          rHive = HKEY_USERS
		Case "HKEY_CURRENT_CONFIG"	  rHive = HKEY_CURRENT_CONFIG
	End Select

	oReg.DeleteValue rHive, strKeyPath, strRegValue
	
End Function

'==========================================================================================================================


'------------------------------------------------------------------------------
	WScript.Quit returnCode
'
'------------------------------------------------------------------------------


Sub StartProgressWindow
 Set objFSO = CreateObject("Scripting.FileSystemObject")
 sTemp = objFSO.GetSpecialFolder(2)
 Set TS = objFSO.CreateTextFile(sTemp & "\" & "Popup.cmd", True)

	Ts.WriteLine "@SET SETUPGUIHEADER=" & SETUPGUIHEADER 
	Ts.WriteLine "@SET SETUPGUIINFO1=" & SETUPGUIINFO1 
	Ts.WriteLine "@SET SETUPGUIINFO2=" & SETUPGUIINFO2 
	Ts.WriteLine "@SET SETUPGUIINFO3=" & SETUPGUIINFO3 
	Ts.WriteLine "@SET SETUPGUIINFO4=" & SETUPGUIINFO4 
	Ts.WriteLine "@SET SETUPGUIINFO5=" & SETUPGUIINFO5 
	Ts.WriteLine "@SET SETUPGUITIMEOUT=" & SETUPGUITIMEOUT 
	Ts.WriteLine "@REM Setup-Fenster oeffnen" 
	Ts.WriteLine "@start %systemroot%\Tools\SetupGUI" 
	Ts.WriteLine "@rem Software installieren"
	TS.Close
	
	WshShell.Run sTemp & "\" & "Popup.cmd", 1, False
	

End Sub


Sub HideProgressWindow
 Dim strComputer
 Dim objWMIService
 Dim colProcessList
 Dim objProcess
 strComputer = "." 
 Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
 Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'SetupGUI.exe'") 
 Wscript.Sleep 1000
 For Each objProcess In colProcessList
       	objProcess.Terminate()
 Next
 Set objWMIService = Nothing
 Set colProcessList = Nothing

End Sub
