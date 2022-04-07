'====================================================================================
' Uninstalls                      : Thermoworks_ThermaDataLogger_SeriesII_Gen_P0
'====================================================================================

On Error Resume Next
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_CLASSES_ROOT	= &H80000000

Dim strAppUninstallPath
Dim strComputer, objReg, oShell, oFSO
Dim var_hivareg, strKeypath, arrSubkeys, var_hivareg1
Dim subkey, var_key2, strSFN

var_hivareg = "HKEY_USERS"
var_key2 = "Software\Microsoft\Active Setup\Installed Components\Thermoworks_ThermaDataLogger_SeriesII_Gen_P0" ' the key
var_hivareg1 = "HKEY_CURRENT_USER"

strComputer = "."
Dim objSubkey
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")
Set oWSH = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("Shell.Application")
Set oEnv = oWSH.Environment("System")
Set WNetwork = CreateObject("WScript.Network")
strUserName = WNetwork.username

strSFN = WScript.ScriptFullName
strAppPath = Left(strSFN, InStrRev(strSFN, "\"))
sSourcePath = strAppPath
strLogLocation = oWSH.ExpandEnvironmentStrings(oEnv("TMP")) & "\" & Chr(33)
strLogFile = strLogLocation & "Thermoworks_ThermaDataLogger_SeriesII_Gen_P0_uninstall.log"
CurPackNum=0


If WSCript.Arguments.length = 0 Then
	oShell.ShellExecute "wscript.exe", Chr(34) & strSFN & Chr(34) & " uac", "", "runas", 1
Else
	
	Set oFile = oFSO.CreateTextFile(strLogFile, TRUE)
	oFile.WriteLine "Uninstallation Started at : " & CDATE(now)
  
	strAppUninstallPath1= oWSH.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{d201afcc-6b0a-420c-b96e-cf91000e2875}\QuietUninstallString")
	strAppUninstallPath2= oWSH.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\LOGERTMB&10C4&87FE\UninstallString")
	strAppUninstallPath3= oWSH.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\TDLCRADL&10C4&8213\UninstallString")
	
	Removex64APP
	Run  strAppUninstallPath1 & " /log " & strLogLocation & "38650u.log"
	Run  strAppUninstallPath2
	Run  strAppUninstallPath3
	 Set file1 = oFSO.GetFile(strLogLocation & "38650u_3_LoggerMSI.log")
	 file1.name = "!Thermoworks_ThermaDataLogger_3.4.16_Gen_P0_msi_uninstall.log"
	 Set file2 = oFSO.GetFile(strLogLocation & "38650u_2_ThermaData_Logger_Cradle_Drivers.msi.log")
	 file2.name = "!Thermoworks_ThermaDataLoggerCradleDrivers_1.0.0.0_Gen_P0_msi_uninstall.log"
	 Set file3 = oFSO.GetFile(strLogLocation & "38650u_1_ThermaData_Logger_Lead_Drivers.msi.log")
	 file3.name = "!Thermoworks_ThermaDataLoggerLeadDrivers_1.0.0.0_Gen_P0_msi_uninstall.log"
	 Set file3 = oFSO.GetFile(strLogLocation & "38650u_0_ThermaData_Logger_TMB_Drivers.msi.log")
	 file3.name = "!Thermoworks_ThermaDataLoggerTMBDrivers_1.0.0.0_Gen_P0_msi_uninstall.log"	
	 Set file4 = oFSO.GetFile(strLogLocation & "38650u.log")
	 file4.name = "!Thermoworks_ThermaDataLogger_3.4.16_Gen_P0_exe_Uninstall_summary.log"
	 
	DeleteARP "Thermoworks ThermaDataLogger SeriesII Gen P0"
	
	RemoveRemainingFiles
	
		'============== remove active setup registries ===================
	If strUserName = "SYSTEM" Then
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
	Set objRegistry=GetObject( "winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	strKeyPath10 = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	objRegistry.EnumKey HKEY_LOCAL_MACHINE, strKeyPath10, arrSubkeys1
	For Each objSubkey1 In arrSubkeys1
		strValueName1 = "ProfileImagePath"
		strSubPath1 = strKeyPath10 & "\" & objSubkey1
		objRegistry.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath1,strValueName1,strValue1
		correct_in strValue1
	Next
	Sub correct_in(dat_path)
		If oFSO.FileExists(dat_path & "\ntuser.dat") Then
			oWSH.Run "cmd /c REG LOAD HKLM\datastream """ & dat_path & "\NTUSER.DAT""", 2, True
			LocationPath = "Software\Microsoft\Active Setup\Installed Components\Thermoworks_ThermaDataLogger_SeriesII_Gen_P0"
			Call location_path("",Locationpath)
			oWSH.Run "cmd /c REG UNLOAD HKLM\datastream\", 2, True
		End If
	End Sub
	Sub SystemRun(ID)
		Call location_path(ID,"")
	End Sub
	'----------------------------------------------------------------------------------------------------------------
	Sub UserSettings
		Call location_path("","")
	End Sub
	'----------------------------------------------------------------------------------------------------------------
	Sub location_path(ID,location)
	If location = "" Then
		If ID = "" Then
			strKeyPath = "Software\Wow6432Node\Microsoft\Active Setup\Installed Components\Thermoworks_ThermaDataLogger_SeriesII_Gen_P0"
			objReg.DeleteKey HKEY_CURRENT_USER,strKeyPath
		Else
			strKeyPath = ID & "\Software\Wow6432Node\Microsoft\Active Setup\Installed Components\Thermoworks_ThermaDataLogger_SeriesII_Gen_P0"
			objReg.DeleteKey HKEY_USERS,strKeyPath
		End If
	Else
		strKeyPath = location
		objReg.DeleteKey HKEY_LOCAL_MACHINE,strKeyPath
	End If
	End Sub


'==========================================================================================================================


	oFile.WriteLine ""
	oFile.WriteLine ""
	oFile.WriteLine "Uninstallation completed successfully at : " & CDATE(now)

End If

'Start of the Run Subroutine
Sub Run  (sRunString)

	CurPackNum = CurPackNum +1
	oFile.WriteLine CSTR(CurPackNum) & " - " & CDATE(now) & " -  Currently Executing = " & sRunString 
	Set oExec = oWSH.Exec(sRunString)

	Do While oExec.Status = 0
		WScript.Sleep 100
	Loop
	'Start of Error Codes
	Select Case oExec.ExitCode
		Case 0
			WScript.Sleep 10
		Case 1605
			WScript.Sleep 10
		Case 1625
			WScript.Sleep 10
			oFile.WriteLine "The system administrator has set policies to prevent this installation. ExitCode = " & oExec.ExitCode & " " & CDATE(now)
			WScript.Quit oExec.ExitCode
		Case 3010
			WScript.Sleep 10
		Case Else
			If InStr(1,sRunString,"ThermaData Logger.exe",1) > 0 Or InStr(1,sRunString,"DriverUninstaller.exe",1) > 0 Or InStr(1,sRunString,"CP210xVCPInstaller_x64.exe",1) > 0 Then
				wscript.sleep 10
				oFile.WriteLine "Detected " & sRunString & " ExitCode = " & oExec.ExitCode & " " & CDATE(now)
				oFile.WriteLine ""
			Else
				wscript.sleep 30000
				oFile.WriteLine "Exiting Uninstallation - Failed at " & sRunString & " ExitCode = " & oExec.ExitCode & " " & CDATE(now)
				WScript.Quit oExec.ExitCode
			End If
	End Select
	'End of Error Codes
End Sub

'----------------------------------------------------------------------------------------------------------------

Sub DeleteARP(strProductName)

	Const HKEY_LOCAL_MACHINE = &H80000002
	Set objReg = GetObject("winmgmts:\\.\root\default:StdRegProv")
	strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & strProductName
	objReg.DeleteKey HKEY_LOCAL_MACHINE, strKeyPath
	oFile.WriteLine strProductName & " ARP deleted"

End Sub

'----------------------------------------------------------------------------------------------------------------


Sub Removex64APP
	Const HKEY_LOCAL_MACHINE = &H80000002
	PF64 = owsh.ExpandEnvironmentStrings("%ProgramW6432%")
	ProcArch = owsh.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
	WinDir = owsh.ExpandEnvironmentStrings("%SystemRoot%")
	If PF64 = "%ProgramW6432%" Then OSArch="x86" Else OSArch="x64"
	If OSArch = "x86" Then
		strAppUninstallPath4 = oWSH.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\D680DEE0F68D64EC53D0C5769879D15D387054CC\UninstallString")
		Run  strAppUninstallPath4 & " /S"
	End If	
	If OSArch="x64" Then
		Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
		objCtx.Add "__ProviderArchitecture", 64
		Set objLocator = CreateObject("Wbemscripting.SWbemLocator")
		Set objServices = objLocator.ConnectServer("","root\default","","",,,,objCtx)
		Set objReg = objServices.Get("StdRegProv") 
		If ProcArch="x86" Then
			strRegPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\D680DEE0F68D64EC53D0C5769879D15D387054CC"
			strRegName = "UninstallString"
			objReg.GetStringValue HKEY_LOCAL_MACHINE, strRegPath, strRegName, strAppUninstallPath4
			Run  strAppUninstallPath4 & " /S"
		Else
			strRegPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\D680DEE0F68D64EC53D0C5769879D15D387054CC"
			strRegName = "UninstallString"
			objReg.GetStringValue HKEY_LOCAL_MACHINE, strRegPath, strRegName, strAppUninstallPath4
			Run  strAppUninstallPath4 & " /S"
		End If		
	End If 
	
End Sub

Sub RemoveRemainingFiles
	PF64 = owsh.ExpandEnvironmentStrings("%ProgramW6432%")
	PF32 = owsh.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
	
	oFSO.DeleteFile PF64 & "\DIFX\E68C45B250901231\CP210xVCPInstaller_x64.exe", True
	
	var2 = PF64 & "\DIFX\E68C45B250901231"
	var1 = PF64 & "\DIFX"
	if(oFSO.GetFolder(var2).SubFolders.Count = 0) AND (oFSO.GetFolder(var2).Files.Count = 0) Then
		oFSO.DeleteFolder var2, True
	end if
	if(oFSO.GetFolder(var1).SubFolders.Count = 0) AND (oFSO.GetFolder(var1).Files.Count = 0) Then
		oFSO.DeleteFolder var1, True
	end if
	var3 = PF32 & "\ETI Ltd\ThermaData Logger"
	var4 = PF32 & "\ETI Ltd"
	var6 = PF32 & "\Silabs\MCU"
	var7 = PF32 & "\Silabs"
	if(oFSO.GetFolder(var3).SubFolders.Count = 0) AND (oFSO.GetFolder(var3).Files.Count = 0) Then
		oFSO.DeleteFolder var3, True
	end if

	if(oFSO.GetFolder(var4).SubFolders.Count = 0) AND (oFSO.GetFolder(var4).Files.Count = 0) Then
		oFSO.DeleteFolder var4, True
	end if
	if(oFSO.GetFolder(var6).SubFolders.Count = 0) AND (oFSO.GetFolder(var6).Files.Count = 0) Then
		oFSO.DeleteFolder var6, True
	end if
	if(oFSO.GetFolder(var7).SubFolders.Count = 0) AND (oFSO.GetFolder(var7).Files.Count = 0) Then
		oFSO.DeleteFolder var7, True
	end if
	
	windir = oWSH.ExpandEnvironmentStrings(oEnv("windir"))
	vbsWriteRegistry = windir & "\installer\CU_WriteRegistry.vbs"
	oFSO.DeleteFile vbsWriteRegistry, True

End Sub
