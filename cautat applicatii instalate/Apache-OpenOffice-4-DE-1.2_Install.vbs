'------------------------------------------------------------------------------
' Description	:	Install Script for Open Office
'------------------------------------------------------------------------------
Option Explicit 
' Variable Declarations -------------------------------------------------------
Const msiInstallStateAbsent = 2
Const msiUILevelNone = 2
Const isDebug = False
Const ForWriting = 2
Const ForReading = 1
Const ForAppending = 8
Const HKLM = &H80000002, HKCU = &H80000001, HKCR = &H80000000
Const LogFileName = "OpenOffice_UninstallPrevVersions_summary.log", Max_LogSize = 100
Const SWRegRoot = "SOFTWARE\"
Const SWRegRoot64 = "SOFTWARE\Wow6432Node\"
Const UninstallRegPath = "Microsoft\Windows\CurrentVersion\Uninstall\"

Const NDLaunchRegPath = "ManageSoft Corp\ManageSoft\Launcher\CurrentVersion"
Const NDLaunchRegValue = "PathName"
' EndConst

'Region Global INIT
Dim oWSO, oFSO, oREG, oWMI, cList, sCurDir, sLogFile,objArgsm, objReg, sParentFolder, objArgs
Dim scurdirbak, slogfilename, colitems, objitem, bforce, brun, arrpacklist, strsqlquery, i, strprefix
Dim colproduct, oproduct, arrproduct, o_installer, str_guid, bx64, arr_PackUninst, arr_3rdUninstall, str_cmdLine, i_res, str_PackUninst
Dim rVal, strArg, vcred2008, ooffice4, oofficelpg4, INSTALLDIR,  sInstallCommand, Parameters
Set oWSO= CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oREG = GetObject("winmgmts:\\.\root\default:StdRegProv")
Set oWMI = GetObject("winmgmts:\\.\root\CIMV2")
Set cList = oWMI.ExecQuery _
	("Select * from Win32_Process Where Name = 'Process.exe'")
	
Set objArgs = WScript.Arguments
Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
sParentFolder  = Replace(WScript.ScriptFullName, WScript.ScriptName, "")

sCurDir = oFSO.GetParentFolderName(WScript.ScriptFullName)
sCurDirBak = oWSO.CurrentDirectory
oWSO.CurrentDirectory = sCurDir

'Region Init log file
	sLOGFileName = oWSO.ExpandEnvironmentStrings("%temp%") & "\" & LogFileName

	If oFSO.FileExists(sLOGFileName) Then
		If oFSO.GetFile(sLOGFileName).Size / 1024 >= Max_LogSize Then
			oFSO.CopyFile sLOGFileName, sLOGFileName & ".bak", vbTrue
			oFSO.DeleteFile sLOGFileName,vbTrue
		End If 		
	End If
'EndRegion
	if objArgs.Count > 0 then
		For Each strArg in objArgs
			If (InStr(strArg, "SOURCELIST=") = 0) Then Parameters = Parameters & " " & strArg
		Next
	End If

'EndRegion

	Logwrite , " Start '" & WScript.ScriptName & "'"
	Logwrite , "OS Architecture is 32-bit" 
	bForce = True
	bRun = vbTrue
	
'Uninstall MSI (Based on ProductName) and Packages if MSI was installed with RayManagesoft Package(Uninstall string like Launcher\ndlaunch * -d "<PackageName>)
		
		strSQLQuery =  " WHERE Name Like 'Apache Open Office 4.0 DE' OR Name Like 'OpenOffice - German'"
		Set colProduct = oWMI.ExecQuery("Select IdentifyingNumber,Name,Version From Win32_Product"&strSQLQuery, , 48)

		Logwrite 9999, "ProductCode"&vbTab&vbTab&"Version"&vbTab&vbTab& "ProductName" 
		For Each oProduct In colProduct
			
			If Not IsArray(arrProduct) Then 
				ReDim arrProduct(0)
			Else	
				ReDim Preserve arrProduct(UBound(arrProduct)+1)
			End If
			arrProduct(UBound(arrProduct)) = oProduct.IdentifyingNumber
			Logwrite 9999,oProduct.IdentifyingNumber&vbTab&oProduct.Version&vbTab&oProduct.Name
		Next
		If IsArray(arrProduct) Then
				UninstallMSIProducts(arrProduct)
		End If

	' Run the install -------------------------------------------------------------

		objReg.GetStringValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{9BE518E6-ECC6-35A9-88E4-87755C07200F}", "DisplayVersion", vcred2008
		If (IsNull(vcred2008) OR Not (vcred2008="9.0.30729.6161")) Then 
			  sInstallCommand = "msiexec.exe /i """ & sParentFolder & "IBM1\MS-VCPP2008Redist-9.0.30729.6161-EN-1.0.msi""" &_
					 " TRANSFORMS=""" & sParentFolder & "IBM1\MS-VCPP2008Redist-9.0.30729.6161-EN-1.0.mst"" " &_
					  "SOURCELIST=""" & sParentFolder & "IBM1""" & Parameters & " /qn"
		  rVal = oWSO.Run(sInstallCommand, 1, TRUE)
		  If rVal <> 0 And rVal <> 3010 And rVal <> 1641 then
			WScript.Quit rVal
		  End If
		End If  
		
		
		objReg.GetStringValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{24B89186-2A56-4D28-B930-6F4FCF224E2F}", "DisplayVersion", ooffice4
		If (IsNull(ooffice4) OR Not (ooffice4="4.01.9714")) Then 
			  sInstallCommand = "msiexec.exe /i """ & sParentFolder & "IBM2\Apache-OpenOffice-4.0.1-EN-1.1.msi""" &_
					 " TRANSFORMS=""" & sParentFolder & "IBM2\Apache-OpenOffice-4.0.1-EN-1.1.mst"" " &_
					  "SOURCELIST=""" & sParentFolder & "IBM2""" & Parameters & " /qn"
		  rVal = oWSO.Run(sInstallCommand, 1, TRUE)
		  If rVal <> 0 And rVal <> 3010 And rVal <> 1641 then
			WScript.Quit rVal
		  End If
		End If  

		objReg.GetStringValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{0C55CCF1-29E2-4481-A31F-1FDF19E038F2}", "DisplayVersion", oofficelpg4
		If (IsNull(oofficelpg4) OR Not (oofficelpg4="4.01.9714")) Then 
			  sInstallCommand = "msiexec.exe /i """ & sParentFolder & "IBM3\Apache-OpenOfficeLanguagePackGerman-4.0.1-DE-1.0.msi""" &_
					 " TRANSFORMS=""" & sParentFolder & "IBM3\Apache-OpenOfficeLanguagePackGerman-4.0.1-DE-1.0.mst"" " &_
					  "SOURCELIST=""" & sParentFolder & "IBM3""" & Parameters & " /qn"
		  rVal = oWSO.Run(sInstallCommand, 1, TRUE)
		  If rVal <> 0 And rVal <> 3010 And rVal <> 1641 then
			WScript.Quit rVal
		  End If
		End If  

	'------------------------------------------------------------------------------
	' Post Install Tasks ----------------------------------------------------------
	'
		oWSO.RegWrite "HKLM\Software\MSITS\Apache-OpenOffice-4-bundle\Package Version","1.2"
		 objReg.GetStringValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\sbase.exe", "Path", INSTALLDIR
		If ofso.FileExists(sParentFolder & "Apache-OpenOffice-W7.x86-4-EN-1.2_Uninstall.vbs") And ofso.FolderExists(INSTALLDIR) Then
		   ofso.CopyFile sParentFolder & "Apache-OpenOffice-W7.x86-4-EN-1.2_Uninstall.vbs", INSTALLDIR
		End If  
		
	'
	'------------------------------------------------------------------------------

Quit(0)

'------------------------------------------End Script--------------------------------------------------


Function GetPackageUninstCMD (str_UninstReg,str_Entry)
	Dim str_Temp,str_cmdLine,arr_temp
'	GetPackageUninstCMD = ""
	str_cmdLine = Null
	oREG.GetStringValue HKLM, SWRegRoot &UninstallRegPath& str_UninstReg, str_Entry, str_Temp
	Select Case str_Entry
		Case "UninstallString"
			If InStr (LCase(str_Temp), "launcher\ndlaunch") > 0 And InStr(str_Temp, " -d ") > 0 Then
				str_cmdLine = str_Temp
				Logwrite , "Uninstall String for Package:" & str_cmdLine
			End If		
		Case Else
			If InStr (LCase(str_Temp), "msiexec.exe") = 0 And Len(str_Temp) > 0	Then 
				str_cmdLine = str_Temp
				Logwrite , "3rdPartyUninstallString String for Package:" & str_cmdLine					
			End If
	End Select
	If Not IsNull (str_cmdLine) Then 
		arr_temp= AddtoArray (str_cmdLine,arr_temp)
	End If

	GetPackageUninstCMD = arr_temp
End Function

Function AddtoArray(str_var, arr_var)
	If IsArray(arr_var) Then
		ReDim	Preserve arr_var(UBound(arr_var) + 1)
		arr_var(UBound(arr_var))=str_var
	Else
		ReDim arr_var(0)
		arr_var(0)=str_var	
	End If
AddtoArray = arr_var

End Function

Sub UninstallMSIProducts(arr_ProductCode)
	Set o_Installer = CreateObject("WindowsInstaller.Installer")
	For Each str_guid In arr_ProductCode
		arr_PackUninst = GetPackageUninstCMD(str_Guid,"UninstallString")
		If o_Installer.ProductState(str_Guid)  >= 0 Then
			str_cmdLine = "msiexec /qn /norestart /x"&str_guid&" /l*v ""%temp%\"&str_guid&".log""" 			
			Logwrite 9999, "Uninstall MSI. Command:'" &str_cmdLine&"'."
			If bRun Then 
				i_res = oWSO.Run (str_cmdLine, 0, vbTrue)
			Else 
				i_res = 0
			End If
			If i_res <> 0 And i_res <> 3010 And i_res <> 1605 And i_res <> 1641 Then
				Logwrite i_res, "Uninstallation of Product with ProductCode " & str_guid & " returned an error .Skip uninstall." 
			Else 	
				Logwrite 9999, "Product  with ProductCode " & str_guid & " successfully uninstalled."
				If bRun Then 
					oREG.DeleteKey HKLM, SWRegRoot & UninstallRegPath & str_guid
				End If
				If IsArray(arr_PackUninst) Then
					For Each str_PackUninst In arr_PackUninst 
						Logwrite 9999, "MSI was installed with RayManagesoft. Try uninstall package"
						If bRun Then
							i_res = oWSO.Run (str_PackUninst, 0, vbTrue)
							if i_res <> 0 then Logwrite i_res , "Package uninstall '" & str_PackUninst & "' failed."
						End If 
					Next
				End If	
			End If
		Else 	
			Logwrite 9999, "Product with ProductCode " & str_guid & " isn't installed."
' If /Force - Try uninstall RayManagesoftPackage and remove RegKey Microsoft\Windows\CurrentVersion\Uninstall\<ProductCode>
				arr_PackUninst = GetPackageUninstCMD(str_Guid,"UninstallString")
				If IsArray(arr_PackUninst) Then
					For Each str_PackUninst In arr_PackUninst 
						Logwrite 9999, "MSI with ProductCode '" & str_guid & "' isn't installed but RegKey is remained. Try cleanup package"
						If bRun Then
							oREG.DeleteValue HKLM, SWRegRoot & UninstallRegPath & str_guid, "3rdPartyUninstallString"
						End If
						Logwrite 9999, "Uninstall Package:'" & str_PackUninst & "'."
						If bRun Then
							ires = oWSO.Run (str_PackUninst, 0, vbTrue)
							if ires <> 0 then Logwrite ires , "Package uninstall '" & str_guid & "' failed."
						End If 
					Next
						
				End If
				Logwrite 9999, "Remove regKey '" & SWRegRoot & UninstallRegPath & str_guid & "'."
				If bRun Then
					oREG.Deletekey HKLM, SWRegRoot & UninstallRegPath & str_guid
				End If
				
		End If
	Next
	Set o_Installer = Nothing

End Sub

Sub Logwrite (errnum, errdesc)
	Dim o_Logfile, str_Outstream
	If Not IsNumeric(errnum) Then
		str_Outstream = Now & ": " & errdesc
	Else
		If errnum = 9999 Then 
			str_Outstream = errdesc
		Else
			str_Outstream = Now & ": " & errdesc & " ReturnCode: " & errnum
		End If
	End If
	Set o_Logfile = oFSO.OpenTextFile(sLOGFileName , ForAppending, vbTrue)
	o_Logfile.WriteLine str_Outstream
	o_Logfile.Close
	
End Sub

Sub Quit(iQuit)
	oWSO.CurrentDirectory = sCurDirBak
	Logwrite iQuit,"Exit script. '" & WScript.ScriptName & "'"
	WScript.Quit(iQuit)
End Sub