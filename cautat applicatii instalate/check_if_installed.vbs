dim strSQLQuery, oWMI, colProduct,oFSO, objFolder,strUninstCommandLine2
dim strProductCode1, strProductCode2, returnVal,sInstallCommand
dim oWSO, oReg, strComputer, strPackageName, strCurrentDir, i_res, pcode
dim execuninst1, execuninst2, strLogFileLocation 
strComputer = "."
Const HKLM = &H80000002
Const app1 = "xxx32"
Const app2 = "InstEd 1.5.7.15"
CONST ForReading  = 1
CONST ForWritting = 2
CONST ForAppending = 8

Set oWSO = CreateObject("WScript.Shell")
Set ofso = CreateObject("Scripting.FileSystemObject")
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
oReg.EnumKey HKLM, strKeyPath, arrSubKeys
strCurrentDir = Left(Wscript.ScriptFullName, (InstrRev(Wscript.ScriptFullName, "\") -1))
strLogFileLocation = oWSO.ExpandEnvironmentStrings("%windir%") & "\temp"
sLOGFileName ="""" & strLogFileLocation & "\" & "OpenOffice_PreviousVersion_uninstall_summary.txt" & """"

'oReg.GetStringValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"&pcode, "DisplayName", pname1
'msgbox pname1
For Each subkey In arrSubKeys
	oReg.GetStringValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"&subkey, "DisplayName", pname
	If UCase(app1) = UCase(pname) then
		strProductCode1 = subkey
		'oReg.GetStringValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"&subkey, "DisplayName", pname
		strUninstCommandLine1 = "msiexec.exe /x " & strProductCode1 & " /qn /l*v " & """" & strLogFileLocation & "\" & app1 & "_Uninstall.log" & """"
		'msgbox strUninstCommandLine2
		returnVal1 = oWSO.Run (strUninstCommandLine2, 0, vbTrue)
			If returnVal1 <> 0 And returnVal1 <> 3010 And returnVal1 <> 1605 And returnVal1 <> 1641 Then
				content = "Uninstallation of Product with ProductCode " & strProductCode1 & " returned an error .Skip uninstall.Please check"  & app1& "_Uninstall.log" & """"
			Else 	
				content = "Product  with ProductCode " & strProductCode1 & " successfully uninstalled."
			End if
			WriteFile (content)
	End if
	If UCase(app2) = UCase(pname) then
		strProductCode2 = subkey
		strUninstCommandLine2 = "msiexec.exe /x " & strProductCode2 & " /qn /l*v " & """" & strLogFileLocation & "\" & app2 & "_Uninstall.log" & """"
		'msgbox strUninstCommandLine2
		returnVal2 = oWSO.Run (strUninstCommandLine2, 0, vbTrue)
		msgbox returnVal2
			If returnVal12 <> 0 And returnVal12 <> 3010 And returnVal12 <> 1605 And returnVal12 <> 1641 Then
				content = "Uninstallation of Product with ProductCode " & strProductCode2 & " returned an error .Skip uninstall.Please check"  & app2 & "_Uninstall.log" & """"
			Else 	
				content = "Product  with ProductCode " & strProductCode2 & " successfully uninstalled."
			End if
			WriteFile (content)
   End if
Next

function WriteFile (content)
			set file = oFSO.CreateTextFile(sLOGFileName)
			strPut =content
			file.Write strPut
			File.close
end function
