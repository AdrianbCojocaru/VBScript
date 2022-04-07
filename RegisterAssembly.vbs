' --------------------------------------------------------------------------
'  Author: 			AC - MS GmbH
'  Description:		Used to register assemblys.
'  Usage: 			wscript.exe "RegisterAssembley.vbs"
' --------------------------------------------------------------------------
Dim Parameters, arr, RegAsmPath, DotNetInstallerPath, oSH
Dim dll, dll2, dll3, dll4, dll5, dll6

Parameters=session.property("CustomActionData")
arr=split(Parameters,";",-1,1)

RegAsmPath = arr(0) & "regasm.exe"
iDir = arr(1)

Set osh = CreateObject("WScript.Shell")

osh.run """" & RegAsmPath & """" & " " & """" & iDir & "Plugins\AnalysisOffice\BiApi.dll" & """" & " /Silent /tlb:" & """" & iDir & "Plugins\AnalysisOffice\BiApi.tlb" & """", 0, true
osh.run """" & RegAsmPath & """" & " " & """" & iDir & "Plugins\AnalysisOffice\BiCore.dll" & """" & " /Silent /Codebase /tlb:" & """" & iDir & "Plugins\AnalysisOffice\BiCore.tlb" & """", 0, true
osh.run """" & RegAsmPath & """" & " " & """" & iDir & "Plugins\EPMAddin\FPMXLClient.OlapUtilities.dll" & """" & " /Silent /Codebase /tlb:" & """" & iDir & "Plugins\EPMAddin\FPMXLClient.OlapUtilities.tlb" & """", 0, true
osh.run """" & RegAsmPath & """" & " " & """" & iDir & "Plugins\EPMAddin\FPMXlClient.dll" & """" & " /Silent /Codebase /tlb:" & """" & iDir & "Plugins\EPMAddin\FPMXlClient.tlb" & """", 0, true
osh.run """" & RegAsmPath & """" & " " & """" & iDir & "Plugins\AnalysisOffice\BiExcelBase.dll" & """" & " /Silent /Codebase /tlb:" & """" & iDir & "Plugins\AnalysisOffice\BiExcelBase.tlb" & """", 0, true
osh.run """" & RegAsmPath & """" & " " & """" & iDir & "Plugins\EPMAddin\EPMOfficeActiveX.dll" & """" & " /Silent /Codebase", 0, true
osh.run """" & RegAsmPath & """" & " " & """" & iDir & "CofInterfaces.dll" & """" & " /Silent /Codebase /tlb:" & """" & iDir & "CofInterfaces.tlb" & """", 0, true
osh.run """" & RegAsmPath & """" & " " & """" & iDir & "ApplicationBuilderComBridge.dll" & """" & " /Silent /Codebase", 0, true