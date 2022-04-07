Dim pathsdb, strRegKeyPath, strRegValName
Set oSH = CreateObject("WScript.Shell")
pathsdb = Session.Property("CustomActionData")


command = "cmd.exe /c sdbinst.exe -q " & """" & pathsdb  & """"
oSH.Run command, 0 , True


'write x64
Const HKLM = &h80000002
Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
objCtx.Add "__ProviderArchitecture", 64
objCtx.Add "__RequiredArchitecture", TRUE
Set objLocator = CreateObject("Wbemscripting.SWbemLocator")
Set objServices = objLocator.ConnectServer("","root\default","","",,,,objCtx)
Set objStdRegProv = objServices.Get("StdRegProv") 
' Use ExecMethod to call the SetDWORDValue method
Set Inparams = objStdRegProv.Methods_("CreateKey").InParameters
Inparams.Hdefkey = HKLM
Inparams.Ssubkeyname = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{6dedf981-203d-4793-a5f6-fc196e49880b}.sdb"
set Outparams = objStdRegProv.ExecMethod_("CreateKey", Inparams,,objCtx)


Set Inparams = objStdRegProv.Methods_("SetDWORDValue").InParameters
Inparams.Hdefkey = HKLM
Inparams.Ssubkeyname = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{6dedf981-203d-4793-a5f6-fc196e49880b}.sdb"
Inparams.Svaluename = "SystemComponent"
Inparams.uValue = 1
set Outparams = objStdRegProv.ExecMethod_("SetDWORDValue", Inparams,,objCtx)