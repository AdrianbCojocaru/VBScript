' --------------------------------------------------------------------------
'  Copyright (C) 2013
'  File:         RemoveActiveSetupRegistries.vbs
'  Description: Used to remove active sertup registries for selfhealing
' --------------------------------------------------------------------------
Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
	strComputer & "\root\default:StdRegProv")

strProductCode = Session.Property("ProductCode")
strKeyPath = "SOFTWARE\Microsoft\Active Setup\Installed Components"
objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys

For Each Subkey in arrSubKeys
	If Instr(1, Subkey, strProductCode) > 0 Then
		objReg.DeleteKey HKEY_LOCAL_MACHINE, strKeyPath & "\" & Subkey
	End If
Next