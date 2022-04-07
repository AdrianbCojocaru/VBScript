On error resume next
Dim pathsdb
pathsdb = Session.Property("CustomActionData")
Set oSH = CreateObject("WScript.Shell")
command = "cmd.exe /c sdbinst.exe -q " & """" & pathsdb  & """"
oSH.Run command, 0 , True

'
Const HKLM = &H80000002
strComputer = "."

Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")
osh.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{8e66c681-ff4a-41f9-9f35-e3affe9c024b}.sdb\SystemComponent", 1, "REG_DWORD"