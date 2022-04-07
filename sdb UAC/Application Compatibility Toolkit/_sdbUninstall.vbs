On error resume next
Dim pathsdb
pathsdb = Session.Property("CustomActionData")
Set oSH = CreateObject("WScript.Shell")
command = "cmd.exe /c sdbinst.exe -u " & """" & pathsdb  & """"
oSH.Run command, 0 , True

'