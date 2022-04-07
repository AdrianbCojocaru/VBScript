Dim pathsdb
Set oSH = CreateObject("WScript.Shell")
pathsdb = Session.Property("CustomActionData")

command = "cmd.exe /c sdbinst.exe -u " & """" & pathsdb  & """"

oSH.Run command, 0 , True