' --------------------------------------------------------------------------
'  File:         DeleteInstalldir.vbs
'  wscript.exe "DeleteInstalldir.vbs"
' ----------------------------------------------------------------------------
Set sh = CreateObject("Wscript.Shell")

path = Session.Property("INSTALLDIR")
sh.run "cmd /c rmdir /s /q " & """" & path & """", 0, True
'vpath = Session.Property("POSTGRESQL")
'sh.run "cmd /c rmdir /q " & """" & vpath & """", 0, True