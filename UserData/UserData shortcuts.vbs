' --------------------------------------------------------------------------

'  Usage: wscript.exe UserData.vbs"

' ----------------------------------------------------------------------------
On Error Resume Next
Set objShell = CreateObject("Wscript.shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")


'PersonalFolder = objShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Personal")

'If PersonalFolder = "" Then
    USERPROFILE = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
	PersonalFolder = USERPROFILE & "\Documents"
'End if


CopyFolder session.property("IBM1"), PersonalFolder & "\IBM\"

sub CopyFolder (Source, Destination)
	Set objShell = CreateObject("WScript.Shell")
	If Right(Source, 1) ="\" then Source = Left(Source, len(Source)-1)
	shellCommand = "cmd.exe /c xcopy " &  """" & Source & """" & " " & """" & Destination & """" & " /E /F /R /H /I /Y"
	objShell.Run shellCommand, 0, true
End sub



Set Shell = CreateObject("WScript.Shell")
DesktopPath = Shell.SpecialFolders("Desktop")


Set link = Shell.CreateShortcut(DesktopPath & "\as400_mtb_1.lnk")
link.Arguments = ""
link.IconLocation = PersonalFolder & "\IBM\iAccessClient\Emulator\pcsws.exe,0"
link.TargetPath = "%USERPROFILE%\Documents\IBM\iAccessClient\Emulator\as400_mtb_1.hod"
link.WindowStyle = 1
link.Save
Set link = Nothing


Set link = Shell.CreateShortcut(DesktopPath & "\as400_mtb_2.lnk")
link.Arguments = ""
link.IconLocation = PersonalFolder & "\IBM\iAccessClient\Emulator\pcsws.exe,0"
link.TargetPath = "%USERPROFILE%\Documents\IBM\iAccessClient\Emulator\as400_mtb_2.hod"
link.WindowStyle = 1
link.Save
Set link = Nothing


Programs = Shell.SpecialFolders("Programs")
CreatePath Programs & "\IBM Access Client Solution"

Set link = Shell.CreateShortcut(Programs & "\IBM Access Client Solution\as400_mtb_1.lnk")
link.Arguments = ""
link.IconLocation = PersonalFolder & "\IBM\iAccessClient\Emulator\pcsws.exe,0"
link.TargetPath = "%USERPROFILE%\Documents\IBM\iAccessClient\Emulator\as400_mtb_1.hod"
link.WindowStyle = 1
link.Save
Set link = Nothing


Set link = Shell.CreateShortcut(Programs & "\IBM Access Client Solution\as400_mtb_2.lnk")
link.Arguments = ""
link.IconLocation = PersonalFolder & "\IBM\iAccessClient\Emulator\pcsws.exe,0"
link.TargetPath = "%USERPROFILE%\Documents\IBM\iAccessClient\Emulator\as400_mtb_2.hod"
link.WindowStyle = 1
link.Save
Set link = Nothing


Sub CreatePath(strPath)
	Dim arrFolders, Folder, strNewPath
	Set oFS = CreateObject("Scripting.FileSystemObject")
	On Error Resume Next
	If Not oFS.FolderExists(strPath) Then
		arrFolders = split(strPath,"\")
		For Each Folder in arrFolders
			strNewPath = strNewPath & Folder
			If not oFS.FolderExists(strPath) Then
				oFS.CreateFolder(strnewPath)
				
			End If
			strNewPath = strNewPath & "\"
		Next
	End If
End Sub