' --------------------------------------------------------------------------
'  File:         CopyUserData.vbs
'  Purpose:  	 used to copy files to user profile
'  Date:          28,June,2017
'  Description:
'  Usage: example usage:
'  wscript.exe "CopyUserData.vbs"
' --------------------------------------------------------------------------

Set objShell = CreateObject("Wscript.shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

source = session.property("USERDATA")
Destination = objShell.ExpandEnvironmentStrings("%localappdata%")

CopyFile source & "Local State", Destination & "\Google\Chrome\User Data\"
CopyFile source & "Secure Preferences", Destination & "\Google\Chrome\User Data\Default\"


sub CopyFile (Source, Destination)	
	If not objFSO.FolderExists(Destination) Then
		mkdirCommand = "cmd.exe /c mkdir " &  """" & Destination & """"
		objShell.Run mkdirCommand, 0, true
	end if
		shellCommand = "cmd.exe /c copy " &  """" & Source & """" & " " & """" & Destination & """" & " /Y"
		objShell.Run shellCommand, 0, true
End sub