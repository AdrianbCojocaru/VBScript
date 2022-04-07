' --------------------------------------------------------------------------
'  File:         CopyUserData.vbs
'  Purpose:  	used to copy user files and folders to	each user profile
'  Date:          28,June,2017
'  Description:
'  Usage: example usage:
'  wscript.exe "CopyUserData.vbs"
' --------------------------------------------------------------------------

on error resume next
dim strComputer, fso
dim regkeycontents1, counter, temp, InitialString, sh
dim source, strKeyPath, strSubPath, objRegistry, arrSubkeys, strValueName, keyval
Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
set sh = CreateObject("wscript.shell")
Const regkey1 = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\SendTo"

source = session.property("USERDATA")

Set objRegistry=GetObject( "winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
objRegistry.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys, keyval, keyname

For Each objSubkey In arrSubkeys
	strValueName = "ProfileImagePath"
	strSubPath = strKeyPath & "\" & objSubkey
	If len(objSubkey) > 8 then
		keyname = "HKEY_USERS\" & objSubkey & "\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Local AppData"
		keyval = sh.RegRead (keyname)
		CopyFile source & "Local State", keyval & "\Google\Chrome\User Data\"
		CopyFile source & "Secure Preferences", keyval & "\Google\Chrome\User Data\Default\"
				
	End if
Next

Set fso = Nothing
set sh = Nothing


sub CopyFile (Source, Destination)	
	dim mkdirCommand, keyval, objShell
	set objShell = CreateObject("wscript.shell")
	Set objFSO = CreateObject( "Scripting.FileSystemObject")
	If not objFSO.FolderExists(Destination) Then
		mkdirCommand = "cmd.exe /c mkdir " &  """" & Destination & """"
		objShell.Run mkdirCommand, 0, true
	end if
		shellCommand = "cmd.exe /c copy " &  """" & Source & """" & " " & """" & Destination & """" & " /Y"
		objShell.Run shellCommand, 0, true
End sub