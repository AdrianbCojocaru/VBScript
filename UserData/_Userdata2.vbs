' ---------------------------------------------------------------------
'  File:         	_UserData.vbs
'  Purpose:  		Used to copy user application files.
'  Description:		Copies user application files.
'  Usage: 			example usage: wscript.exe _UserData.vbs
' ---------------------------------------------------------------------

set osh = createobject("wscript.shell")

strFolderSource = osh.expandenvironmentstrings("%windir%") & "\Installer\InterPlot Organizer Connect 10 DE\_UserData\"
'MsgBox strFolderSource

AppDataFolder = osh.expandenvironmentstrings("%LOCALAPPDATA%")

strFolderDestination = AppDataFolder & "\"
'MsgBox strFolderDestination

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strFolderSource)
CreateMissingFolder objFolder.Path

Set colFiles = objFolder.Files

For Each objFile In colFiles
  CopyIfMissing objFile.Path
Next
ShowSubFolders(objFolder)
ShowSubFolders(objFolder)
 

Sub ShowSubFolders(objFolder)
'msgbox objFolder
  Set colFolders = objFolder.SubFolders
  For Each objSubFolder In colFolders
  CreateMissingFolder objSubFolder.Path
    Set colFiles = objSubFolder.Files
    For Each objFile In colFiles
      CopyIfMissing objFile.Path
    Next
    ShowSubFolders(objSubFolder)
  Next
End Sub

Sub CopyIfMissing(strSource)
destinationTempFile = Replace(strSource, strFolderSource, strFolderDestination, 1, -1, 1)

If Not objFSO.FileExists(destinationTempFile) Then
	CreateMissingFolder(strFolderDestination)
	msgbox "lala " & destinationTempFile
   objFSO.CopyFile strSource, destinationTempFile
End If 

End Sub


Sub CreateMissingFolder(strFolder)
msgbox "CreateMissingFolder "strFolder
DestinationTempFolder = Replace(strFolder,strFolderSource,strFolderDestination)
If Not objFSO.FolderExists(DestinationTempFolder) Then 
      objFSO.CreateFolder DestinationTempFolder
End If 

End Sub