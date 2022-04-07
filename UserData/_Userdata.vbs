' ---------------------------------------------------------------------
'  File:         	_UserData.vbs
'  Purpose:  		Used to copy user application files.
'  Date:            06-11-2015
'  Description:		Copies user application files.
'  Usage: 			example usage: wscript.exe _UserData.vbs
' ---------------------------------------------------------------------

set osh = createobject("wscript.shell")

strFolderSource=session.property("USERDATA")

AppDataFolder = osh.expandenvironmentstrings("%APPDATA%")

strFolderDestination = AppDataFolder & "\"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strFolderSource)
CreateMissingFolder objFolder.Path

Set colFiles = objFolder.Files

For Each objFile In colFiles
  CopyIfMissing objFile.Path
Next
ShowSubFolders(objFolder)
 

Sub ShowSubFolders(objFolder)
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
   objFSO.CopyFile strSource,destinationTempFile
End If 

End Sub


Sub CreateMissingFolder(strFolder)
DestinationTempFolder = Replace(strFolder,strFolderSource,strFolderDestination)
If Not objFSO.FolderExists(DestinationTempFolder) Then 
      objFSO.CreateFolder DestinationTempFolder
End If 

End Sub