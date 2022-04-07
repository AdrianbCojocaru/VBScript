' --------------------------------------------------------------------------
'  File:         userdata.vbs
'  Purpose:  	used for user settings
'  Date:          21,July,2014
'  Description:
'  Usage: example usage:
'  wscript.exe "userdata.vbs"
' ----------------------------------------------------------------------------
set osh = createobject("wscript.shell")

strFolderSource=session.property("USERDATA")
aa = osh.expandenvironmentstrings("%APPDATA%")

strFolderDestination = aa & "\"
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
destinationTempFile = Replace(lcase(strSource),lcase(strFolderSource),strFolderDestination)

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