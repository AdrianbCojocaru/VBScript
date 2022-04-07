 ' --------------------------------------------------------------------------
'  File:         ShortPath.vbs
'  Purpose:  used to solve hardcoded paths 
'  Date:          27 February,2014
'  Description:
'  Usage: example usage:
'  wscript.exe "ShortPath.vbs"
' ----------------------------------------------------------------------------

 Dim strLogText, MyPath, MyShortPath, FullPath, PartialPath, i
 Dim objFSO

 Set objFSO = CreateObject("Scripting.FileSystemObject")
 MyPath = Session.Property("APPDIR")
 If Not objFSO.FolderExists(MyPath) Then
    FullPath = Split(MyPath, "\", -1, 1)
    PartialPath = FullPath(0)
    For i = 1 To UBound(FullPath)
       PartialPath = PartialPath & "\" & FullPath(i)
       If Not objFSO.FolderExists(PartialPath) Then
          objFSO.CreateFolder(PartialPath)
       End If
      Next
 End If
 Set MyShortPath = objFSO.GetFolder(MyPath)
 Session.Property("SHORTINSTSL") = MyShortPath.ShortPath
 Set MyShortPath = nothing  
 Set objFSO = nothing