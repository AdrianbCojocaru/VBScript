 Set objFSO = CreateObject("Scripting.FileSystemObject")
strLocFis = "C:\Program Files (x86)\ARM"
Set ShortPath = objFSO.GetFolder(strLocFis)
msgbox shortpath.ShortPath