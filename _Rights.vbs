' --------------------------------------------------------------------------
'  Author: 			AC - MS GmbH
'  Description:		Sets rights on a msi folder.
'  Usage: 			example usage: wscript.exe _Rights.vbs
' --------------------------------------------------------------------------

Dim oSH, fso, f
Dim cmd, WrkDir, path
Dim returnVal

Set fso = CreateObject("Scripting.FileSystemObject")
Set oSH = CreateObject("Wscript.Shell")
WrkDir = oSH.expandenvironmentstrings("%windir%") & "\Temp"

path = session.property("CustomActionData")

Set f = fso.CreateTextFile(WrkDir + "\secpolicy.inf", True)
f.WriteLine "[Unicode]"
f.WriteLine "Unicode=No"
f.WriteLine "[Version]"
f.WriteLine "signature=" + QStr("$CHICAGO$")
f.WriteLine "Revision=1"
f.WriteLine "[File Security]"
f.Writeline Qstr(path) + ",0," + Qstr("D:AR(A;OICI;0x1301bf;;;BU)")
  f.WriteLine "[Registry Keys]"
  f.WriteLine QStr("MACHINE\SOFTWARE\Metro\Erik")+",0,"+Qstr("D:AR(A;CI;KA;;;BU)")
f.Close

cmd = "secedit.exe /configure /db " + QStr(WrkDir + "\secpolicy.sdb") + " /cfg " + QStr(WrkDir + "\secpolicy.inf") + " /overwrite /quiet"
returnVal = osh.Run (cmd, 0, true)
f = fso.DeleteFile(WrkDir + "\secpolicy.inf")
f = fso.DeleteFile(WrkDir + "\secpolicy.sdb")
If fso.FileExists(WrkDir + "\secpolicy.jfm") Then
    f = fso.DeleteFile(WrkDir + "\secpolicy.jfm")
End If

function QStr(S)
QStr = """" + S + """"
end function
