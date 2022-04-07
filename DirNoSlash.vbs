' --------------------------------------------------------------------------
'  Author: 			AC - MS GmbH
'  Description:		Used to remove the last backslash.
'  Usage: 			example usage: wscript.exe DirNoSlash.vbs
' --------------------------------------------------------------------------
str1 = Session.Property("INSTALLDIR")
'str2 = Session.Property("PTV_AG1")
Session.Property("DIRNOSLASH") = left(str1,len(str1)-1)
'Session.Property("VENDORNS") = left(str2,len(str2)-1)