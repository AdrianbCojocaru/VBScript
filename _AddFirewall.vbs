' --------------------------------------------------------------------------
'  Description:		used to add firewalls rules
'  Usage:			wscript.exe "addFirewall.vbs
' ----------------------------------------------------------------------------
Set oSH = CreateObject("WScript.Shell")
sCustPropValue = Session.Property("CustomActionData")
arrArgs = Split(sCustPropValue, ";", 2)
exe=arrArgs(0)
prop=arrArgs(1)
shellCommand01 = "cmd.exe /c ""netsh firewall add allowedprogram program=""" & exe & """ profile=STANDARD name=Firefox(""" & prop & """)"""
returnVal01 = osh.Run (shellCommand01, 0, True)
Set oSH = Nothing 