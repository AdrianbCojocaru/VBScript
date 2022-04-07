' --------------------------------------------------------------------------
'  Author:			AC - MS GmbH
'  Description:		used to remove firewalls rules
'  Usage:			wscript.exe "removeFirewall.vbs
' ----------------------------------------------------------------------------
Set oSH = CreateObject("WScript.Shell")
sCustPropValue = Session.Property("CustomActionData")
arrArgs = Split(sCustPropValue, ";", 2)
exe=arrArgs(0)
prop=arrArgs(1)
shellCommand01 = "cmd.exe /c netsh advfirewall firewall delete rule program=""" & exe & """ name=Firefox(""" & prop & """)"""
returnVal01 = osh.Run (shellCommand01, 0, True) 
Set oSH = Nothing