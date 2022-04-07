Const HKEY_CURRENT_USER = &H80000001
strComputer = "."

Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")

strKeyPath = "Software\Adobe\Dreamweaver CS6\Settings"
objReg.CreateKey HKEY_CURRENT_USER,strKeyPath

strEntryName1 = "initialFileTypeDlg"
strValue = "TRUE"
objReg.SetStringValue HKEY_CURRENT_USER,strKeyPath,strEntryName1,strValue
objReg.SetStringValue HKEY_CURRENT_USER,strKeyPath,strEntryName2,strValue
