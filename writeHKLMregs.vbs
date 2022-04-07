RegWrite "64Bit", "SOFTWARE\Metro\MPOS", "Server", "DEF01MPSU000000"

Function RegWrite(strArhitecture, strRegKeyPath, strRegValName, strValue)
	
	hexRegRoot = &H80000002
    strComputer = "."
    Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet") 
	strArhitecture = lcase (strArhitecture)
	Select Case strArhitecture
	   Case "32bit" 
		objCtx.Add "__ProviderArchitecture", 32	
	   Case "64bit" 
		objCtx.Add "__ProviderArchitecture", 64 
	End Select
    Set objLocator = CreateObject("Wbemscripting.SWbemLocator") 
    Set objServices = objLocator.ConnectServer("","root\default","","",,,,objCtx) 
    Set objReg = objServices.Get("StdRegProv")
	
	SubKeyExists = False
    strRegKeyPathNBS = Left(strRegKeyPath, InStrRev(strRegKeyPath, "\") - 1) 
    strSubKeyToMatch = Right(strRegKeyPath, Len(strRegKeyPath) - InStrRev(strRegKeyPath, "\"))
    objReg.EnumKey hexRegRoot, strRegKeyPathNBS, arrSubKeys
    If IsArray(arrSubKeys) Then

       For Each strSubKey In arrSubKeys
          If strSubKey = strSubKeytoMatch Then 
          	  SubKeyExists = True
          End If
       Next
    End If
	
     If Not SubKeyExists Then
         objReg.CreateKey hexRegRoot, strRegKeyPath
      End If
	statusCode = objReg.SetStringValue(hexRegRoot, strRegKeyPath, strRegValName, strValue) 
End Function