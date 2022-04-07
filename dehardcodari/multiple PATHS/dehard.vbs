' --------------------------------------------------------------------------
'  File:         dehard.vbs
'  Purpose:  used for hardcoded files 
'  Date:          22,August,2013
'  Description:
'  Usage: example usage:
'  wscript.exe "dehard.vbs"
' ----------------------------------------------------------------------------
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim strFilePath, strToReplace, strNewValue, strArgs, arrArgs
Dim SHOST, DSPORT, FSPORT, SDSPORT, ADD1, PORT1, ADD2, PORT2, ADD3, ADD4, ADD5, ADD6
strArgs = Session.Property("CustomActionData")
arrArgs = Split(strArgs, ";", -1, 1)

strLocFis = arrArgs(0) 'installdir
SHOST = arrArgs(1)  'SHOST	10.111.11.36
DSPORT = arrArgs(2) 'DSPORT	9898
FSPORT = arrArgs(3) 'FSPORT 9899
SDSPORT = arrArgs(4)'SDSPORT	80
ADD1 =  arrArgs(5)	'ADD1	129.148.70.86
PORT1 =  arrArgs(6)	'PORT1	10000
ADD2 = arrArgs(7)	'ADD2 127.0.0.1
PORT2 = arrArgs(8)	'PORT2	8899
ADD3 = arrArgs(9)	'ADD3	0.0.0.0
ADD4 = arrArgs(10)	'ADD4	1.7.1.103
ADD5 = arrArgs(11)	'ADD5 12.34.56.78
ADD6 = arrArgs(12)	'ADD6 17.0.5.20

Function Ascii(strToReplace, strNewValue, strFilePath)

	Dim objFSO, objFile, strText, re
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	if objFSO.FileExists(strFilePath) Then
		Set objFile = objFSO.OpenTextFile(strFilePath, ForReading, True)
			strText = objFile.ReadAll
		objFile.Close
		Set objFile = Nothing
		
		strText = Replace(strText, strToReplace, strNewValue, 1, -1, 0)
	
		Set objFile = objFSO.CreateTextFile(strFilePath, True)
		objFile.Write strText
		objFile.Close
		Set objFile = Nothing
	
	End If

	Set objFSO = Nothing

End Function


strToReplace = "C:\Program Files (x86)\PTC\Creo Elements\Direct Manager Server 17.0"
strNewValue = Left(strLocFis, (Len(strLocFis) - 1))
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\config\appclient.bat"
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\config\asenv.bat"
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\bin\capture-schema.bat"
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\bin\asadmin.bat"
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\bin\asant.bat"
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\bin\asapt.bat"
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\bin\asupgrade.bat"
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\bin\jspc.bat"
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\bin\package_appclient.bat"
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\bin\schemagen.bat"
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\bin\verifier.bat"
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\bin\wscompile.bat"
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\bin\wsdeploy.bat"
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\bin\wsdeploy.bat"
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\bin\wsgen.bat"
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\bin\wsimport.bat"
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\bin\xjc.bat"
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\updatecenter\bin\updatetool.bat"
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\updatecenter\config\ucenv.bat"
Ascii strToReplace, strNewValue, strLocFis & "WebServicesServer\imq\etc\imqenv.conf"

strToReplace = "10.111.11.36"
Ascii strToReplace, SHOST, strLocFis & "config\custom.xml"
Ascii strToReplace, SHOST, strLocFis & "custom.xml"
Ascii strToReplace, SHOST, strLocFis & "local\db_defaults"
Ascii strToReplace, SHOST, strLocFis & "clntwin\mmbuild.nsi"
Ascii strToReplace, SHOST, strLocFis & "clntwin\sapintegrationbuild.nsi"
Ascii strToReplace, SHOST, strLocFis & "clntwin\wmbuild.nsi"
strToReplace = "9898"
Ascii strToReplace, DSPORT, strLocFis & "config\custom.xml"
Ascii strToReplace, DSPORT, strLocFis & "custom.xml"
Ascii strToReplace, DSPORT, strLocFis & "local\db_defaults"
strToReplace = "9899"
Ascii strToReplace, FSPORT, strLocFis & "config\custom.xml"
Ascii strToReplace, FSPORT, strLocFis & "custom.xml"
Ascii strToReplace, FSPORT, strLocFis & "local\db_defaults"
strToReplace ="8899"
Ascii strToReplace, PORT2, strLocFis & "config\custom.xml"
strToReplace = "129.148.70.86"
Ascii strToReplace, ADD1, strLocFis & "jdk\sample\jnlp\corba\war\app\helloworld.jnlp"
strToReplace = "10000"
Ascii strToReplace, PORT1, strLocFis & "jdk\sample\jnlp\corba\war\app\helloworld.jnlp"
strToReplace = "0.0.0.0"
Ascii strToReplace, ADD3, strLocFis & "WebServicesServer\lib\install\templates\developer\domain.xml"
Ascii strToReplace, ADD3, strLocFis & "WebServicesServer\lib\install\templates\enterprise\domain.xml"
Ascii strToReplace, ADD3, strLocFis & "WebServicesServer\lib\install\templates\default-config.xml"
Ascii strToReplace, ADD3, strLocFis & "WebServicesServer\lib\install\templates\ee\default-config.xml"
Ascii strToReplace, ADD3, strLocFis & "WebServicesServer\lib\install\templates\default-domain.xml.template"
Ascii strToReplace, ADD3, strLocFis & "WebServicesServer\lib\install\templates\default-domain.xml.template.darwin"
Ascii strToReplace, ADD3, strLocFis & "WebServicesServer\lib\install\templates\ee\default-domain.xml.template"
Ascii strToReplace, ADD3, strLocFis & "WebServicesServer\lib\install\templates\ee\default-domain.xml.template.darwin"
Ascii strToReplace, ADD3, strLocFis & "WebServicesServer\lib\install\templates\ri-domain.xml.template"
Ascii strToReplace, ADD3, strLocFis & "WebServicesServer\lib\install\templates\ee\ri-domain.xml.template"
Ascii strToReplace, ADD3, strLocFis & "WebServicesServer\lib\install\templates\samples-domain.xml.template"
Ascii strToReplace, ADD3, strLocFis & "WebServicesServer\lib\install\templates\ee\samples-domain.xml.template"
Ascii strToReplace, ADD3, strLocFis & "WebServicesServer\lib\install\templates\temp-domain.xml"
Ascii strToReplace, ADD3, strLocFis & "WebServicesServer\lib\install\templates\ee\temp-domain.xml"
strToReplace = "1.7.1.103"
Ascii strToReplace, ADD4, strLocFis & "jdk\lib\visualvm\platform10\update_tracking\org-jdesktop-layout.xml"
Ascii strToReplace, ADD4, strLocFis & "jdk\lib\visualvm\platform10\config\Modules\domain.xml"
strToReplace = "12.34.56.78"
Ascii strToReplace, ADD5, strLocFis &  "Apache\conf\httpd.conf.in"
Ascii strToReplace, ADD5, strLocFis &  "Apache\conf\httpd.conf.in"
strToReplace = "80"
Ascii strToReplace, SDSPORT, strLocFis & "Apache\conf\httpd.conf.in"
Ascii strToReplace, SDSPORT, strLocFis & "Apache\conf\httpd.conf.in"
Ascii strToReplace, SDSPORT, strLocFis & "clntwin\mmbuild.nsi"
Ascii strToReplace, SDSPORT, strLocFis & "clntwin\sapintegrationbuild.nsi"
Ascii strToReplace, SDSPORT, strLocFis & "clntwin\wmbuild.nsi"
Ascii strToReplace, SDSPORT, strLocFis & "config\custom.xml"
Ascii strToReplace, SDSPORT, strLocFis & "custom.xml"