' --------------------------------------------------------------------------
'  File:         append_tnsnames.vbs
'  Purpose:  used to delete tnsnames.ora configuration
'  Date:          11,October,2013
'  Description:
'  Usage: example usage:
'  wscript.exe "delete_tnsnames.vbs"
' ----------------------------------------------------------------------------

	PropertyLine = session.property("CustomActionData")

	ArrayPropertyLine = Split (PropertyLine, ";")
	
	SPORT = ArrayPropertyLine(0)
	ora_p = ArrayPropertyLine(1)
	VHOST1 = ArrayPropertyLine(2)
	VHOST2 = ArrayPropertyLine(3)
	TDIR = ArrayPropertyLine(4)
	
	ora_p1 = ora_p & "\network\admin\tnsnames.ora"
	ora_p2= TDIR & "Oracle\Ora92\network\ADMIN\tnsnames.ora"
	ora_p3= TDIR & "Oracle\Ora10g\NETWORK\ADMIN\tnsnames.ora"
	ora_p4= TDIR & "Oracle\Ora11g\client\network\admin\tnsnames.ora"
	ora_p5= TDIR & "Oracle\Ora11.202\Client\network\admin\tnsnames.ora"
	
	delora (ora_p1)
	delora (ora_p2)
	delora (ora_p3)
	delora (ora_p4)
	delora (ora_p5)

Function delora (ora_path)
	content1 = content1 & vbCrLf
	content1 = content1 & "TCSP="& vbCrLf
	content1 = content1 & " (DESCRIPTION="& vbCrLf  
	content1 = content1 & "   (ADDRESS_LIST="& vbCrLf
	content1 = content1 & "      (ADDRESS="& vbCrLf
	content1 = content1 & "        (PROTOCOL=TCP)"& vbCrLf
	content1 = content1 & "          (HOST="& VHOST1 &")"& vbCrLf
	content1 = content1 & "          (PORT="& SPORT &")"& vbCrLf
	content1 = content1 & "      )"& vbCrLf
	content1 = content1 & "      (ADDRESS="& vbCrLf
	content1 = content1 & "        (PROTOCOL=TCP)"& vbCrLf
	content1 = content1 & "          (HOST="& VHOST2 &")" & vbCrLf
	content1 = content1 & "          (PORT="& SPORT &")" & vbCrLf
	content1 = content1 & "      )"& vbCrLf
	content1 = content1 & "    )"& vbCrLf
	content1 = content1 & "    (CONNECT_DATA="& vbCrLf
	content1 = content1 & "      (SERVICE_NAME=TCSP)"& vbCrLf
	content1 = content1 & "    )"& vbCrLf
	content1 = content1 & "  )"& vbCrLf
	content1 = content1 & vbCrLf
	
	set FSO = CreateObject("Scripting.FileSystemOBject")
	CONST ForReading  = 1
    CONST ForWritting = 2
    CONST ForAppending = 8
	if FSO.FileExists(ora_path) then
		strContent = ""
		set ora_file = FSO.OpenTextFile(ora_path, ForReading)
		if not ora_file.AtEndOfStream then
			strContent = ora_file.ReadAll
		end if
		ora_file.close
		strContent = replace(strContent,content1,"",1,-1,1)
			
		if strContent= "" then
			FSO.DeleteFile(ora_path)
		else
			set ora_file = FSO.OpenTextFile(ora_path, ForWritting)
			ora_file.Write strContent
			ora_File.close
		end if
	end if
End Function