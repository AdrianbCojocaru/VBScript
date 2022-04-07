' --------------------------------------------------------------------------
'  File:         append_tnsnames.vbs
'  Purpose:  used to add new configuration for tnsnames.ora
'  Date:          11,October,2013
'  Description:
'  Usage: example usage:
'  wscript.exe "append_tnsnames.vbs"
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
	
	WriteOra (ora_p1)
	WriteOra (ora_p2)
	WriteOra (ora_p3)
	WriteOra (ora_p4)
	WriteOra (ora_p5)
	
Function WriteOra (ora_path)	
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
		set ora_file = FSO.OpenTextFile(ora_path, ForReading)
		if not ora_file.AtEndOfStream then
			strContent = ora_file.ReadAll
		end if
		ora_file.close
		strPut = ""
		if instr(strContent,content1) = 0 then
			strPut = strPut & content1
		end if
				
		set ora_file = FSO.OpenTextFile(ora_path, ForAppending)
		ora_file.Write strPut
		ora_File.close
	else
		nr_car = len (ora_path)
		path_ok_nr = nr_car - 13
		if fso.FolderExists (left (ora_path, path_ok_nr)) then	'path - tnsnames.ora
			set ora_file = FSO.CreateTextFile(ora_path)
			strPut =content1
			ora_file.Write strPut
			ora_File.close
		end if
	end if
End Function