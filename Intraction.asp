<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
Function CleanMe(xxx)
	CleanMe = xxx
	If Not IsNull(xxx) Or xxx <> "" Then CleanMe = Replace(xxx, "'", " ")
End Function
If request("action") = 1 Then 'add/edit user
	If Request("selIntr") <> 0 Then
		Set rsIntr = Server.CreateObject("ADODB.RecordSet")
		sqlIntr = "SELECT * FROM interpreter_T WHERE index = " & Request("selIntr")
		rsIntr.Open sqlIntr, g_strCONN, 1, 3
		If Not rsIntr.EOF Then
			tmpIntr =  Request("selIntr")
			If Request("txtIntrLname") <> "" And Request("txtIntrFname") <> "" Then
				rsIntr("First Name") = CleanMe(Request("txtIntrFname"))
				rsIntr("Last Name") = CleanMe(Request("txtIntrLname"))
				rsIntr("E-mail") = Request("txtIntrEmail")
				rsIntr("Phone1") = Request("txtIntrP1")
				rsIntr("P1Ext") = Request("txtIntrExt")
				rsIntr("Phone2") = Request("txtIntrP2")
				rsIntr("Fax") = Request("txtIntrFax")
				rsIntr("Address1") = CleanMe(Request("txtIntrAddr"))
				rsIntr("City") = Request("txtIntrCity")
				rsIntr("State") = Request("txtIntrState")
				rsIntr("Zip Code") = Request("txtIntrZip")
				rsIntr("IntrAdrI") = CleanMe(Request("txtIntrAddrI"))
				'rsIntr("IntrAdrI") = Request("txtHire")
				rsIntr("InHouse") = False
				If Request("chkInHouse") <> "" Then rsIntr("InHouse") = True
				rsIntr("Stat") = Request("radioStatIntr")
				rsIntr("Rate") = Request("selIntrRate")
				'rsIntr("Crime") = False
				'If Request("chkCrim") <> "" Then rsIntr("crime") = True
				'rsIntr("drive") = False
				rsIntr("filenum") = Request("txtfilenum")
				If Request("txtdrivedate") <> "" Then 
					If IsDate(Request("txtdrivedate")) Then
						rsIntr("DriveDate") = Request("txtdrivedate")
					End If
				Else 
					rsIntr("DriveDate") = Empty
				End If
				If Request("txtcrimedate") <> "" Then 
					If IsDate(Request("txtcrimedate")) Then
						rsIntr("CrimeDate") = Request("txtcrimedate")
					End If
				Else 
					rsIntr("CrimeDate") = Empty
				End If
				rsIntr("ssnum") = Request("txtss")
				rsIntr("passnum") = Request("txtpass")
				If Request("txtpassexp") <> "" Then 
					If IsDate(Request("txtpassexp")) Then
						rsIntr("passexp") = Request("txtpassexp")
					End If
				Else 
					rsIntr("passexp") = Empty
				End If
				rsIntr("drivenum") = Request("txtdrive")
				If Request("txtdriveexp") <> "" Then 
					If IsDate(Request("txtdriveexp")) Then
						rsIntr("driveexp") = Request("txtdriveexp")
					End If
				Else 
					rsIntr("driveexp") = Empty
				End If
				rsIntr("greennum") = Request("txtgreen")
				If Request("txtgreenexp") <> "" Then 
					If IsDate(Request("txtgreenexp")) Then
						rsIntr("greenexp") = Request("txtgreenexp")
					End If
				Else 
					rsIntr("greenexp") = Empty
				End If
				rsIntr("employnum") = Request("txtemploy")
				If Request("txtemployexp") <> "" Then 
					If IsDate(Request("txtemployexp")) Then
						rsIntr("employexp") = Request("txtemployexp")
					End If
				Else 
					rsIntr("employexp") = Empty
				End If
				rsIntr("carnum") = Request("txtcar")
				If Request("txtcarexp") <> "" Then 
					If IsDate(Request("txtcarexp")) Then
						rsIntr("carexp") = Request("txtcarexp")
					End If
				Else 
					rsIntr("carexp") = Empty
				End If
				'If Request("chkdriv") <> "" Then rsIntr("drive") = True
				rsIntr("train") = Request("txttrain")
				rsIntr("Active") = False
				
				If Request("txtvacfrom") <> "" Then 
					If IsDate(Request("txtvacfrom")) Then
						rsIntr("vacfrom") = Request("txtvacfrom")
					Else
						rsIntr("vacfrom") = Empty
						rsIntr("vacto") = Empty
					End If
				Else 
					rsIntr("vacfrom") = Empty
				End If
				If Request("txtvacto") <> "" Then 
					If IsDate(Request("txtvacto")) Then
						rsIntr("vacto") = Request("txtvacto")
					Else
						rsIntr("vacfrom") = Empty
						rsIntr("vacto") = Empty
					End If
				Else 
					rsIntr("vacto") = Empty
				End If
				
				rsIntr("datehired") = Empty
				If Request("txthire") <> "" Then
	    	If isDate(Request("txthire")) THen rsIntr("datehired") = Request("txthire")
	    End If
	    rsIntr("dateterm") = Empty
	    If Request("txtterm") <> "" Then
	    	If isDate(Request("txtterm")) THen rsIntr("dateterm") = Request("txtterm")
	    End If
				If Request("radioStatIntr1") = 0 Then rsIntr("Active") = True
				rsIntr("Comments") = Request("txtIntrCom")	
				If Request("SelIntrLang") <> "0" Then 'SAVE LANGUAGES OF INTERPRETER
					If rsIntr("Language1") = "" Or IsNull(rsIntr("Language1")) Then 
						rsIntr("Language1") = Request("SelIntrLang")
					Else
						If rsIntr("Language2") = ""  Or IsNull(rsIntr("Language2")) Then
							rsIntr("Language2") = Request("SelIntrLang")
						Else
							If rsIntr("Language3") = ""  Or IsNull(rsIntr("Language3")) Then
								rsIntr("Language3") = Request("SelIntrLang")
							Else
								If rsIntr("Language4") = "" Or IsNull(rsIntr("Language4")) Then
									rsIntr("Language4") = Request("SelIntrLang")
								Else
									If rsIntr("Language5") = "" Or IsNull(rsIntr("Language5")) Then rsIntr("Language5") = Request("SelIntrLang")
								End If
							End If
						End If 	
					End If
				End If
				'DELETE LANGUAGES OF INTERPRETER
				If Request("chkLang1") <> "" Then  rsIntr("Language1") = ""
				If Request("chkLang2") <> "" Then  rsIntr("Language2") = ""
				If Request("chkLang3") <> "" Then  rsIntr("Language3") = ""
				If Request("chkLang4") <> "" Then  rsIntr("Language4") = ""
				If Request("chkLang5") <> "" Then  rsIntr("Language5") = ""
				'CREATE LOG
			on error resume next
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
				strLog = Now & vbTab & "Interpreter (ID: " & Request("selIntr") & ") was edited by " & Session("UsrName") & "."
				LogMe.WriteLine strLog
				Set LogMe = Nothing
				Set fso = Nothing
				rsIntr.Update
			Else
				Session("MSG") = Session("MSG") & "<br>Error: Interpreter's name cannot be blank."
			End If
		End If
		rsIntr.Close
		Set rsIntr = Nothing
		response.redirect "adminIntr.asp?IntrID=" & Request("selIntr") & "&type=" & Request("tmpType")
	Else
		Set rsIntr = Server.CreateObject("ADODB.RecordSet")
		sqlIntr = "SELECT * FROM interpreter_T"
		rsIntr.Open sqlIntr, g_strCONN, 1, 3
		'If Not rsIntr.EOF Then
			
			If Request("txtIntrLname") <> "" And Request("txtIntrFname") <> "" Then
				rsIntr.Addnew
				newID = rsIntr("index")
				rsIntr("First Name") = CleanMe(Request("txtIntrFname"))
				rsIntr("Last Name") = CleanMe(Request("txtIntrLname"))
				rsIntr("E-mail") = Request("txtIntrEmail")
				rsIntr("Phone1") = Request("txtIntrP1")
				rsIntr("P1Ext") = Request("txtIntrExt")
				rsIntr("Phone2") = Request("txtIntrP2")
				rsIntr("Fax") = Request("txtIntrFax")
				rsIntr("Address1") = CleanMe(Request("txtIntrAddr"))
				rsIntr("City") = Request("txtIntrCity")
				rsIntr("State") = Request("txtIntrState")
				rsIntr("Zip Code") = Request("txtIntrZip")
				rsIntr("IntrAdrI") = CleanMe(Request("txtIntrAddrI"))
				'rsIntr("IntrAdrI") = Request("txtHire")
				rsIntr("InHouse") = False
				If Request("chkInHouse") <> "" Then rsIntr("InHouse") = True
				rsIntr("Stat") = Request("radioStatIntr")
				rsIntr("Rate") = Request("selIntrRate")
				'rsIntr("Crime") = False
				'If Request("chkCrim") <> "" Then rsIntr("crime") = True
				'rsIntr("drive") = False
				rsIntr("filenum") = Request("txtfilenum")
				If Request("txtdrivedate") <> "" Then 
					If IsDate(Request("txtdrivedate")) Then
						rsIntr("DriveDate") = Request("txtdrivedate")
					End If
				Else 
					rsIntr("DriveDate") = Empty
				End If
				If Request("txtcrimedate") <> "" Then 
					If IsDate(Request("txtcrimedate")) Then
						rsIntr("CrimeDate") = Request("txtcrimedate")
					End If
				Else 
					rsIntr("CrimeDate") = Empty
				End If
				rsIntr("ssnum") = Request("txtss")
				rsIntr("passnum") = Request("txtpass")
				If Request("txtpassexp") <> "" Then 
					If IsDate(Request("txtpassexp")) Then
						rsIntr("passexp") = Request("txtpassexp")
					End If
				Else 
					rsIntr("passexp") = Empty
				End If
				rsIntr("drivenum") = Request("txtdrive")
				If Request("txtdriveexp") <> "" Then 
					If IsDate(Request("txtdriveexp")) Then
						rsIntr("driveexp") = Request("txtdriveexp")
					End If
				Else 
					rsIntr("driveexp") = Empty
				End If
				rsIntr("greennum") = Request("txtgreen")
				If Request("txtgreenexp") <> "" Then 
					If IsDate(Request("txtgreenexp")) Then
						rsIntr("greenexp") = Request("txtgreenexp")
					End If
				Else 
					rsIntr("greenexp") = Empty
				End If
				rsIntr("employnum") = Request("txtemploy")
				If Request("txtemployexp") <> "" Then 
					If IsDate(Request("txtemployexp")) Then
						rsIntr("employexp") = Request("txtemployexp")
					End If
				Else 
					rsIntr("employexp") = Empty
				End If
				rsIntr("carnum") = Request("txtcar")
				If Request("txtcarexp") <> "" Then 
					If IsDate(Request("txtcarexp")) Then
						rsIntr("carexp") = Request("txtcarexp")
					End If
				Else 
					rsIntr("carexp") = Empty
				End If
				'If Request("chkdriv") <> "" Then rsIntr("drive") = True
				rsIntr("train") = Request("txttrain")
				rsIntr("Active") = False
				rsIntr("datehired") = Empty
				If Request("txthire") <> "" Then
	    	If isDate(Request("txthire")) THen rsIntr("datehired") = Request("txthire")
	    End If
	    rsIntr("dateterm") = Empty
	    If Request("txtterm") <> "" Then
	    	If isDate(Request("txtterm")) THen rsIntr("dateterm") = Request("txtterm")
	    End If
				If Request("radioStatIntr1") = 0 Then rsIntr("Active") = True
				rsIntr("Comments") = Request("txtIntrCom")	
				If Request("SelIntrLang") <> "0" Then 'SAVE LANGUAGES OF INTERPRETER
					If rsIntr("Language1") = "" Or IsNull(rsIntr("Language1")) Then 
						rsIntr("Language1") = Request("SelIntrLang")
					Else
						If rsIntr("Language2") = ""  Or IsNull(rsIntr("Language2")) Then
							rsIntr("Language2") = Request("SelIntrLang")
						Else
							If rsIntr("Language3") = ""  Or IsNull(rsIntr("Language3")) Then
								rsIntr("Language3") = Request("SelIntrLang")
							Else
								If rsIntr("Language4") = "" Or IsNull(rsIntr("Language4")) Then
									rsIntr("Language4") = Request("SelIntrLang")
								Else
									If rsIntr("Language5") = "" Or IsNull(rsIntr("Language5")) Then rsIntr("Language5") = Request("SelIntrLang")
								End If
							End If
						End If 	
					End If
				End If
				'DELETE LANGUAGES OF INTERPRETER
				If Request("chkLang1") <> "" Then  rsIntr("Language1") = ""
				If Request("chkLang2") <> "" Then  rsIntr("Language2") = ""
				If Request("chkLang3") <> "" Then  rsIntr("Language3") = ""
				If Request("chkLang4") <> "" Then  rsIntr("Language4") = ""
				If Request("chkLang5") <> "" Then  rsIntr("Language5") = ""
				'CREATE LOG
			on error resume next
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
				strLog = Now & vbTab & "Interpreter (ID: " & Request("selIntr") & ") was edited by " & Session("UsrName") & "."
				LogMe.WriteLine strLog
				Set LogMe = Nothing
				Set fso = Nothing
				rsIntr.Update
			Else
				Session("MSG") = Session("MSG") & "<br>Error: Interpreter's name cannot be blank."
			End If
		'End If
		rsIntr.Close
		Set rsIntr = Nothing
		response.redirect "adminIntr.asp?IntrID= " & newID & "&type=" & Request("tmpType")
	End If
ElseIf request("action") = 2 Then 'delete intr
	set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "DELETE interpreter_T WHERE index = " & Request("selIntr")
	rsIntr.Open sqlIntr, g_strCONN, 3, 1
	Set rsIntr = Nothing
	set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "DELETE InterpreterEval_T WHERE intrID = " & Request("selIntr")
	rsIntr.Open sqlIntr, g_strCONN, 3, 1
	Set rsIntr = Nothing
	set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "DELETE IntrTraining_T WHERE intrID = " & Request("selIntr")
	rsIntr.Open sqlIntr, g_strCONN, 3, 1
	Set rsIntr = Nothing
	Session("MSG") = "Interpreter deleted."
	Response.Redirect "adminIntr.asp?type=" & Request("tmpType")
ElseIf request("action") = 3 Then 'eval add
	Set tblSite = Server.CreateObject("ADODB.RecordSet")
	sqlSite = "SELECT * FROM InterpreterEval_T WHERE IntrID = " & Request("IntrID")
	tblSite.Open sqlSite, g_strCONN, 1, 3
	If Not tblSite.EOF Then
		If Request("ctr") <> "" Then 
			ctr = Request("ctr")
			For i = 0 to ctr 
				tmpctr = Request("chkeval" & i)
				If tmpctr <> "" Then
					strTmp = "index= " & tmpctr 
					tblSite.Find(strTmp)
					If Not tblSite.EOF Then
						tblSite("date") = Request("txtdate" & i)
						tblSite("comment") = Request("txtcom" & i)
						tblSite.Update
					End If
				End If
				tblSite.Movefirst
			Next
		End If 
	End If
	tblSite.Close
	Set tblSite = Nothing
	If Request("txtdate") <> "" And Request("txtcom") <> "" Then
		If IsDate(Request("txtdate")) Then 
			Set rsLog = Server.CreateObject("ADODB.RecordSet")
			sqlLog = "SELECT * FROM InterpreterEval_T"' WHERE IntrID = " & Request("IntrID")
			rsLog.Open sqlLog, g_strCONN, 1, 3
			rsLog.Addnew
			rsLog("IntrID") = Request("IntrID")
			rsLog("date") = Request("txtdate")
			rsLog("comment") = Request("txtcom")
			rsLog.Update
			rsLog.Close
			Set rsLog = Nothing
			Session("MSG") = "Evaluation/Feedback SAVED."
		Else
			Session("MSG") = "Invalid date for Evaluation/Feedback."
		End If
	End If
	Response.redirect "intreval.asp?intrID=" & Request("IntrID") & "&type=" & Request("tmpType")
ElseIf request("action") = 4 Then 'eval delete
	Set tblSite = Server.CreateObject("ADODB.RecordSet")
	sqlSite = "SELECT * FROM InterpreterEval_T WHERE IntrID = " & Request("IntrID")
	tblSite.Open sqlSite, g_strCONN, 1, 3
	If Not tblSite.EOF Then
		If Request("ctr") <> "" Then 
			ctr = Request("ctr")
			For i = 0 to ctr 
				tmpctr = Request("chkeval" & i)
				If tmpctr <> "" Then
					strTmp = "index= " & tmpctr 
					tblSite.Movefirst
					tblSite.Find(strTmp)
					If Not tblSite.EOF Then
						tblSite.Delete
						tblSite.Update
					End If
				End If
			Next
		End If 
	End If
	tblSite.Close
	Set tblSite = Nothing
	Response.redirect "intreval.asp?intrID=" & Request("IntrID") & "&type=" & Request("tmpType")
ElseIf Request("action")= 5 Then 'add trainig
	Set tblSite = Server.CreateObject("ADODB.RecordSet")
	sqlSite = "SELECT * FROM IntrTraining_T WHERE IntrID = " & Request("IntrID")
	tblSite.Open sqlSite, g_strCONN, 1, 3
	If Not tblSite.EOF Then
		If Request("ctr") <> "" Then 
			ctr = Request("ctr")
			For i = 0 to ctr 
				tmpctr = Request("chkeval" & i)
				If tmpctr <> "" Then
					strTmp = "index= " & tmpctr 
					tblSite.Find(strTmp)
					If Not tblSite.EOF Then
						tblSite("Date") = Request("txtdate" & i)
						tblSite("hours") = Request("txthrs" & i)
						tblSite("type") = Request("selTrain" & i)
						If Request("hidcert" & i) = 3 Then
							tblSite("type") = 3
							tblSite("cert") = Request("txtCert" & i)
						End If
						tblSite.Update
					End If
				End If
				tblSite.Movefirst
			Next
		End If 
	End If
	tblSite.Close
	Set tblSite = Nothing
	If Request("txtdate") <> "" And Request("txthrs") <> "" And Request("SelTrain") <> 0 Then
		If IsDate(Request("txtdate")) Then 
			Set rsLog = Server.CreateObject("ADODB.RecordSet")
			sqlLog = "SELECT * FROM IntrTraining_T"' WHERE IntrID = " & Request("IntrID")
			rsLog.Open sqlLog, g_strCONN, 1, 3
			rsLog.Addnew
			rsLog("IntrID") = Request("IntrID")
			rsLog("date") = Request("txtdate")
			rsLog("hours") = Request("txthrs")
			rsLog("type") = Request("selTrain")
			rsLog("cert") = Request("txtCert")
			rsLog.Update
			rsLog.Close
			Set rsLog = Nothing
			Session("MSG") = "Training SAVED."
		Else
			Session("MSG") = "Invalid date for Training."
		End If
	End If
	Response.redirect "intrtrain.asp?intrID=" & Request("IntrID") & "&type=" & Request("tmpType")
ElseIf Request("action") = 6 Then
	Set tblSite = Server.CreateObject("ADODB.RecordSet")
	sqlSite = "SELECT * FROM IntrTraining_T WHERE IntrID = " & Request("IntrID")
	tblSite.Open sqlSite, g_strCONN, 1, 3
	If Not tblSite.EOF Then
		If Request("ctr") <> "" Then 
			ctr = Request("ctr")
			For i = 0 to ctr 
				tmpctr = Request("chkeval" & i)
				If tmpctr <> "" Then
					strTmp = "index= " & tmpctr 
					tblSite.Movefirst
					tblSite.Find(strTmp)
					If Not tblSite.EOF Then
						tblSite.Delete
						tblSite.Update
					End If
				End If
			Next
		End If 
	End If
	tblSite.Close
	Set tblSite = Nothing
	Response.redirect "intrtrain.asp?intrID=" & Request("IntrID") & "&type=" & Request("tmpType")
End If
%>