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
If Request("ctrl") = 1 Then
	'add/edit
	'institution
	If Request("selInst") = 0 Then
		If Request("txtNewInst") <> "" Then
			Set rsInst = Server.CreateObject("ADODB.RecordSet")
			sqlInst = "SELECT * FROM Institution_T WHERE UCase(Facility) = '" & Trim(Ucase(Request("txtNewInst"))) & "'"
			rsInst.Open sqlInst, g_strCONN, 1 , 3
			If rsInst.EOF Then
				rsInst.AddNew
				rsInst("facility") = Request("txtNewInst")
				rsInst("Date") = Date
				rsInst.Update
				Session("MSG") = Session("MSG") & Request("txtNewInst") & " Institution saved.<br>" 
			Else
				Session("MSG") = Session("MSG") & Request("txtNewInst") & " Institution already exists in the database.<br>" 
			End If
			rsInst.Close
			Set rsInst = Nothing
		End If
	Else
		If Request("txtNewInst") <> "" Then
			Set rsInst = Server.CreateObject("ADODB.RecordSet")
			sqlInst = "SELECT * FROM institution_T WHERE index = " & Request("selInst")
			rsInst.Open sqlInst, g_strCONN, 1 , 3
			If NOT rsInst.EOF Then
				rsInst("facility") = Request("txtNewInst")
				rsInst.Update
				Session("MSG") = Session("MSG") & Request("txtNewInst") & " Institution saved.<br>" 
			Else
				Session("MSG") = Session("MSG") & Request("txtNewInst") & " Institution cannot be found in the database.<br>" 
			End If
			rsInst.Close
			Set rsInst = Nothing
		End If
	End If
	'department
	If Request("selDept") = 0 Then
		If Request("selInst") <> 0 Then
			If Request("txtNewDept") <> "" Then
				Set rsDept = Server.CreateObject("ADODB.REcordSet")
				sqlDept = "SELECT * FROM Dept_T WHERE InstID = " & Request("selInst") & " AND Ucase(dept) = '" & Ucase(trim(Request("txtNewDept"))) & "'"
				rsDept.Open sqlDept, g_strCONN, 1, 3
				If rsDept.EOF Then
					rsDept.AddNew
					rsDept("InstID") = Request("selInst")
					rsDept("dept") = Request("txtNewDept")
					rsDept("Class") = Request("selClass")
					rsDept("Address") = CleanMe(Request("txtInstAddr"))
					rsDept("City") = Request("txtInstCity")
					rsDept("State") = Request("txtInstState")
					rsDept("Zip") = Request("txtInstZip")
					rsDept("InstAdrI") = CleanMe(Request("txtInstAddrI"))
					rsDept("Blname") = CleanMe(Request("txtBlname"))
					rsDept("BAddress") = CleanMe(Request("txtBillAddr"))
					rsDept("BCity") = Request("txtBillCity")
					rsDept("BState") = Request("txtBillState")
					rsDept("BZip") = Request("txtBillZip")
					rsDept.Update
					Session("MSG") = Session("MSG") & Request("txtNewDept") & " Department for " & GetInst2(Request("selInst")) & " saved.<br>" 
				Else
					Session("MSG") = Session("MSG") & Request("txtNewDept") & " Department for " & GetInst2(Request("selInst")) & " already exists in the database.<br>"  
				End If
				rsDept.CLose
				Set rsDept = Nothing
			End If
		Else
			Session("MSG") = Session("MSG") & "Select Institution first for New Department.<br>"
		End If
	Else
		If Request("selInst") <> 0 Then
			If Request("txtNewDept") <> "" Then
				Set rsDept = Server.CreateObject("ADODB.RecordSet")
				sqlDept = "SELECT * FROM Dept_T WHERE index = " & Request("selDept")
				rsDept.Open sqlDept, g_strCONN, 1, 3
				If Not rsDept.EOF Then
					rsDept("InstID") = Request("selInst")
					rsDept("dept") = Request("txtNewDept")
					rsDept("Class") = Request("selClass")
					rsDept("Address") = CleanMe(Request("txtInstAddr"))
					rsDept("City") = Request("txtInstCity")
					rsDept("State") = Request("txtInstState")
					rsDept("Zip") = Request("txtInstZip")
					rsDept("InstAdrI") = CleanMe(Request("txtInstAddrI"))
					rsDept("Blname") = CleanMe(Request("txtBlname"))
					rsDept("BAddress") = CleanMe(Request("txtBillAddr"))
					rsDept("BCity") = Request("txtBillCity")
					rsDept("BState") = Request("txtBillState")
					rsDept("BZip") = Request("txtBillZip")
					rsDept.Update
					Session("MSG") = Session("MSG") & Request("txtNewDept") & " Department for " & GetInst2(Request("selInst")) & " saved.<br>" 
				Else
					Session("MSG") = Session("MSG") & Request("txtNewDept") & " Department for " & GetInst2(Request("selInst")) & " cannot be found in the database.<br>" 
				End If
				rsDept.Close
				Set rsDept = Nothing	
			End If
		End If
	End If
	'requesting person
	If Request("selReq") = 0 Then
		If Request("selDept") <> 0 Then
			Set rsReq = Server.CreateObject("ADODB.RecordSet")
			sqlReq = "SELECT * FROM requester_T WHERE UCase(lname) = '" & ucase(trim(Request("txtReqLname"))) & "' AND UCase(fname) = '" & ucase(trim(Request("txtReqfname"))) & "'"
			rsReq.Open sqlReq, g_strCONN,1, 3
			If Not rsReq.EOF Then
				ReqID = rsReq("index")
				Set rsReqDept = Server.CreateObject("ADODB.RecordSet")
				sqlReqDept = "SELECT * FROM reqdept_T WHERE ReqID = " & ReqID & " AND deptID = " & Request("selDept")
				rsReqDept.Open sqlReqDept, g_strCONN, 1, 3
				If rsReqDept.EOF Then
					ExistReq = False
				Else
					ExistReq = True
				End If
				rsReqDept.Close
				Set rsReqDept = Nothing
				If ExistReq Then
					Session("MSG") = Session("MSG") & "Requesting Person for " & GetDept(Request("selDept")) & " already exists in the database.<br>"  	
				Else
					If Request("txtphone") = "" And Request("txtfax") = "" And Request("txtemail") = "" Then
						Session("MSG") = Session("MSG") & "Requesting person should at least have 1 contact information.<br>"
					Else
						rsReq.AddNew
						rsReq("Lname") = CleanMe(Request("txtReqLname"))
						rsReq("Fname") = CleanMe(Request("txtReqFname"))
						rsReq("Phone") = Request("txtphone")
						rsReq("pExt") = Request("txtReqExt")
						rsReq("Fax") = Request("txtfax")
						rsReq("Email") = Request("txtemail")
						myPrime = Request("radioPrim1")
						If IsNull(myPrime) Then myPrime = 2
						rsReq("prime") = myPrime
						rsReq.Update
						tmpReq = rsReq("index")
						Set rsReqDept = Server.CreateObject("ADODB.RecordSet")
						sqlReqDept = "SELECT * FROM reqdept_T"
						rsReqDept.Open sqlReqDept, g_strCONN, 1, 3
						rsReqDept.AddNew
						rsReqDept("ReqID") = tmpReq
						rsReqDept("DeptID") = Request("selDept")
						rsReqDept.Update
						rsReqDept.Close
						Set rsReqDept = Nothing
						Session("MSG") = Session("MSG") & "Requesting person for " & GetDept(Request("selDept")) & " saved.<br>"
					End If
				End If
			Else
				If Request("txtphone") = "" And Request("txtfax") = "" And Request("txtemail") = "" Then
					Session("MSG") = Session("MSG") & "Requesting person should at least have 1 contact information.<br>"
				Else
					rsReq.AddNew
					rsReq("Lname") = CleanMe(Request("txtReqLname"))
					rsReq("Fname") = CleanMe(Request("txtReqFname"))
					rsReq("Phone") = Request("txtphone")
					rsReq("pExt") = Request("txtReqExt")
					rsReq("Fax") = Request("txtfax")
					rsReq("Email") = Request("txtemail")
					myPrime = Request("radioPrim1")
					If IsNull(myPrime) Then myPrime = 2
					rsReq("prime") = myPrime
					rsReq.Update
					tmpReq = rsReq("index")
					Set rsReqDept = Server.CreateObject("ADODB.RecordSet")
					sqlReqDept = "SELECT * FROM reqdept_T"
					rsReqDept.Open sqlReqDept, g_strCONN, 1, 3
					rsReqDept.AddNew
					rsReqDept("ReqID") = tmpReq
					rsReqDept("DeptID") = Request("selDept")
					rsReqDept.Update
					rsReqDept.Close
					Set rsReqDept = Nothing
					Session("MSG") = Session("MSG") & "Requesting person for " & GetDept(Request("selDept")) & " saved.<br>"
				End If
			End If
			rsReq.Close
			Set rsReq = Nothing
		Else
			Session("MSG") = Session("MSG") & "Select Department first for New Requesting Person."
		End If	
	Else
		If Request("selDept") <> 0 Then
			If Request("txtReqLname") = "" And Request("txtReqFname") = "" Then
				Session("MSG") = Session("MSG") & "Requesting Person's name cannot be blank.<br>"
			Else
				Set rsReq = Server.CreateObject("ADODB.RecordSet")
				sqlReq = "SELECT * FROM requester_T WHERE index = " & Request("selReq")
				rsReq.Open sqlReq, g_strCONN, 1, 3
				If Not rsReq.EOF Then
					rsReq("Lname") = CleanMe(Request("txtReqLname"))
					rsReq("Fname") = CleanMe(Request("txtReqFname"))
					rsReq("Phone") = Request("txtphone")
					rsReq("pExt") = Request("txtReqExt")
					rsReq("Fax") = Request("txtfax")
					rsReq("Email") = Request("txtemail")
					myPrime = Request("radioPrim1")
					If IsNull(myPrime) Then myPrime = 2
					rsReq("prime") = myPrime
					rsReq.Update
				End If
				rsReq.Close
				Set rsReq = Nothing
				Set rsReqDept = Server.CreateObject("ADODB.RecordSet")
				sqlReqDept = "SELECT * FROM reqdept_T WHERE ReqID = " & Request("selReq") & " AND DeptID = " & Request("selDept")
				rsReqDept.Open sqlReqDept, g_strCONN, 1, 3
				If rsReqDept.EOF Then
					rsReqDept.AddNew
					rsReqDept("ReqID") = Request("selReq")
					rsReqDept("DeptID") = Request("selDept")
					rsReqDept.Update
				End If
				rsReqDept.Close
				Set rsReqDept = Nothing
				Session("MSG") = Session("MSG") & "Requesting person for " & GetDept(Request("selDept")) & " saved.<br>"
			End If
		Else
			Session("MSG") = Session("MSG") & "Select Department first for New Requesting Person.<br>"
		End If
	End If
	Response.Redirect "admintools.asp"
ElseIf Request("ctrl") = 2 Then
	If Request("chkDelInst") <> "" And Request("selInst") <> 0 Then
		Session("MSG") = Session("MSG") & "Institution " & GetInst2(Request("selInst")) & " deleted.<br>"
		Set rsInst = Server.CreateObject("ADODB.RecordSet")
		sqlInst = "DELETE FROM Institution_T WHERE index = " & Request("selInst")
		rsInst.Open sqlInst, g_strCONN, 1, 3
		Set rsInst = Nothing
		Set rsInst = Server.CreateObject("ADODB.RecordSet")
		sqlInst = "DELETE FROM Dept_T WHERE InstID = " & Request("selInst")
		rsInst.Open sqlInst, g_strCONN, 1, 3
		Set rsInst = Nothing
	Else
		Session("MSG") = Session("MSG") & "Select Institution to Delete.<br>"
	End If
	If Request("chkDelDept") <> "" And Request("selDept") <> 0 Then
		Session("MSG") = Session("MSG") & "Department " & GetDept(Request("selDept")) & " deleted.<br>"
		Set rsInst = Server.CreateObject("ADODB.RecordSet")
		sqlInst = "DELETE FROM Dept_T WHERE index = " & Request("selDept")
		rsInst.Open sqlInst, g_strCONN, 1, 3
		Set rsInst = Nothing
		Set rsInst = Server.CreateObject("ADODB.RecordSet")
		sqlInst = "DELETE FROM ReqDept_T WHERE DeptID = " & Request("selDept")
		rsInst.Open sqlInst, g_strCONN, 1, 3
		Set rsInst = Nothing
	Else
		Session("MSG") = Session("MSG") & "Select Department to Delete.<br>"
	End If
	If Request("chkDelReq") <> "" And Request("selReq") <> 0 Then
		Session("MSG") = Session("MSG") & "Requesting Person " & GetReq(Request("selReq")) & " deleted.<br>"
		Set rsInst = Server.CreateObject("ADODB.RecordSet")
		sqlInst = "DELETE FROM Requester_T WHERE index = " & Request("selReq")
		rsInst.Open sqlInst, g_strCONN, 1, 3
		Set rsInst = Nothing
		Set rsInst = Server.CreateObject("ADODB.RecordSet")
		sqlInst = "DELETE FROM ReqDept_T WHERE ReqID = " & Request("selReq")
		rsInst.Open sqlInst, g_strCONN, 1, 3
		Set rsInst = Nothing
	Else
		Session("MSG") = Session("MSG") & "Select Requesting Person to Delete.<br>"
	End If
	Response.Redirect "admintools.asp"
End If
%>