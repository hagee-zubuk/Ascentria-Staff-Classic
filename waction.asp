<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
Function CleanMe(xxx)
	' clean string
	CleanMe = xxx
	If Not IsNull(xxx) Or xxx <> "" Then CleanMe = Replace(xxx, ",", " ")
End Function
Function GetInstAdr(zzz)
	GetInstAdr = "N/A"
	If zzz = "" Then Exit Function
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT * FROM dept_t WHERE [index] = " & zzz
	rsInst.Open sqlInst, g_strCONN, 3, 1
	If Not rsInst.EOF Then
		GetInstAdr = rsInst("Address") & " | "& rsInst("City") & " | " & rsInst("State") & " | " & rsInst("Zip")
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
If Request("ctrl") = 1 Then
	If Request("txtNewInst") = "" Then
		myInst = Request("selInst")
	Else
		'CHECK INSTITUITION
		If Request("txtNewInst") <> "" Then
			Set rsRP = Server.CreateObject("ADODB.RecordSet")
			sqlRP = "SELECT * FROM institution_T WHERE facility = '" & Request("txtNewInst") & "' "
			rsRP.Open sqlRP, g_strCONN, 3, 1
			If NOT rsRP.EOF Then
				Session("MSG") = "ERROR: Institution already exists."	
			End If
			rsRP.Close
			Set rsRP =Nothing
			If Session("MSG") = "" Then
				Set rsInst = Server.CreateObject("ADODB.RecordSet")
				sqlInst = "SELECT * FROM institution_T"
				rsInst.Open sqlInst, g_strCONN, 1, 3
				rsInst.AddNew
				rsInst("Facility") = Request("txtNewInst")
				rsInst("Date") = Date
				rsInst.Update
				myInst = rsInst("Index")
				rsInst.Close
				Set rsInst = Nothing
			Else
				Response.Redirect "wMain1.asp"
			End If
		End If 
	End If
	Set rsW1 = Server.CreateObject("ADODB.RecordSet")
	sqlW1 = "SELECT * FROM Wrequest_T"
	rsW1.Open sqlW1, g_strCONNW, 1, 3
	rsW1.AddNew
	rsW1("InstID") = myInst
	rsW1("Emergency") = False
	If Request("chkEmer") <> "" Then rsW1("Emergency") = True
	rsW1("EmerFee") = False
	If Request("chkEmerFee") <> "" Then rsW1("EmerFee") = True
	'rsW1("InstRate") = Request("selInstRate")
	rsW1("AppNum") = Request("selNum")
	rsW1.Update
	myID = rsW1("index")
	rsW1.Close
	Set rsW1 = Nothing
	Response.Redirect "wMain2.asp?tmpID=" & myID
ElseIf Request("ctrl") = 2 Then
	If Request("txtInstDept") = "" Then
		myDept = Request("selDept")
	Else
		'CHECK DEPARTMENT
		If Request("txtInstDept") <> "" Then
			Set rsRP = Server.CreateObject("ADODB.RecordSet")
			sqlRP = "SELECT * FROM dept_T WHERE dept = '" & Request("txtInstDept") & "' AND InstID = " & Request("tmpInst")
			rsRP.Open sqlRP, g_strCONN, 3, 1
			If NOT rsRP.EOF Then
				Session("MSG") = Session("MSG") & "ERROR: Department already exists for this insitution."	
			End If
			rsRP.Close
			Set rsRP =Nothing
			If Session("MSG") = "" Then
				Set rsDept = Server.CreateObject("ADODB.RecordSet")
				sqlDept = "SELECT * FROM dept_T"
				rsDept.Open sqlDept, g_strCONN, 1, 3
				rsDept.AddNew
				rsDept("dept") = Request("txtInstDept")
				rsDept("InstID") = Request("tmpInst")
				rsDept("Address") = CleanMe(Request("txtInstAddr"))
				rsDept("City") = Request("txtInstCity")
				rsDept("State") = Request("txtInstState")
				rsDept("Zip") = Request("txtInstZip")
				rsDept("Class") = Request("selClass")
				rsDept("Blname") = Request("txtBlname")
				rsDept("InstAdrI") = Request("txtInstAddrI")
				If Request("chkBill") <> "" Then
					rsDept("BAddress") = CleanMe(Request("txtBilAddr"))
					rsDept("BCity") = Request("txtBilCity")
					rsDept("BState") = Request("txtBilState")
					rsDept("BZip") = Request("txtBilZip")
				Else
					rsDept("BAddress") = CleanMe(Request("txtInstAddr"))
					rsDept("BCity") = Request("txtInstCity")
					rsDept("BState") = Request("txtInstState")
					rsDept("BZip") = Request("txtInstZip")
				End If
				rsDept.Update
				myDept = rsDept("index")
				rsDept.Close
				Set rsDept = Nothing	
			Else
				Response.Redirect "wMain2.asp?tmpID= " & Request("tmpID")
			End If
		End If 
	End If
	Set rsW1 = Server.CreateObject("ADODB.RecordSet")
	sqlW1 = "SELECT * FROM Wrequest_T WHERE [index] = " & Request("tmpID")
	rsW1.Open sqlW1, g_strCONNW, 1, 3
	If Not rsW1.EOF Then
		rsW1("DeptID") = myDept
		rsW1("InstRate") = Z_Cdbl(Request("selInstRate"))
		rsW1("training") = False
		If Request("chkTrain") = 1 Then rsW1("training") = True
		rsW1.Update
	End If
	rsW1.Close
	Set rsW1 = Nothing
	Response.Redirect "wMain3.asp?tmpID=" & Request("tmpID")
ElseIf Request("ctrl") = 3 Then
	If Request("txtReqLname") = "" And Request("txtReqFname") = "" Then
		myReq = Request("selReq")
	Else
		Set rsReq = Server.CreateObject("ADODB.RecordSet")
		sqlReq = "SELECT * FROM requester_T"
		rsReq.Open sqlReq, g_strCONN, 1, 3
		rsReq.AddNew
		rsReq("Lname") = CleanMe(Request("txtReqLname"))
		rsReq("Fname") = CleanMe(Request("txtReqFname"))
		rsReq("phone") = Request("txtphone")
		rsReq("pExt") = Request("txtReqExt")
		rsReq("eMail") = Request("txtemail")
		rsReq("fax") = Request("txtfax")
		myPrime = Request("radioPrim1")
		If IsNull(myPrime) Then myPrime = 2
		rsReq("prime") = myPrime
		rsReq.Update
		myReq = rsReq("Index")
		rsReq.Close
		Set rsReq = Nothing
	End If
	Set rsW1 = Server.CreateObject("ADODB.RecordSet")
	sqlW1 = "SELECT * FROM Wrequest_T WHERE [index] = " & Request("tmpID")
	rsW1.Open sqlW1, g_strCONNW, 1, 3
	If Not rsW1.EOF Then
		rsW1("ReqID") = myReq
		rsW1.Update
	End If
	rsW1.Close
	Set rsW1 = Nothing
	'SAVE REQUESTER TO DEPARTMENT RELATIONSHIP
	Set rsReqDept = Server.CreateObject("ADODB.RecordSet")
	sqlReqDept = "SELECT * FROM reqdept_T WHERE ReqID = " & myReq & " AND DeptID = " & Request("tmpDep")
	rsReqDept.Open sqlReqDept, g_strCONN, 1, 3
	If rsReqDept.EOF Then
		rsReqDept.AddNew
		rsReqDept("ReqID") = myReq
		rsReqDept("DeptID") = Request("tmpDep")
		rsReqDept.Update
	End If
	rsReqDept.Close
	Set rsReqDept = Nothing
	Response.Redirect "wMain4.asp?tmpID=" & Request("tmpID")
ElseIf Request("ctrl") = 4 Then
	Response.Cookies("LBREQUESTW4") = Z_DoEncrypt(Request("tmpID") & "|" & Request("txtClilname") & "|" & Request("txtClifname") & _
		"|" & Request("chkClient") & "|" & Request("txtCliAddrI") & "|" & Request("txtCliAdd") & "|" & Request("chkClientAdd") & _
		"|" & Request("txtCliFon") & "|" & Request("txtCliCity") & "|" & Request("txtCliState") & "|" & Request("txtCliZip") & _
		"|" & Request("txtAlter") & "|" & Request("txtCliDir") & "|" & Request("txtCliCir") & "|" & Request("txtDOB") & _
		"|" & Request("selLang") & "|" & Request("txtAppDate") & "|" & Request("txtAppTFrom") & "|" & Request("txtAppTTo") & _
		"|" & Request("txtAppLoc") & "|" & Request("txtDocNum") & "|" & Request("txtCrtNum") & "|" & Request("txtcom")& _
		"|" & Request("selGender") & "|" & Request("chkMinor") & "|" & Request("chkout") & "|" & Request("chkmed") & "|" & Request("MCnum")& _
		"|" & Request("txtJudge") & "|" & Request("txtClaim") & "|" & Request("chkcall") & "|" & Request("chkleave"))
	If Request("txtDOB") <> "" Then
		If Not IsDate(Request("txtDOB")) Then Session("MSG") = Session("MSG") & "ERROR: Invalid Date of Birth."
	End If
	If Request("txtAppdate") <> "" Then
		If Not IsDate(Request("txtAppdate")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Appointment date."
	End If
	If Request("txtAppTFrom") <> "" Then
		If Not IsDate(Request("txtAppTFrom")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Appointment Time (From:)."
	End If
	If Request("txtAppTTo") <> "" Then
		If Not IsDate(Request("txtAppTTo")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Appointment Time (To:)."
	End If
	' check number of appointments
	Set rsMain = Server.CreateObject("ADODB.RecordSet")
	sqlMain = "SELECT * FROM wrequest_T WHERE [index] = " & Request("tmpID")
	rsMain.Open sqlMain, g_strCONNW, 1, 3
	If Not rsMain.EOF Then
		myAppNum = rsMain("AppNum")
	End If
	rsMain.Close
	Set rsMain = Nothing
	If myAppNum > 1 Then
		ctrApp = 2
		Do Until ctrApp = myAppNum + 1
			If Request("txtAppdate" & ctrApp) <> "" Then
				If Not IsDate(Request("txtAppdate" & ctrApp)) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Appointment date (" & ctrApp & ")."
			End If
			If Request("txtAppTFrom" & ctrApp) <> "" Then
				If Not IsDate(Request("txtAppTFrom" & ctrApp)) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Appointment Time (From:)(" & ctrApp & ")."
			End If
			If Request("txtAppTTo" & ctrApp) <> "" Then
				If Not IsDate(Request("txtAppTTo" & ctrApp)) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Appointment Time (To:)(" & ctrApp & ")."
			End If
			ctrApp = ctrApp + 1
		Loop
	End If
	'
	If Session("MSG") = "" Then
		Set rsMain = Server.CreateObject("ADODB.RecordSet")
		sqlMain = "SELECT * FROM wrequest_T WHERE [index] = " & Request("tmpID")
		rsMain.Open sqlMain, g_strCONNW, 1, 3
		If Not rsMain.EOF Then
			rsMain("clname") = CleanMe(Request("txtClilname"))
			rsMain("cfname") = CleanMe(Request("txtClifname"))
			rsMain("Client") = False
			If Request("chkClient") <> "" Then rsMain("Client") = True
			rsMain("Caddress") = CleanMe(Request("txtCliAdd"))
			rsMain("Ccity") = Request("txtCliCity")
			rsMain("Cstate") = Ucase(Request("txtCliState"))
			rsMain("Czip") = Request("txtCliZip")
			rsMain("directions") = Request("txtCliDir")
			rsMain("spec_cir") = Request("txtCliCir")
			rsMain("DOB") = Z_DateNull(Request("txtDOB"))
			rsMain("LangID") = Request("selLang")
			rsMain("appDate") = Z_DateNull(Request("txtAppDate"))
			rsMain("appTimeFrom") = Z_DateNull(Request("txtAppDate") & " " & Request("txtAppTFrom"))
			rsMain("appTimeTo") = Z_DateNull(Request("txtAppDate") & " " & Request("txtAppTTo"))
			rsMain("appLoc") = Request("txtAppLoc")
			rsMain("docNum") = Request("txtDocNum")
			rsMain("CrtRumNum") = Request("txtCrtNum")
			rsMain("Comment") = Request("txtcom")
			rsMain("Cphone") = Request("txtCliFon")
			rsMain("CAphone") = Request("txtAlter")
			rsMain("CliAdd") = False
			If Request("chkClientAdd") <> "" Then rsMain("CliAdd") = True
			rsMain("CliAdrI") = Request("txtCliAddrI")
			rsMain("Comment") = Request("txtcom")
			rsMain("Gender") = Request("selGender")
			rsMain("Child") = False
			If Request("chkMinor") <> "" Then rsMain("Child") = True
			rsMain("outpatient") = False
			If Request("chkout") <> "" Then rsMain("outpatient") = True
			rsMain("hasmed") = False
			If Request("chkmed") <> "" Then rsMain("hasmed") = True
			rsMain("medicaid") = Request("MCnum")
			rsMain("meridian") = Request("MHPnum")
			rsMain("nhhealth") = Request("NHHFnum")
			rsMain("wellsense") = Request("WSHPnum")
			rsMain("acknowledge") = false
			If Request("chkawk") <> "" Then rsMain("acknowledge") = True
			rsMain("autoacc") = False
			If Request("chkacc") <> "" Then rsMain("autoacc") = True
			rsMain("wcomp") = False
			If Request("chkcomp") <> "" Then rsMain("wcomp") = True
			rsMain("secins") = Request("selIns")
			rsMain("PDamount") = Z_CZero(Request("txtPDamount"))
			rsMain("UploadFile") = False
			rsMain("ApprovePDF") = False
			If FileUpload(Request("h_tmpfilename")) Then 
				rsMain("UploadFile") = True
				rsMain("filename") = Request("h_tmpfilename") & ".PDF"
			End If
			rsMain("mrrec") = Request("mrrec")
			rsMain("blockSched") = false
			If Request("chkblock") <> "" Then rsMain("blockSched") = True
			rsMain("Judge") = Request("txtJudge")
			rsMain("Claimant") = Request("txtclaim")
			rsMain("courtcall") = false
			If Request("chkcall") <> "" Then rsMain("courtcall") = True
			rsMain("leavemsg") = false
			If Request("chkleave") <> "" Then rsMain("leavemsg") = True
			rsMain.Update
			myDept = rsMain("DeptID")
			myAppNum = rsMain("AppNum")
			myLang = rsMain("LangID")
			myInst = rsMain("InstID")
		End If
		rsMain.Close
		Set rsMain = Nothing
	Else
		Response.Redirect "wMain4.asp?tmpID= " & Request("tmpID")
	End If
	
	'SAVE TO LB DB
	TimeNow = Now

	Set rsWiz = Server.CreateObject("ADODB.RecordSet")
	
	sqlWiz = "SELECT * FROM wrequest_T WHERE [index] = " & Request("tmpID")

	
	rsWiz.Open sqlWiz, g_strCONNW, 3, 1

	If Not rsWiz.EOF Then
		ctrApp = 1
		Do Until ctrApp = myAppNum + 1
			Set rsLB = Server.CreateObject("ADODB.RecordSet")
			sqlLB = "SELECT * FROM request_T WHERE timestamp = '" & now & "'"
			rsLB.Open sqlLB, g_strCONN, 1, 3
			rsLB.AddNew
			x = 1
			'On error resume next
			'response.write "xxx: " & rsLB.Fields.Count & "<br>"
	    Do Until x = rsLB.Fields.Count
	    	If ctrApp > 1 Then 
	    		If x = 3 Or x = 4 Or x = 5 Then 'Date
	    			If x = 3 Then rsLB.Fields(3).value = Request("txtAppDate" & ctrApp)
	    			If x = 4 Then rsLB.Fields(4).value = Request("txtAppDate" & ctrApp)  & " " & Request("txtAppTFrom" & ctrApp)
	    			If x = 5 Then rsLB.Fields(5).value = Request("txtAppDate" & ctrApp)  & " " & Request("txtAppTTo" & ctrApp)
	    		Else
	        	rsLB.Fields(x).Value = Z_NullFix(rsWiz.Fields(x).Value)
	        End If
	      Else
	      	'response.write "x: " & x & "<br>"
	    		rsLB.Fields(x).Value = Z_NullFix(rsWiz.Fields(x).Value)
	    	End If 	
	        x = x + 1
	    Loop
	    rsLB("timestamp") = TimeNow 
	    rsLB.Update
	  	myID = rsLB("index")
	  	rsLB.Close
	  	Set rsLB = Nothing
	  	'LOG UPLOAD
	  	If FileUpload(Request("h_tmpfilename")) Then  'save Form on DB
				Set rsFile = Server.CreateObject("ADODB.RecordSet")
				sqlFile = "SELECT * FROM pdf_T"
				rsFile.Open sqlFile, g_strCONN, 1, 3
				rsFile.AddNew
				rsFile("appID") = myID
				rsFile("filename") = Request("h_tmpfilename") & ".PDF"
				rsFile("datestamp") = Now
				rsFile.Update
				rsFile.Close
				Set rsFile = Nothing
			End If
	  	'SAVE HISTORY
	
			Set rsHist = Server.CreateObject("ADODB.RecordSet")
			sqlHist = "SELECT * FROM History_T WHERE [index] = 0"
			rsHist.Open sqlHist, g_strCONNHist, 1, 3 
			rsHist.AddNew
			rsHist("reqID") = myID
			rsHist("Creator") = Request.Cookies("LBUsrName")
			rsHist("date") = Z_DateNull(Request("txtAppDate"))
			rsHist("dateTS") = TimeNow
			rsHist("dateU") = Request.Cookies("LBUsrName")
			rsHist("Stime") = Z_DateNull(Request("txtAppDate") & " " & Request("txtAppTFrom"))
			rsHist("StimeTS") = TimeNow
			rsHist("StimeU") = Request.Cookies("LBUsrName")
			If Request("chkClient") <> "" Then
				tmpHistAdr = CleanMe(Request("txtCliAdd")) & "|" & Request("txtCliCity") & "|" & Ucase(Request("txtCliState")) & "|" & Request("txtCliZip")
			Else
				tmpHistAdr = GetInstAdr(myDept)
			End If
			rsHist("location") = tmpHistAdr
			rsHist("locationTS") = TimeNow
			rsHist("locationU") = Request.Cookies("LBUsrName")
			rsHist.Update
			rsHist.Close
			Set rsHist = Nothing
			If SaveHist(myID, "waction.asp") Then
			
			End If
			Call ActiveSage(myDept)
			
			
			'dont send if ASL and no appt hours(479)
			If Request("selLang") <> 52 And Request("selLang") <> 78 And Request("selLang") <> 81 And Request("selLang") <> 90 _
				And Request("selLang") <> 85 And Request("chkblock") = "" And Request("myInst") <> 479 Then 
				'send job to intr
				call Z_EmailJob(myID) 
			End If
	  	ctrApp = ctrApp + 1
	  Loop
	End If
	
	rsWiz.Close

	Set rsWiz = Nothing
	
	'DELETE RECORD ON WIZARD DB
	Set rsWiz = Server.CreateObject("ADODB.RecordSet")
	sqlWiz = "DELETE FROM wrequest_T WHERE [index] = " & Request("tmpID")
	rsWiz.Open sqlWiz, g_strCONNW, 1, 3
	Set rsLB = Nothing

	Response.Redirect "reqconfirm.asp?ID=" & myID
End If
%>