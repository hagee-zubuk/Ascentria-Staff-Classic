<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
server.scripttimeout = 600000 '10mins
'USER CHECK
If Cint(Request.Cookies("LBUSERTYPE")) = 2 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If
Function GetMyStatus(xxx)
	Select Case xxx
		Case 1
			GetMyStatus = "COMPLETED"
		Case 2
			GetMyStatus = "MISSED"
		Case 3
			GetMyStatus = "CANCELED"
		Case 4
			GetMyStatus = "CANCELED (BILLABLE)"
		Case Else
			GetMyStatus = "PENDING"
	End Select
End Function
Function CleanMe(xxx)
	CleanMe = xxx
	If Not IsNull(xxx) Or xxx <> "" Then CleanMe = Replace(xxx, "'", " ")
End Function
Function CleanFax(strFax)
	CleanFax = Replace(strFax, "-", "") 
End Function
Function GetPrime(xxx)
	'get primary number of requesting person
	GetPrime = ""
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT * FROM requester_T WHERE [index] = " & xxx
	rsRP.Open sqlRP, g_strCONN, 3, 1
	If Not rsRP.EOF Then
		If rsRP("prime") = 0 Then
			GetPrime = rsRP("Email")
		ElseIf rsRP("prime") = 1 Then
			'GetPrime = rsRP("Phone")
			GetPrime = ""
		ElseIf rsRP("prime") = 2 Then
			GetPrime = CleanFax(Trim(rsRP("Fax"))) & "@emailfaxservice.com" 
		End If
	End If
	rsRP.Close
	set rsRP = Nothing
End Function
Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function
server.scripttimeout = 360000
tmpPage = "document.frmAssign."
'GET REQUEST
Set rsConfirm = Server.CreateObject("ADODB.RecordSet")
sqlConfirm = "SELECT * FROM Request_T WHERE [index] = " & Request("ID")
rsConfirm.Open sqlConfirm, g_strCONN, 3, 1
If Not rsConfirm.EOF Then
	TS = rsConfirm("timestamp")
	tmpReqP = rsConfirm("reqID") 
	tmpClient = ""
	tmpDeptaddr = ""
	If rsConfirm("client") = True Then tmpClient = " (LSS Client)"
	tmpName = rsConfirm("clname") & ", " & rsConfirm("cfname") & tmpClient
	tmpAddr = rsConfirm("caddress") & ", " & rsConfirm("CliAdrI") & ", " & rsConfirm("cCity") & ", " &  rsConfirm("cstate") & ", " & rsConfirm("czip")
	If rsConfirm("CliAdd") = True Then 
		tmpDeptaddrG = rsConfirm("CAddress") &", " & rsConfirm("CCity") & ", " & rsConfirm("CState") & ", " & rsConfirm("CZip")
	End If
	tmpFon = rsConfirm("Cphone")
	tmpAFon = rsConfirm("CAphone")
	tmpDir = rsConfirm("directions")
	tmpSC = rsConfirm("spec_cir")
	tmpDOB = rsConfirm("DOB")
	tmpLang = rsConfirm("langID")
	tmpAppDate = rsConfirm("appDate")
	tmpAppTFrom = rsConfirm("appTimeFrom") 
	tmpAppTTo = rsConfirm("appTimeTo")
	tmpAppLoc = rsConfirm("appLoc")
	tmpInst = rsConfirm("instID")
	tmpDept = rsConfirm("DeptID")
	tmpInstRate = Z_FormatNumber(rsConfirm("InstRate"), 2)
	tmpDoc = rsConfirm("docNum")
	tmpCRN = rsConfirm("CrtRumNum")
	tmpIntr = rsConfirm("IntrID")
	tmpIntrRate = Z_FormatNumber(rsConfirm("IntrRate"), 2)
	tmpCom = rsConfirm("Comment")
	Statko = GetMyStatus(rsConfirm("Status"))
	tmpBilHrs = rsConfirm("Billable")
	tmpActTFrom = Z_FormatTime(rsConfirm("astarttime")) 
	tmpActTTo = Z_FormatTime(rsConfirm("aendtime"))
	tmpBilTInst = rsConfirm("TT_Inst")
	tmpBilTIntr = rsConfirm("TT_Intr")
	tmpBilMInst = rsConfirm("M_Inst")
	tmpBilMIntr = rsConfirm("M_Intr")
	tmpEmer = ""
	If rsConfirm("Emergency") = True Then tmpEmer = "(EMERGENCY)" 
	If rsConfirm("emerFEE") = True Then tmpEmer = "(EMERGENCY - Fee applied)"
	tmpHPID = Z_CZero(rsConfirm("HPID"))
	tmpcomintr = rsConfirm("intrcomment")
	tmpcombil = rsConfirm("bilcomment")
	tmpLBcom = rsConfirm("LBcomment")
End If
rsConfirm.Close
Set rsConfirm = Nothing
If tmpHPID > 0 And Request.Cookies("UID") <> 2 And Request.Cookies("UID") <> 5 Then hpid = "disabled"
'RESUPPLY DATA IN ERROR EVENT
If Session("MSG") <> "" And Request("ID") = "" Then
	tmpEntry = Split(Z_DoDecrypt(Request.Cookies("LBREQUEST")), "|")
	tmpNewInst = Split(Z_DoDecrypt(Request.Cookies("LBINST")), "|")
	tmpNewDept = Split(Z_DoDecrypt(Request.Cookies("LBDEPT")), "|")
	tmpNewReq = Split(Z_DoDecrypt(Request.Cookies("LBREQ")), "|")
	tmpNewIntr = Split(Z_DoDecrypt(Request.Cookies("LBINTR")), "|")
	tmpNewIntrBTN = tmpNewIntr(10)
	If tmpNewReq(6) = "BACK" Then
		tmpNewReqLN = tmpNewReq(0)
		tmpNewReqFN = tmpNewReq(1)
		tmpReqExt = tmpNewReq(8)
		tmpNewReqPhone = tmpNewReq(2)
		tmpNewReqEmail = tmpNewReq(3)
		tmpNewReqFax = tmpNewReq(4)
		tmpNewReqPrim = tmpNewReq(7)
		selRPEmail = ""
		selRPPhone = ""
		selRPFax = ""
		if tmpNewReqPrim = 0 Then selRPEmail = "checked"
		if tmpNewReqPrim = 1 Then selRPPhone = "checked"
		if tmpNewReqPrim = 2 Then selRPFax = "checked"
	Else
		tmpReqP = tmpEntry(1)
	End If
	tmplName = tmpEntry(2)
	tmpfName = tmpEntry(3)
	chkClient = ""
	If tmpEntry(20) <> "" Then chkClient = "checked"
	tmpAddr = tmpEntry(4)
	chkUClientadd = ""
	If tmpEntry(29) <> "" Then chkUClientadd = "checked"
	tmpCity = tmpEntry(5)
	tmpState = tmpEntry(6)
	tmpZip = tmpEntry(7)
	tmpCAdrI = tmpEntry(30)
	tmpDir = tmpEntry(8)
	tmpSC = tmpEntry(9)
	tmpDOB = tmpEntry(10)
	tmpLang = tmpEntry(11)
	tmpAppDate = tmpEntry(12)
	tmpAppTFrom = Z_FormatTime(tmpEntry(13))
	tmpAppTTo = Z_FormatTime(tmpEntry(14))
	tmpAppLoc = tmpEntry(15)
	If tmpNewInst(6) = "BACK" Then 
		tmpNewInstchk = tmpNewInst(6)
		tmpNewInstTxt = tmpNewInst(0)
	Else
		tmpInst = tmpEntry(16)
	End If
	If tmpNewDept(6) = "BACK" Then
		tmpNewInstchk = tmpNewDept(6)
		tmpNewInstDept = tmpNewDept(0)
		SocSer = ""
		Priv = ""
		Legal = ""
		Med = ""
		tmpClass = tmpNewDept(7)
		If tmpNewDept(7) = 1 Then SocSer = "selected"
		If tmpNewDept(7) = 2 Then Priv = "selected"
		If tmpNewDept(7) = 3 Then Court = "selected"
		If tmpNewDept(7) = 4 Then Med = "selected"
		If tmpNewDept(7) = 5 Then Legal = "selected"
		tmpNewInstAddr = tmpNewDept(2)
		tmpNewInstCity = tmpNewDept(3)
		tmpNewInstState = tmpNewDept(4)
		tmpNewInstZip = tmpNewDept(5)
		tmpNewInstAddrI = tmpNewDept(16)
		If tmpNewInst(8) <> "" Then 
			chkBillMe = "checked"
		Else
			tmpBilInstAddr = tmpNewDept(9)
			tmpBilInstCity = tmpNewDept(10)
			tmpBilInstState = tmpNewDept(11)
			tmpBilInstZip = tmpNewDept(12)
		End If	
		tmpBLname =  tmpNewDept(13)
	Else
		tmpDept = tmpEntry(26)
	End If
	tmpRate = tmpEntry(17)
	tmpDoc = tmpEntry(18)
	tmpCRN = tmpEntry(19)
	tmpCFon = tmpEntry(21)
	If Request.Cookies("LBACTION") = 1 Then
		tmpCAFon = tmpEntry(27)
	Else
		tmpCAFon = tmpEntry(38)
	End If
	If tmpNewIntrBTN = "BACK" Then
		tmpIntrLname = tmpNewIntr(0)
		tmpIntrFname = tmpNewIntr(1)
		tmpIntrEmail = tmpNewIntr(2)
		tmpIntrP1 = tmpNewIntr(3)
		tmpIntrExt = tmpNewIntr(13)
		tmpIntrFax = tmpNewIntr(4)
		tmpIntrP2 = tmpNewIntr(5)
		tmpIntrAddr = tmpNewIntr(6)
		tmpIntrCity = tmpNewIntr(7)
		tmpIntrState = tmpNewIntr(8)
		tmpNewIntrAddrI = tmpNewIntr(16)
		tmpIntrZip = tmpNewIntr(9)
		tmpInHouse = ""
		If tmpNewIntr(11) <> "" Then tmpInHouse = "checked"
		tmpIntrPrim = tmpNewIntr(12)
		selIntrFax = ""
		selIntrP2 = ""
		selIntrP1 = ""
		selIntrEmail = ""
		if tmpIntrPrim = 0 Then selIntrEmail = "checked"
		if tmpIntrPrim = 1 Then selIntrP1 = "checked"
		if tmpIntrPrim = 2 Then selIntrP2 = "checked"
		if tmpIntrPrim = 0 Then selIntrFax = "checked"
		tmpIntrRate =  tmpNewIntr(14)
	Else
		tmpIntr = tmpEntry(22)
	End If
	tmpEmer = ""
	If tmpEntry(24) <> "" Then tmpEmer = "checked"
	If Request.Cookies("LBACTION") = 1 Then	
		tmpCom =  tmpEntry(25)
	Else
		tmpCom =  tmpEntry(30)
	End If
End If

'GET INSTITUTION
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM institution_T WHERE [index] = " & tmpInst
rsInst.Open sqlInst, g_strCONN, 3, 1
If Not rsInst.EOF Then
	tmpIname = rsInst("Facility") 
End If
rsInst.Close
Set rsInst = Nothing 
'GET DEPARTMENT
Set rsDept = Server.CreateObject("ADODB.RecordSet")
sqlDept = "SELECT * FROM dept_T WHERE Active = 1 AND [index] = " & tmpDept
rsDept.Open sqlDept, g_strCONN, 3, 1
If Not rsDept.EOF Then
	tmpDname = rsDept("dept") 
	tmpDeptaddr = rsDept("InstAdrI") & " " & rsDept("address") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	tmpBaddr = rsDept("Baddress") & ", " & rsDept("BCity") & ", " &  rsDept("Bstate") & ", " & rsDept("Bzip")
	tmpBContact = rsDept("Blname")
End If
rsDept.Close
Set rsDept = Nothing 
'GET LANGUAGE
Set rsLang = Server.CreateObject("ADODB.RecordSet")
sqlLang  = "SELECT * FROM language_T WHERE [index] = " & tmpLang
rsLang.Open sqlLang , g_strCONN, 3, 1
If Not rsLang.EOF Then
	tmpSalita = rsLang("language") 
End If
rsLang.Close
Set rsLang = Nothing 
'GET INSTITUTION RATES
Set rsRate = Server.CreateObject("ADODB.RecordSet")
sqlRate = "SELECT * FROM rate_T ORDER BY Rate"
rsRate.Open sqlRate, g_strCONN, 3, 1
Do Until rsRate.EOF
	RateKo = ""
	If tmpInstRate = rsRate("Rate") Then Rateko = "selected"
	strRate1 = strRate1 & "<option " & Rateko & " value='" & rsRate("Rate") & "'>$" & Z_FormatNumber(rsRate("Rate"), 2) & "</option>" & vbCrLf
	rsRate.MoveNext
Loop
rsRate.Close
Set rsRate = Nothing
'GET AVAILABLE REQUESTING PERSON PER DEPARTMENT
Set rsInstReq = Server.CreateObject("ADODB.RecordSet")
sqlInstReq = "SELECT * FROM dept_T WHERE Active = 1 AND InstID = " & tmpInst & " ORDER BY dept"
rsInstReq.Open sqlInstReq, g_strCONN, 3, 1
Do Until rsInstReq.EOF
	tmpDpt = ""
	If Cint(tmpDept) = rsInstReq("index") Then tmpDpt = "selected"
	DeptName = rsInstReq("Dept")
	strDept2 = strDept2	& "<option " & tmpDpt & " value='" & rsInstReq("Index") & "'>" &  DeptName & "</option>" & vbCrlf
	
	tmpOLDAddr = rsInstReq("address") & "|" & rsInstReq("city") & "|" & rsInstReq("state") & "|" & rsInstReq("zip")
	strDept = strDept & "if (dept == " & rsInstReq("Index") & ") " & vbCrLf & _
		"{document.frmAssign.txtInstAddr.value = """ & rsInstReq("address") &"""; " & vbCrLf & _
		"document.frmAssign.selDept.value = " & rsInstReq("Index") & "; " & vbCrLf & _
		"document.frmAssign.txtInstCity.value = """ & rsInstReq("city") &"""; " & vbCrLf & _
		"document.frmAssign.txtInstState.value = """ & rsInstReq("state") &"""; " & vbCrLf & _
		"document.frmAssign.txtInstZip.value = """ & rsInstReq("zip") &"""; " & vbCrLf & _
		"document.frmAssign.txtInstAddrI.value = """ & rsInstReq("InstAdrI") &"""; " & vbCrLf & _
		"document.frmAssign.txtBlname.value = """ & rsInstReq("BLname") &"""; " & vbCrLf & _
		"document.frmAssign.txtBilAddr.value = """ & rsInstReq("Baddress") &"""; " & vbCrLf & _
		"document.frmAssign.txtBilCity.value = """ & rsInstReq("Bcity") &"""; " & vbCrLf & _
		"document.frmAssign.txtBilState.value = """ & rsInstReq("Bstate") &"""; " & vbCrLf & _
		"document.frmAssign.txtBilZip.value = """ & rsInstReq("Bzip") &"""; " & vbCrLf
		If tmpInstRate = 0 Then strDept = strDept & "document.frmAssign.selInstRate.value = """ & rsInstReq("defrate") &"""; " & vbCrLf 
		strDept = strDept & "document.frmAssign.OldAddr.value = """ & tmpOLDAddr &"""; " & vbCrLf & _
		"document.frmAssign.selClass.value = """ & GetClass(rsInstReq("Class")) &"""; }" & vbCrLf 
	
	InstReq = rsInstReq("index")
	strInstReqDept = strInstReqDept & "if (dept == " & InstReq & "){" & vbCrLf
	Set rsReqInst = Server.CreateObject("ADODB.RecordSet")
	sqlReqInst = "SELECT requester_T.[index] as rpID, lname, fname, Phone, pExt, Fax, Email, prime, aphone FROM requester_T, reqdept_T WHERE  ReqID = requester_T.[index] AND DeptID = " & InstReq & " ORDER BY lname, fname"
	rsReqInst.Open sqlReqInst, g_strCONN, 3, 1
	Do Until rsReqInst.EOF
		
		If tmpReqP = "" Then tmpReqP = -1
	If CInt(tmpReqP) = rsReqInst("rpID") Then ReqSel = "selected"
	tmpReqName = CleanMe(rsReqInst("lname")) & ", " & CleanMe(rsReqInst("fname"))
	strReq2 = strReq2 & "<option " & ReqSel & " value='" & rsReqInst("rpID") & "'>" & tmpReqName & "</option>" & vbCrLf
	
		tmpReqName = CleanMe(rsReqInst("lname")) & ", " & CleanMe(rsReqInst("fname"))
		strInstReqDept = strInstReqDept	& "if(req != "& rsReqInst("rpID") & ")" & vbCrLf & _
			"{var ChoiceReq = document.createElement('option');" & vbCrLf & _
			"ChoiceReq.value = " & rsReqInst("rpID") & ";" & vbCrLf & _
			"ChoiceReq.appendChild(document.createTextNode(""" & tmpReqName & """));" & vbCrLf & _
			"document.frmAssign.selReq.appendChild(ChoiceReq);}" & vbCrLf
		
		strJScript3 = strJScript3 & "if (Req == " & rsReqInst("rpID") & ") " & vbCrLf & _
		"{document.frmAssign.txtphone.value = """ & rsReqInst("Phone") &"""; " & vbCrLf & _
		"document.frmAssign.selReq.value = " & rsReqInst("rpID") & "; " & vbCrLf & _
		"document.frmAssign.txtReqExt.value = """ & rsReqInst("pExt") &"""; " & vbCrLf & _
		"document.frmAssign.txtfax.value = """ & rsReqInst("Fax") &"""; " & vbCrLf & _
		"document.frmAssign.txtaphone.value = """ & rsReqInst("aphone") &"""; " & vbCrLf & _
		"document.frmAssign.txtemail.value = """ & rsReqInst("Email") &"""; " & vbCrLf
		If rsReqInst("prime") = 0 Then
			strJScript3 = strJScript3 & "document.frmAssign.radioPrim1[2].checked = true;" & vbCrLf 
		ElseIf rsReqInst("prime") = 1 Then
			strJScript3 = strJScript3 & "document.frmAssign.radioPrim1[0].checked = true;" & vbCrLf 
		ElseIf rsReqInst("prime") = 2 Then
			strJScript3 = strJScript3 & "document.frmAssign.radioPrim1[1].checked = true;" & vbCrLf 
		End If
		strJScript3 = strJScript3 & "}"	
		
		rsReqInst.MoveNext
	Loop
	rsReqInst.Close
	Set rsReqInst = Nothing
	rsInstReq.MoveNext
	strInstReqDept = strInstReqDept & "}"
Loop
rsInstReq.Close
Set rsLangIntr = Nothing

'GET AVAILABLE DEPARTMENTS
Set rsInstDept = Server.CreateObject("ADODB.RecordSet")
sqlInstDept = "SELECT * FROM institution_T WHERE [index] = " & tmpInst
rsInstDept.Open sqlInstDept, g_strCONN, 3, 1
Do Until rsInstDept.EOF
	tmpDO = ""
	If Cint(tmpInst) = rsInstDept("index") Then tmpDO = "selected"
	InstName = rsInstDept("Facility")
	strInst = strInst	& "<option " & tmpDO & " value='" & rsInstDept("Index") & "'>" &  InstName & "</option>" & vbCrlf
	
	InstDept = rsInstDept("Index")
	strInstDept = strInstDept & "if (inst == " & InstDept & "){" & vbCrLf
	Set rsDeptInst = Server.CreateObject("ADODB.RecordSet")
	sqlDeptInst = "SELECT * FROM dept_T WHERE Active = 1 AND InstID = " &  InstDept & " ORDER BY Dept"
	rsDeptInst.Open sqlDeptInst, g_strCONN, 3, 1
	If Not rsDeptInst.EOF Then
		Do Until rsDeptInst.EOF
			strInstDept = strInstDept & "if (dept != " & rsDeptInst("index") & ")" & vbCrLf & _
				"{var ChoiceInst = document.createElement('option');" & vbCrLf & _
				"ChoiceInst.value = " & rsDeptInst("index") & ";" & vbCrLf & _
				"ChoiceInst.appendChild(document.createTextNode(""" & rsDeptInst("Dept") & """));" & vbCrLf & _
				"document.frmAssign.selDept.appendChild(ChoiceInst);} " & vbCrlf
			rsDeptInst.MoveNext
		Loop
	End If
	rsDeptInst.Close
	Set rsDeptInst = Nothing
	rsInstDept.MoveNext
	strInstDept = strInstDept & "}"
Loop
rsInstDept.Close
Set rsInstDept = Nothing
'GET INTERPRETER INFO
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT InHouse, [Last Name], [First Name]  FROM interpreter_T WHERE [index] = " & tmpIntr
rsIntr.Open sqlIntr, g_strCONN, 3, 1
If Not rsIntr.EOF Then
	tmpInHouse = ""
	If rsIntr("InHouse") = 1 Then tmpInHouse = "(In-House)"
	tmpIntrName = rsIntr("Last Name") & ", " & rsIntr("First Name") & " " & tmpInHouse
Else
	tmpIntrName = "<i>To be assigned.</i>"
	tmpIntr = 0
End If
rsIntr.Close
Set rsIntr = Nothing
'HP DATA
If tmpHPID <> 0  THen
	Set rsHP = Server.CreateObject("ADODB.RecordSet")
		sqlHP = "SELECT * FROM Appointment_T WHERE [index] = " & tmpHPID
	rsHP.Open sqlHP, g_StrCONNHP, 3, 1
	If Not rsHP.EOF Then
		tmpCallMe = ""
		If rsHP("callme") = True Then tmpCallMe = "* Call patient to remind of appointment"
		tmpReason = GetReas(Z_Replace(rsHP("reason"),", ", "|"))
		tmpClin = rsHP("clinician")  
		InHP = 0
		tmpMeet = ""
		If rsHP("mwhere") = 1 Then
			InHP = 1
			tmpMeet = UCase(GetLoc(rsHP("mlocation")))
			If tmpMeet = "OTHER" Then tmpMeet = rsHP("mother")
		End If
		tmpMinor = ""
		If rsHP("minor") = True Then tmpMinor = "* Minor"
		tmpParents = ""
		If rsHP("parents") <> "" Then tmpParents = rsHP("parents") 
	End If
	rsHp.Close
	Set rsHp = Nothing
End If
%>
<!-- #include file="_closeSQL.asp" -->
<html>
	<head>
		<title>Language Bank - Interpreter Request Form - Edit Contact - <%=Request("ID")%></title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function bawal(tmpform)
		{
			var iChars = ",|\"\'";
			var tmp = "";
			for (var i = 0; i < tmpform.value.length; i++)
			 {
			  	if (iChars.indexOf(tmpform.value.charAt(i)) != -1)
			  	{
			  		alert ("This character is not allowed.");
			  		tmpform.value = tmp;
			  		return;
		  		}
			  	else
		  		{
		  			tmp = tmp + tmpform.value.charAt(i);
		  		}
		  	}
		}
		function maskMe(str,textbox,loc,delim)
		{
			var locs = loc.split(',');
			for (var i = 0; i <= locs.length; i++)
			{
				for (var k = 0; k <= str.length; k++)
				{
					 if (k == locs[i])
					 {
						if (str.substring(k, k+1) != delim)
					 	{
					 		str = str.substring(0,k) + delim + str.substring(k,str.length);
		     			}
					}
				}
		 	}
			textbox.value = str
		}
		function SaveAss(xxx)
		{
			<% If tmpHPID <> 0 Then %>
				if (document.frmAssign.hideHPID.value != 0)
				{
					if (document.frmAssign.selInst.value != document.frmAssign.hideInst.value)
					{
						alert("ERROR: User cannot change Institution.")
						document.frmAssign.selInst.value =  document.frmAssign.hideInst.value;
						return;
					}
					if (document.frmAssign.selDept.value != document.frmAssign.hideDept.value)
					{
						alert("ERROR: User cannot change Department.")
						document.frmAssign.selDept.value =  document.frmAssign.hideDept.value;
						return;
					}
					if (document.frmAssign.selReq.value != document.frmAssign.hideReq.value)
					{
						alert("ERROR: User cannot change Requesting Person.")
						document.frmAssign.selReq.value =  document.frmAssign.hideReq.value;
						return;
					}
				}
			<% End If %>
			//CHECK IF ADDRESS HAS BEEN CHANGED
			var strNewAddr = document.frmAssign.txtInstAddr.value + "|" + document.frmAssign.txtInstCity.value + "|" + document.frmAssign.txtInstState.value + "|" + document.frmAssign.txtInstZip.value;
			if (strNewAddr != document.frmAssign.OldAddr.value)
			{
				var ans = window.confirm("WARNING: Changing of institution address will be effective for all instances of that institution. Click Cancel to stop.");
				if (!ans)
				{
					return;
				}
			}
			if (document.frmAssign.radioPrim1[2].checked == true && document.frmAssign.txtemail.value == "")
			{
				alert("ERROR: Please supply an E-mail address to requesting person."); 
				document.frmAssign.txtemail.focus();
				return;
			}
			if (document.frmAssign.radioPrim1[0].checked == true && document.frmAssign.txtphone.value == "")
			{
				alert("ERROR: Please supply a Phone Number to requesting person."); 
				document.frmAssign.txtphone.focus();
				return;
			}
			if (document.frmAssign.radioPrim1[1].checked == true && document.frmAssign.txtfax.value == "")
			{
				alert("ERROR: Please supply a Fax Number to requesting person."); 
				document.frmAssign.txtfax.focus();
				return;
			}
			//CHECK VALID FAX
			if (document.frmAssign.radioPrim1[1].checked == true && document.frmAssign.txtfax.value != "")
			{
				var tmpFax =  document.frmAssign.txtfax.value
				tmpFax = tmpFax.replace("-", "")
				if (tmpFax.length < 10) 
				{
					alert("ERROR: Please include area code in Fax Number to requesting person."); 
					document.frmAssign.txtfax.focus();
					return;
				}
			}
			<%=strReqCHK%>
			if (pnt = 1)
			{
				document.frmAssign.action = "action.asp?ctrl=12&ReqID=" + xxx;
				document.frmAssign.submit();
			} 
		}
		function CalendarView(strDate)
		{
			document.frmAssign.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmAssign.submit();
		}
		function chkPrim()
		{
			if (document.frmAssign.radioPrim1[0].checked == true)
			{
				document.frmAssign.txtPRim1.value = "Phone";
			}
			if (document.frmAssign.radioPrim1[1].checked == true)
			{
				document.frmAssign.txtPRim1.value = "Fax";
			}
			if (document.frmAssign.radioPrim1[2].checked == true)
			{
				document.frmAssign.txtPRim1.value = "E-Mail";
			}
		}
		function InstInfo(Inst)
		{
			<%=strJScript%>
			if (Inst == -1)
			{
				document.frmAssign.selInst.value = -1; 
				document.frmAssign.txtInstDept.value = ""; 
				document.frmAssign.txtInstAddr.value = ""; 
				document.frmAssign.txtInstCity.value = ""; 
				document.frmAssign.txtInstState.value = ""; 
				document.frmAssign.txtInstZip.value = ""; 
				document.frmAssign.txtInstAddrI.value = "";
				document.frmAssign.txtBlname.value = ""; 
				document.frmAssign.txtBilAddr.value = ""; 
				document.frmAssign.txtBilCity.value = ""; 
				document.frmAssign.txtBilState.value = ""; 
				document.frmAssign.txtBilZip.value = ""; 
				document.frmAssign.selClass.value = "1";
			}
		}
		function hideNewInts() 
		{
			if (document.frmAssign.txtNewInst.value == "")
			{	
				document.frmAssign.txtNewInst.style.visibility = 'hidden';
				//document.frmAssign.btnNew.value = 'NEW';
				document.frmAssign.txtNewInst.value = "";
				document.frmAssign.HnewInt.value = 'NEW';
			}
			else
			{
				document.frmAssign.txtNewInst.style.visibility = 'visible';
				//document.frmAssign.btnNew.value = 'BACK';
				document.frmAssign.selInst.disabled = true;
				document.frmAssign.txtInstDept.value = '<%=tmpNewInstDept%>';
				document.frmAssign.txtInstAddr.value = '<%=tmpNewInstAddr%>';
				document.frmAssign.txtInstCity.value = '<%=tmpNewInstCity%>';
				document.frmAssign.txtInstState.value = '<%=tmpNewInstState%>';
				document.frmAssign.txtInstZip.value = '<%=tmpNewInstZip%>';
				document.frmAssign.txtInstAddrI.value = '<%=tmpNewInstAddrI%>';
				document.frmAssign.HnewInt.value = 'BACK';
			}
		}
		function hideNewReq() 
		{
			if (document.frmAssign.txtReqLname.value == "" && document.frmAssign.txtReqFname.value == "")
			{	
				document.frmAssign.txtReqLname.style.visibility = 'hidden';
				document.frmAssign.txtReqFname.style.visibility = 'hidden';
				document.frmAssign.txtcoma2.style.visibility = 'hidden';
				document.frmAssign.txtformat2.style.visibility = 'hidden';
				//document.frmAssign.btnNewReq.value = 'NEW';
				document.frmAssign.txtReqLname.value = "";
				document.frmAssign.txtReqFname.value = "";
				document.frmAssign.HnewReq.value = 'NEW';
			}
			else
			{
				document.frmAssign.txtReqLname.style.visibility = 'visible';
				document.frmAssign.txtReqFname.style.visibility = 'visible';
				document.frmAssign.txtcoma2.style.visibility = 'visible';
				document.frmAssign.txtformat2.style.visibility = 'visible';
				//document.frmAssign.btnNewReq.value = 'BACK';
				document.frmAssign.selReq.disabled = true;
				document.frmAssign.txtReqLname.value = '<%=tmpNewReqLN%>';
				document.frmAssign.txtReqFname.value = '<%=tmpNewReqFN%>';
				document.frmAssign.txtemail.value = '<%=tmpNewReqeMail%>';
				document.frmAssign.txtReqExt.value = '<%=tmpReqExt%>';
				document.frmAssign.txtphone.value = '<%=tmpNewReqPhone%>';
				document.frmAssign.txtfax.value = '<%=tmpNewReqFax%>';
				document.frmAssign.HnewReq.value = 'BACK';
			}
		}
		function textboxchangeDept()
		{

				//document.frmAssign.btnNewDept.value = 'NEW';
				document.frmAssign.selDept.disabled = false;
				document.frmAssign.txtInstDept.value = "";
				document.frmAssign.txtInstDept.style.visibility = 'hidden';
				DeptInfo(document.frmAssign.selDept.value);
				document.frmAssign.HnewDept.value = 'NEW';
		
		}
		function hideNewDept() 
		{
			if (document.frmAssign.txtInstDept.value == "")
			{	
				document.frmAssign.txtInstDept.style.visibility = 'hidden';
				//document.frmAssign.btnNewDept.value = 'NEW';
				document.frmAssign.txtInstDept.value = "";
				document.frmAssign.HnewDept.value = 'NEW';
			}
			else
			{
				document.frmAssign.txtInstDept.style.visibility = 'visible';
				//document.frmAssign.btnNewDept.value = 'BACK';
				document.frmAssign.selDept.disabled = true;
				document.frmAssign.selClass.value = '<%=tmpClass%>';
				document.frmAssign.txtInstAddr.value = '<%=tmpNewInstAddr%>';
				document.frmAssign.txtInstCity.value = '<%=tmpNewInstCity%>';
				document.frmAssign.txtInstState.value = '<%=tmpNewInstState%>';
				document.frmAssign.txtInstZip.value = '<%=tmpNewInstZip%>';
				document.frmAssign.txtInstAddrI.value = '<%=tmpNewInstAddrI%>';
				document.frmAssign.txtBlname.value = '<%=tmpBLname%>';
				if (document.frmAssign.chkBill.checked != true)
				{
					document.frmAssign.chkBill.checked = false;
					document.frmAssign.txtBilAddr.value = '<%=tmpBilInstAddr%>';
					document.frmAssign.txtBilCity.value = '<%=tmpBilInstCity%>';
					document.frmAssign.txtBilState.value = '<%=tmpBilInstState%>';
					document.frmAssign.txtBilZip.value = '<%=tmpBilInstZip%>';
				}
				else
				{
					document.frmAssign.chkBill.checked = true;
					document.frmAssign.txtBilAddr.value = "";
					document.frmAssign.txtBilCity.value = "";
					document.frmAssign.txtBilState.value = "";
					document.frmAssign.txtBilZip.value = "";
				}
				document.frmAssign.HnewDept.value = 'BACK';
			}
		}
		function DeptInfo(dept)
		{
			if (dept == 0 && document.frmAssign.txtInstDept.value == "" )
			{
				document.frmAssign.selDept.value =0;
				document.frmAssign.txtInstDept.value = "";
				document.frmAssign.txtInstAddr.value = "";
				document.frmAssign.txtInstCity.value = "";
				document.frmAssign.txtInstState.value = "";
				document.frmAssign.txtInstZip.value = "";
				document.frmAssign.txtInstAddrI.value = "";
				document.frmAssign.txtBlname.value = "";
				document.frmAssign.txtBilAddr.value = "";
				document.frmAssign.txtBilCity.value = "";
				document.frmAssign.txtBilState.value = "";
				document.frmAssign.txtBilZip.value = "";
				document.frmAssign.OldAddr.value = "";
			}
			else
			{
				hideNewDept();
			}
			<%=strDept%>
		}
		function DeptChoice(inst, dept)
		{
			var i;
			for(i=document.frmAssign.selDept.options.length-1;i>=1;i--)
			{
				if (dept != "undefined")
				{
					if (document.frmAssign.selDept.options[i].value != dept)
					{
						document.frmAssign.selDept.remove(i);
					}
				}
				else
				{
					document.frmAssign.selReq.remove(i);
				}
			}
			<%=strInstDept%>
		}
		function ReqChoice(dept, req)
		{
			 var i;
			for(i=document.frmAssign.selReq.options.length-1;i>=1;i--)
			{
				if (req != "undefined")
				{
					if (document.frmAssign.selReq.options[i].value != req)
					{
						document.frmAssign.selReq.remove(i);
					}
				}
				else
				{
					document.frmAssign.selReq.remove(i);
				}
			}
			<%=strInstReqDept%>
		}
		function ReqShowMe()
		{
			
				ReqChoice(document.frmAssign.selDept.value);
	
		}
		function ReqInfo(Req)
		{
			if (Req == " -1")
			{
				if  (document.frmAssign.txtReqLname.value == "" || document.frmAssign.txtReqFname.value == "")
					{
						hideNewReq();
					}
					else
					{document.frmAssign.txtphone.value = ""; 
					document.frmAssign.txtReqExt.value = ""; 
					document.frmAssign.radioPrim1[1].checked = true;
					document.frmAssign.txtfax.value = ""; 
					document.frmAssign.txtemail.value = ""; }
			}
			<%=strJScript3%>
			chkPrim();
		}
		function PopMe()
		{
			newwindow = window.open('find.asp','name','height=150,width=400,scrollbars=0,directories=0,status=0,toolbar=0,resizable=0');
			if (window.focus) {newwindow.focus()}
		}
		function textboxchangeInst() 
		{
			if (document.frmAssign.btnNew.value == 'NEW')
			{
				alert("To save a new Institution, complete the form and click 'Save Request' button.");
				//document.frmAssign.btnNew.value = 'BACK';
				document.frmAssign.selInst.disabled = true;
				document.frmAssign.txtNewInst.style.visibility = 'visible';
				document.frmAssign.txtInstDept.value = "";
				document.frmAssign.txtInstAddr.value = "";
				document.frmAssign.txtInstCity.value = "";
				document.frmAssign.txtInstState.value = "";
				document.frmAssign.txtInstZip.value = "";
				document.frmAssign.txtInstAddrI.value = "";
				document.frmAssign.txtBlname.value = ""; 
				document.frmAssign.txtBilAddr.value = ""; 
				document.frmAssign.txtBilCity.value = ""; 
				document.frmAssign.txtBilState.value = ""; 
				document.frmAssign.txtBilZip.value = ""; 
				document.frmAssign.selClass.value = "1";
				document.frmAssign.txtNewInst.focus();
				document.frmAssign.HnewInt.value = 'BACK';
				DeptChoice();
			}
			else
			{
				//document.frmAssign.btnNew.value = 'NEW';
				document.frmAssign.selInst.disabled = false;
				document.frmAssign.txtNewInst.value = "";
				document.frmAssign.txtNewInst.style.visibility = 'hidden';
				document.frmAssign.HnewInt.value = 'NEW';
				DeptChoice(document.frmAssign.selInst.value);
			}
		}
		function textboxchangeReq() 
		{
			
				//document.frmAssign.btnNewReq.value = 'NEW';
				document.frmAssign.selReq.disabled = false;
				document.frmAssign.txtReqLname.value = "";
				document.frmAssign.txtReqFname.value = "";
				document.frmAssign.txtReqLname.style.visibility = 'hidden';
				document.frmAssign.txtReqFname.style.visibility = 'hidden';
				document.frmAssign.txtcoma2.style.visibility = 'hidden';
				document.frmAssign.txtformat2.style.visibility = 'hidden';
				ReqInfo(document.frmAssign.selReq.value);
				document.frmAssign.HnewReq.value = 'NEW';
		
		}
		function ReqShowMe()
		{
			if (document.frmAssign.chkAll.checked == true) 
			{
				for(i=document.frmAssign.selReq.options.length-1;i>=1;i--)
				{
					document.frmAssign.selReq.remove(i);
				}
				<%=strReq%>
			}
			else
			{
				ReqChoice(document.frmAssign.selDept.value);
			}
		}
		-->
		</script>
		<body onload='InstInfo(<%=tmpInst%>); hideNewInts(); hideNewReq(); chkPrim(); hideNewDept();
			DeptInfo(<%=tmpDept%>); DeptChoice(<%=tmpInst%>, <%=tmpDept%>); ReqChoice(<%=tmpDept%>, <%=tmpReqP%>); ReqInfo(<%=tmpReqP%>);'>
			<form method='post' name='frmAssign'>
				<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
					<tr>
						<td valign='top'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<tr>
						<td valign='top' >
							<table cellSpacing='2' cellPadding='0' width="100%" border='0'>
								<!-- #include file="_greetme.asp" -->
								<tr>
								<td class='title' colspan='10' align='center'><nobr> Interpreter Request Form - Edit Contact</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td  align='center' colspan='10'>
										<div name="dErr" style="width: 250px; height:55px;OVERFLOW: auto;">
											<table border='0' cellspacing='1'>		
												<tr>
													<td><span class='error'><%=Session("MSG")%></span></td>
												</tr>
											</table>
										</div>
									</td>
								</tr>
								<tr>
									<td class='header' colspan='10'><nobr>Contact Information </td>
								</tr>
								<tr>
									<td align='right'>Request ID:</td>
									<td class='confirm' width='300px'><%=Request("ID")%>&nbsp;<%=tmpEmer%>
									<input type='hidden' name='hideHPID' value='<%=Z_CZero(tmpHPID)%>'>	
									</td>
								</tr>
								<tr>
									<td align='right'>Timestamp:</td>
									<td class='confirm' width='300px'><%=TS%></td>
								</tr>
								<tr>
									<td align='right'>Status:</td>
									<td class='confirm' width='300px'><%=Statko%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
										<td align='right'>*Institution:</td>
										<td width='350px'>
											<select class='seltxt' name='selInst'  style='width:250px;' onfocus='DeptChoice(document.frmAssign.selInst.value);DeptInfo(document.frmAssign.selDept.value); ' onchange='DeptChoice(document.frmAssign.selInst.value);DeptInfo(document.frmAssign.selDept.value); '>
												<option value='-1'>&nbsp;</option>
												<%=strInst%>
											</select>
											<!--<input type='button'  disabled value="FIND" <%=HpLock%>  name="findReq" onclick='PopMe();' title='Search instiution' class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
											<input class='btnLnk' disabled type='button' name='btnNew' value='NEW'  <%=HpLock%> onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'"
											onclick='textboxchangeInst();'>//-->
											<input type='hidden' name='hideInst' value='<%=tmpInst%>'>
										</td>
										<td>
											<input type='hidden' name='HnewInt'>
										</td>
									</tr>
									<tr>
										<td align='right'>&nbsp;</td>
										<td><input size='50' class='main' maxlength='50' name='txtNewInst' value='<%=tmpNewInstTxt%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td align='right' width='15%'>*Department:</td>
										<td>	
											<select class='seltxt' name='selDept'  style='width:250px;' onfocus='DeptInfo(document.frmAssign.selDept.value); ReqChoice(document.frmAssign.selDept.value); '  onchange='DeptInfo(document.frmAssign.selDept.value); ReqChoice(document.frmAssign.selDept.value); '>
												<option value='0'>&nbsp;</option>
												<%=strDept2%>
											</select>
											<!--<input class='btnLnk' type='button' name='btnNewDept' <%=HpLock%> value='NEW' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'"
											onclick='textboxchangeDept();'>//-->
											<input type='hidden' name='HnewDept'>
											<input type='hidden' name='hideDept' value='<%=tmpDept%>'>
										</td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td><input class='main' size='50' maxlength='50' name='txtInstDept' value='<%=tmpNewInstDept%>' onkeyup='bawal(this);'></td>
										<% If Request.Cookies("LBUSERTYPE") = 1 Then %>
										<td align='right' width='15%'>Rate:</td>
										<td>
											
											<input class='main' size='5' maxlength='5'  readonly name='txtInstRate' value='<%=tmpInstRate%>'>
											<select class='seltxt' style='width: 70px;' name='selInstRate'>
												<option value='0' >&nbsp;</option>
												<%=strRate1%>
											</select>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">
												*Rate varies per request
											</span>
											<% Else %>
												<input class='main' size='5' maxlength='5' type='hidden' name='txtInstRate' value='<%=tmpInstRate%>'>
											<% End If %>
										</td>
									</tr>
									<tr>
										<td align='right'>Classification:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='selClass' value='<%=tmpClass%>' onkeyup='bawal(this);' Readonly>
										</td>
									</tr>
									<tr>
										<td align='right'>Apartment/Suite Number:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtInstAddrI' value='<%=tmpNewInstAddrI%>' onkeyup='bawal(this);' readonly>
										</td>
									</tr>
									<tr>
										<td align='right'>*Appointment Address:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtInstAddr' value='<%=tmpNewInstAddr%>' onkeyup='bawal(this);' readonly>
										</td>
									</tr>
									<tr>
										<td align='right'>City:</td>
										<td colspan='5'>
											<input class='main' size='25' maxlength='25' name='txtInstCity' value='<%=tmpNewInstCity%>' onkeyup='bawal(this);' readonly>&nbsp;State:
											<input class='main' size='2' maxlength='2' name='txtInstState' value='<%=tmpNewInstState%>' onkeyup='bawal(this);' readonly>&nbsp;Zip:
											<input class='main' size='10' maxlength='10' name='txtInstZip' value='<%=tmpNewInstZip%>' onkeyup='bawal(this);' readonly>
											<input type='hidden' name='OldAddr'>
										</td>
									</tr>
									<tr>
										<td align='right'>Billed To:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtBlname' value='<%=tmpBLname%>' onkeyup='bawal(this);' readonly>
										</td>
									</tr>
									<tr>
										<td align='right'>Billing Address:</td>
										
									</tr>
									<tr>
										<td align='right'>Address:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtBilAddr' value='<%=tmpBilInstAddr%>' onkeyup='bawal(this);' readonly>
										</td>
									</tr>
									<tr>
										<td align='right'>City:</td>
										<td colspan='5'>
											<input class='main' size='25' maxlength='25' name='txtBilCity' value='<%=tmpBilInstCity%>' onkeyup='bawal(this);' readonly>&nbsp;State:
											<input class='main' size='2' maxlength='2' name='txtBilState' value='<%=tmpBilInstState%>' onkeyup='bawal(this);' readonly>&nbsp;Zip:
											<input class='main' size='10' maxlength='10' name='txtBilZip' value='<%=tmpBilInstZip%>' onkeyup='bawal(this);' readonly>
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td align='right'>*Requesting Person:</td>
										<td width='200px'>
											<nobr>
											<select id='selReq' class='seltxt' name='selReq'  style='width:250px;' onfocus='JavaScript:ReqInfo(document.frmAssign.selReq.value);' onchange='JavaScript:ReqInfo(document.frmAssign.selReq.value);'>
												<option value='-1'>&nbsp;</option>
												<%=strReq2%>
											</select>
											<!--<input class='btnLnk' type='button' name='btnNewReq' <%=HpLock%> value='NEW' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'"
											onclick='textboxchangeReq();'>//-->
											<!--<input type='checkbox' <%=HpLock%> name='chkAll' onclick='ReqShowMe(); ReqInfo(document.frmAssign.selReq.value);'>
											Show All//-->
											<input type='hidden' name='HnewReq'>
											<input type='hidden' name='hideReq' value='<%=tmpReqP%>'>
										</td>
									</tr>
									<tr>
										<td align='right'>&nbsp;</td>
										<td align='left'>
											<input class='main' size='20' maxlength='20' name='txtReqLname' value='<%=tmpNewReqLN%>' onkeyup='bawal(this);'>
											<input class='trans' style='width: 5px;' name='txtcoma2' readonly value=', '>
											<input class='main' size='20' maxlength='20' name='txtReqFname' value='<%=tmpNewReqFN%>' onkeyup='bawal(this);'>
											<input class='transmall' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" size='22' name='txtformat2' readonly value='last name, first name'>
										</td>
									</tr>
									<tr>
										<td align='right'><b>*Contact Numbers:</b></td>
										<td align='left'><b>(any of the following)</b></td>
									</tr>
									<tr>
										<td align='right'>Primary:</td>
										<td>
											<input class='main2'  name='txtPRim1'  readonly size='6'>
										</td>
									</tr>
									<tr>
										<td align='right'>
											<input type='radio' name='radioPrim1' value='1' <%=selRPPhone%> onclick='chkPrim();'>
											Phone:
										</td>
										<td>
											<input class='main' size='12' maxlength='12' name='txtphone' value='<%=tmpNewReqPhone%>' onkeyup='bawal(this);' readonly    >
											&nbsp;Ext:<input class='main' size='12' maxlength='12' name='txtReqExt' value='<%=tmpReqExt%>' onkeyup='bawal(this);' readonly>
										</td>
										<td align='right'>
											<input type='radio' name='radioPrim1' value='2'  <%=selRPFax%> onclick='chkPrim(); ' >
											Fax:
										</td>
										<td width='300px'><input class='main' size='12' maxlength='12' name='txtfax' value='<%=tmpNewReqFax%>' onkeyup='bawal(this);' readonly></td>
									</tr>
									<tr>
										<td align='right'>
											<input type='radio' name='radioPrim1' value='0' <%=selRPEmail%> onclick='chkPrim();'>
											E-Mail:
										</td>
										<td><input class='main' size='50' maxlength='50' name='txtemail' value='<%=tmpNewReqeMail%>' onkeyup='bawal(this);' readonly></td>
										<td>&nbsp;</td>
										<td>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Please include area code on fax number</span>
										</td>
									</tr>
									<tr>
										<td align='right'>
											
											Alternate Phone:
										</td>
										<td><input class='main' size='12' maxlength='12' name='txtaphone' value='<%=tmpNewReqaphone%>' onkeyup='bawal(this);' Readonly></td>
										<td>&nbsp;</td>
										
									</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' class='header'><nobr>Appointment Information</td>
								</tr>
								<tr>
									<td align='right'>Client Name:</td>
									<td class='confirm'><%=tmpName%>
									<% If tmpHPID <> 0  Then%>
										&nbsp;<%=tmpMinor%>
									<% End If%>	
									</td>
								</tr>
								<tr>
									<td align='right'>Client Address:</td>
									<td class='confirm'><%=tmpAddr%></td>
								</tr>
								<tr>
									<td align='right'>Language:</td>
									<td class='confirm'><%=tmpSalita%></td>
								</tr>
								<tr>
									<td align='right'>Appointment Date:</td>
									<td class='confirm'><%=tmpAppDate%></td>
								</tr>
								<tr>
									<td align='right'>Appointment Time:</td>
									<td class='confirm'><%=tmpAppTFrom%> - <%=tmpAppTTo%></td>
								</tr>
								<tr>
									<td align='right'>Docket Number:</td>
									<td class='confirm'><%=tmpDoc%></td>
								</tr>
								<tr>
									<td align='right'>Court Room No:</td>
									<td class='confirm'><%=tmpCRN%></td>
								</tr>
								<% If tmpHPID <> 0  Then%>
									<tr>
										<td align='right' valign='top'>Reason:</td>
										<td class='confirm'><%=tmpReason%></td>
									</tr>
									<tr>
										<td align='right'>Clinician:</td>
										<td class='confirm'><%=tmpClin%></td>
									</tr>
									<% If tmpParents <> "" Then%>
										<tr>
											<td align='right'>Parents:</td>
											<td class='confirm'><%=tmpParents%></td>
										</tr>
									<%End If%>
								<%End If%>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Appointment Comment:</td>
									<td class='confirm'><%=tmpCom%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' class='header'><nobr>Interpreter Information</td>
								</tr>
								<tr>
									<td align='right'>Interpreter:</td>
									<td class='confirm'><%=tmpIntrName%></td>
								</tr>
							
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Interpreter Comment:</td>
									<td class='confirm'><%=tmpComintr%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
										<tr><td>&nbsp;</td></tr>
									<tr>
									<td colspan='10' class='header'><nobr>Billing Information</td>
								</tr>
								<% If Request.Cookies("LBUSERTYPE") = 1 Then %>
								<tr>
									<td align='right'>Billable Hours:</td>
									<td class='confirm'><%=tmpBilHrs%></td>
								</tr>
								<tr>
									<td align='right'>Actual Time:</td>
									<td class='confirm'><%=tmpActTFrom%> - <%=tmpActTTo%></td>
								</tr>
								<tr>
									<td align='right'>&nbsp;</td>
									<td rowspan='3' valign='top'>
										<table cellSpacing='2' cellPadding='0' border='0'>
											<tr>
												<td align='left'>Bill To Institution </td>
												<td>|</td>
												<td>Pay To Interpreter</td>
											</tr>
											<tr>
												<td class='confirm' align='center'><%=tmpBilTInst%></td>
												<td>|</td>
												<td class='confirm' align='center'><%=tmpBilTIntr%></td>
											</tr>
											<tr>
												<td class='confirm' align='center'><%=tmpBilMInst%> </td>
												<td>|</td>
												<td class='confirm' align='center'> <%=tmpBilMIntr%></td>
											</tr>
										</table>
									</td>
								</tr>
								<tr>
									<td align='right'>Travel Time:</td>
								</tr>
								<tr>
									<td align='right'>Mileage:</td>
								</tr>
								<% end if %>
								<tr><td>&nbsp;</td></tr>
										<tr>
									<td align='right'>Billing Comment:</td>
									<td class='confirm'><%=tmpCombil%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' class='header'><nobr>Language Bank Notes	
									</td>
								</tr>
								<tr>
									<td align='right' valign='top'>Notes:</td>
									<td class='confirm'>
										<textarea name='txtLBcom' class='main' onkeyup='bawal(this);' style='width: 375px;' readonly><%=tmpLBcom%></textarea>
									</td>
								</tr>
										<tr><td>&nbsp;</td></tr>
										<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
										<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='10' align='center' height='100px' valign='bottom'>
												<input type='hidden' name="HID" value='<%=Request("ID")%>'>
												<input type='hidden' name="hidInstRate" value='<%=tmpInstRate%>'>
												<input type='hidden' name="hidIntrRate" value='<%=tmpIntrRate%>'>
												<input class='btn' type='button' value='Save' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" <%=hpid%> onclick='SaveAss(<%=Request("ID")%>) ;'>
												<input class='btn' type='button' value='Back' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='reqconfirm.asp?ID=<%=Request("ID")%>';">
											</td>
									</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td valign='bottom'>
							<!-- #include file="_footer.asp" -->
						</td>
					</tr>
				</table>
			</form>
		</body>
	</head>
</html>
<%
If Session("MSG") <> "" Then
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>