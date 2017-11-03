<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
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
	RP = rsConfirm("reqID") 
	tmpClient = ""
	tmpDeptaddr = ""
	tmplName = rsConfirm("clname") 
	tmpfName = rsConfirm("cfname")
	chkClient = ""
	If rsConfirm("Client") = True Then chkClient = "checked"
	chkUClientadd = ""
	If  rsConfirm("CliAdd")  = True Then chkUClientadd = "checked"
	tmpAddr = rsConfirm("caddress") 
	tmpCity = rsConfirm("cCity") 
	tmpState = rsConfirm("cstate") 
	tmpZip = rsConfirm("czip")
	tmpCAdrI = rsConfirm("CliAdrI")
	tmpCFon = rsConfirm("Cphone")
	tmpCAFon = rsConfirm("CAphone")
	tmpDir = rsConfirm("directions")
	tmpSC = rsConfirm("spec_cir")
	tmpDOB = rsConfirm("DOB")
	tmpLang = rsConfirm("langID")
	tmpAppDate = rsConfirm("appDate")
	tmpAppTFrom = Z_FormatTime(rsConfirm("appTimeFrom"))
	tmpAppTTo = Z_FormatTime(rsConfirm("appTimeTo"))
	tmpAppLoc = rsConfirm("appLoc")
	tmpInst = rsConfirm("instID")
	tmpDept = rsConfirm("deptID")
	tmpInstRate = rsConfirm("InstRate")
	tmpDoc = rsConfirm("docNum")
	tmpCRN = rsConfirm("CrtRumNum")
	tmpInst = rsConfirm("instID")
	tmpDept = rsConfirm("DeptID")
	tmpInstRate = Z_FormatNumber(rsConfirm("InstRate"), 2)
	tmpDoc = rsConfirm("docNum")
	tmpCRN = rsConfirm("CrtRumNum")
	tmpIntr = rsConfirm("IntrID")
	tmpIntrRate = Z_FormatNumber(rsConfirm("IntrRate"), 2)
	tmpEmer = ""
	If rsConfirm("Emergency") = True Then tmpEmer = "(EMERGENCY)" 
	tmpCom = rsConfirm("Comment")
	tmpComintr = rsConfirm("IntrComment")
	tmpcombil = rsConfirm("bilComment")
	Statko = GetMyStatus(rsConfirm("Status"))
	tmpBilHrs = rsConfirm("Billable")
	tmpActTFrom = Z_FormatTime(rsConfirm("astarttime")) 
	tmpActTTo = Z_FormatTime(rsConfirm("aendtime"))
	tmpBilTInst = rsConfirm("TT_Inst")
	tmpBilTIntr = rsConfirm("TT_Intr")
	tmpBilMInst = rsConfirm("M_Inst")
	tmpBilMIntr = rsConfirm("M_Intr")
	tmpLBcom = rsConfirm("LBcomment")
	'timestamp on sent/print
	tmpSentReq = "Request email has not been sent to Requesting Person."
	If rsConfirm("SentReq") <> "" Then tmpSentReq = "Request email was last sent to Requesting Person on <b>" & rsConfirm("SentReq") & "</b>."
	tmpSentIntr = "Request email has not been sent to Interpreter."
	If rsConfirm("SentIntr") <> "" Then tmpSentIntr = "Request email was last sent to Interpreter on <b>" & rsConfirm("SentIntr") & "</b>."
	tmpPrint = "Request has not been printed."
	If rsConfirm("Print") <> "" Then tmpPrint = "Request was last printed on<b> " & rsConfirm("Print") & "</b>."
	tmpHPID = Z_CZero(rsConfirm("HPID"))
End If
rsConfirm.Close
Set rsConfirm = Nothing
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
'GET INTERPRETER INFO
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT * FROM interpreter_T WHERE [index] = " & tmpIntr
rsIntr.Open sqlIntr, g_strCONN, 3, 1
If Not rsIntr.EOF Then
	tmpInHouse = ""
	If rsIntr("InHouse") = True Then tmpInHouse = "(In-House)"
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
		if rsHP("reason") <> "" Then mytmpReas = Z_Replace(rsHP("reason"),", ", "|")
		tmpReason = GetReas(mytmpReas)
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
'GET REQUESTING PERSON
Set rsReq = Server.CreateObject("ADODB.RecordSet")
sqlReq = "SELECT * FROM requester_T WHERE [index] = " & RP
rsReq.Open sqlReq, g_strCONN, 3, 1
If Not rsReq.EOF Then
	tmpRP = rsReq("Lname") & ", " & rsReq("Fname") 
	Fon = rsReq("phone") 
	If rsReq("pExt") <> "" Then Fon = Fon & " ext. " & rsReq("pExt")
	Fax = rsReq("fax")
	email = rsReq("email")
	Pcon = GetPrime(RP)
End If
rsReq.Close
Set rsReq = Nothing
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
sqlDept = "SELECT * FROM dept_T WHERE [index] = " & tmpDept
rsDept.Open sqlDept, g_strCONN, 3, 1
If Not rsDept.EOF Then
	tmpDname = rsDept("dept") 
	tmpDeptaddr = rsDept("address") & ", " & rsDept("InstAdrI") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	tmpBaddr = rsDept("Baddress") & ", " & rsDept("BCity") & ", " &  rsDept("Bstate") & ", " & rsDept("Bzip")
	tmpBContact = rsDept("Blname")
	tmpZipInst = ""
	If rsDept("zip") <> "" Then tmpZipInst = rsDept("zip")
	If tmpDeptaddrG = "" Then 
		tmpDeptaddrG = rsDept("address") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	End If
End If
rsDept.Close
Set rsDept = Nothing 
'GET AVAILABLE LANGUAGES
Set rsLang = Server.CreateObject("ADODB.RecordSet")
sqlLang = "SELECT * FROM language_T ORDER BY [Language]"
rsLang.Open sqlLang, g_strCONN, 3, 1
Do Until rsLang.EOF
	tmpL = ""
	If tmpLang = "" Then tmpLang = -1
	If CInt(tmpLang) = rsLang("index") Then tmpL = "selected"
	strLang = strLang	& "<option " & tmpL & " value='" & rsLang("Index") & "'>" &  rsLang("language") & "</option>" & vbCrlf
	strLangChk = strLangChk & "if (xxx == """ & Trim(rsLang("Language")) & """){ " & vbCrLf & _
		"return " & rsLang("index") & ";}"
	rsLang.MoveNext
Loop
rsLang.Close
Set rsLang = Nothing
%>
<!-- #include file="_closeSQL.asp" -->
<html>
	<head>
		<title>Language Bank - Interpreter Request Form - Edit Notes - <%=Request("ID")%></title>
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
			document.frmAssign.action = "action.asp?ctrl=16&ReqID=" + xxx;
			document.frmAssign.submit();
		}
		function CalendarView(strDate)
		{
			document.frmAssign.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmAssign.submit();
		}
		
		-->
		</script>
		<body >
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
								<td class='title' colspan='10' align='center'><nobr> Interpreter Request Form - Edit Notes</td>
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
									<td colspan='10' class='header'><nobr>Language Bank Notes
									</td>
								</tr>
								<tr>
									<td align='right' valign='top'>Notes:</td>
									<td class='confirm' valign='top'>
										<textarea name='txtLBcom' class='main' onkeyup='bawal(this);' style='width: 375px;' rows='6'><%=tmpLBcom%></textarea>
										<input class='btnLnk' type='button' value='Save' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'" onclick='SaveAss(<%=Request("ID")%>) ;'>
										<input class='btnLnk' type='button' value='Back' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'" onclick="document.location='reqconfirm.asp?ID=<%=Request("ID")%>';">
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td class='header' colspan='10'><nobr>Contact Information </td>
								</tr>
								<tr>
									<td align='right'>Request ID:</td>
									<td class='confirm' width='300px'><%=Request("ID")%>&nbsp;<%=tmpEmer%></td>
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
									<td align='right'>Institution:</td>
									<td class='confirm'><%=tmpIname%></td>
								</tr>
								<tr>
									<td align='right'>Department:</td>
									<td class='confirm'><%=tmpDname%></td>
								</tr>
								<tr>
									<td align='right'>Address:</td>
									<td class='confirm'><%=tmpDeptaddr%></td>
								</tr>
								<tr>
									<td align='right'>Billed To:</td>
									<td class='confirm'><%=tmpBContact%></td>
								</tr>
								<tr>
									<td align='right'>Billing Address:</td>
									<td class='confirm'><%=tmpBaddr%></td>
								</tr>
								<% If Request.Cookies("LBUSERTYPE") <> 4 Then %>
									<tr>
										<td align='right' width='15%'>Rate:</td>
										<td class='confirm'><%=tmpInstRate%></td>
									</tr>
								<% End If %>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Requesting Person:</td>
									<td class='confirm'><%=tmpRP%></td>
								</tr>
								<tr>
									<td align='right'>Phone:</td>
									<td class='confirm'><%=fon%></td>
								</tr>
								<tr>
									<td align='right'>Fax:</td>
									<td class='confirm'><%=fax%></td>
								</tr>
								<tr>
									<td align='right'>E-Mail:</td>
									<td class='confirm'><%=email%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' class='header'><nobr>Appointment Information</td>
								</tr>
								<tr>
									<td align='right'>Client Name:</td>
									<td class='confirm'><%=tmplName%>, <%=tmpfName%>
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
								<tr>
									<td align='right'>Rate:</td>
									<td class='confirm'><%=tmpIntrRate%></td>
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
								<tr><td>&nbsp;</td></tr>
										<tr>
									<td align='right'>Billing Comment:</td>
									<td class='confirm'><%=tmpCombil%></td>
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
												<input class='btn' type='button' value='Save' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SaveAss(<%=Request("ID")%>) ;'>
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