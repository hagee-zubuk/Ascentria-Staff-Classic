<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function
Function CleanFax(strFax)
	CleanFax = Replace(strFax, "-", "") 
End Function
Function GetPrime(xxx)
	GetPrime = ""
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT * FROM requester_T WHERE index = " & xxx
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
Function GetPrime2(xxx)
	GetPrime2 = ""
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT * FROM interpreter_T WHERE index = " & xxx
	rsRP.Open sqlRP, g_strCONN, 3, 1
	If Not rsRP.EOF Then
		If rsRP("prime") = 0 Then
			GetPrime2 = rsRP("E-mail")
		ElseIf rsRP("prime") = 1 Or rsRP("prime") = 2 Then
			'GetPrime = rsRP("Phone")
			GetPrime2 = ""
		ElseIf rsRP("prime") = 3 Then
			GetPrime2 = CleanFax(Trim(rsRP("Fax"))) & "@emailfaxservice.com" 
		End If
	End If
	rsRP.Close
	set rsRP = Nothing
End Function
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
Function GetLoc(xxx)
	Select Case xxx
		Case 0 
			GetLoc = "Front Door"
		Case 1
			GetLoc = "Cafeteria"
		Case 2
			GetLoc = "Registration Desk"
		Case 3
			GetLoc = "Department"
		Case 4
			GetLoc = "OTHER"
	End Select
End Function
tmpPage = "document.frmConfirm."
'GET REQUEST
Set rsConfirm = Server.CreateObject("ADODB.RecordSet")
sqlConfirm = "SELECT * FROM Request_T WHERE index = " & Request("ID")
rsConfirm.Open sqlConfirm, g_strCONN, 3, 1
If Not rsConfirm.EOF Then
	TS = rsConfirm("timestamp")
	RP = rsConfirm("reqID") 
	tmpClient = ""
	If rsConfirm("client") = True Then tmpClient = " (LSS Client)"
	tmpName = rsConfirm("clname") & ", " & rsConfirm("cfname") & tmpClient
	tmpAddr = rsConfirm("caddress") & ", " & rsConfirm("cCity") & ", " &  rsConfirm("cstate") & ", " & rsConfirm("czip")
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
	tmpEmer = ""
	If rsConfirm("Emergency") = True Then tmpEmer = "(EMERGENCY)" 
	tmpCom = rsConfirm("Comment")
	Statko = GetMyStatus(rsConfirm("Status"))
	tmpBilHrs = rsConfirm("Billable")
	tmpActTFrom = Z_FormatTime(rsConfirm("astarttime")) 
	tmpActTTo = Z_FormatTime(rsConfirm("aendtime"))
	tmpBilTInst = rsConfirm("TT_Inst")
	tmpBilTIntr = rsConfirm("TT_Intr")
	tmpBilMInst = rsConfirm("M_Inst")
	tmpBilMIntr = rsConfirm("M_Intr")
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
'HP DATA
If tmpHPID <> 0  THen
	Set rsHP = Server.CreateObject("ADODB.RecordSet")
		sqlHP = "SELECT * FROM Appointment_T WHERE index = " & tmpHPID
	rsHP.Open sqlHP, g_StrCONNHP, 3, 1
	If Not rsHP.EOF Then
		tmpCallMe = ""
		If rsHP("callme") = True Then tmpCallMe = "* Call patient to remind of appointment"
		'tmpReason = rsHP("reason")
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
'GET REQUESTING PERSON
Set rsReq = Server.CreateObject("ADODB.RecordSet")
sqlReq = "SELECT * FROM requester_T WHERE index = " & RP
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
sqlInst = "SELECT * FROM institution_T WHERE index = " & tmpInst
rsInst.Open sqlInst, g_strCONN, 3, 1
If Not rsInst.EOF Then
	tmpIname = rsInst("Facility") 
	'tmpIaddr = rsInst("address") & ", " & rsInst("City") & ", " &  rsInst("state") & ", " & rsInst("zip")
	'tmpBaddr = rsInst("Baddress") & ", " & rsInst("BCity") & ", " &  rsInst("Bstate") & ", " & rsInst("Bzip")
	'tmpBContact = rsInst("Blname") & ", " & rsInst("Bfname")
End If
rsInst.Close
Set rsInst = Nothing 
'GET DEPARTMENT
Set rsDept = Server.CreateObject("ADODB.RecordSet")
sqlDept = "SELECT * FROM dept_T WHERE index = " & tmpDept
rsDept.Open sqlDept, g_strCONN, 3, 1
If Not rsDept.EOF Then
	tmpDname = rsDept("dept") 
	tmpDeptaddr = rsDept("address") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	tmpBaddr = rsDept("Baddress") & ", " & rsDept("BCity") & ", " &  rsDept("Bstate") & ", " & rsDept("Bzip")
	tmpBContact = rsDept("Blname")
	tmpZipInst = ""
	If rsDept("zip") <> "" Then tmpZipInst = rsDept("zip")
End If
rsDept.Close
Set rsDept = Nothing 
'GET LANGUAGE
Set rsLang = Server.CreateObject("ADODB.RecordSet")
sqlLang  = "SELECT * FROM language_T WHERE index = " & tmpLang
rsLang.Open sqlLang , g_strCONN, 3, 1
If Not rsLang.EOF Then
	tmpSalita = rsLang("language") 
End If
rsLang.Close
Set rsLang = Nothing 
'GET INTERPRETER INFO
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT * FROM interpreter_T WHERE index = " & tmpIntr
rsIntr.Open sqlIntr, g_strCONN, 3, 1
If Not rsIntr.EOF Then
	tmpInHouse = ""
	If rsIntr("InHouse") = True Then tmpInHouse = "(In-House)"
	tmpIntrName = rsIntr("Last Name") & ", " & rsIntr("First Name") & " " & tmpInHouse
	tmpIntrEmail = rsIntr("E-mail")
	tmpIntrP1 = rsIntr("Phone1")
	If rsIntr("P1Ext") <> "" Then tmpIntrP1 = tmpIntrP1 & " ext. " &  rsIntr("P1Ext")
	tmpIntrP2 = rsIntr("Phone2")
	tmpIntrFax = rsIntr("Fax")
	tmpIntrAdd = rsIntr("address1") & ", " & rsIntr("City") & ", " &  rsIntr("state") & ", " & rsIntr("Zip Code")
	PconIntr = GetPrime2(tmpIntr)
	tmpZipIntr = ""
	If rsIntr("Zip Code") <> "" Then tmpZipIntr = rsIntr("Zip Code")
Else
	tmpIntrName = "<i>To be assigned.</i>"
End If
rsIntr.Close
Set rsIntr = Nothing
%>
<html>
	<head>
		<title>Language Bank - Request Confirmation</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function chkEmail(tmpemail)
		{
			if (tmpemail == undefined || tmpemail == "")
				{
					alert("ERROR: Primary Contact is blank or invalid.");
				}
			else
				{
					var ans = window.confirm("This action will send an email/fax to the requesting person.\nClick Cancel to stop.");
					if (ans)
					{
						document.frmConfirm.action = "email.asp?sino=0&emailadd='" + tmpemail + "' &HID=" + <%=Request("ID")%>;
						document.frmConfirm.submit();
					}
				}
		}
		function chkEmail2(tmpemail)
		{
			if (tmpemail == undefined || tmpemail == "")
				{
					alert("ERROR: Primary Contact is blank or invalid.");
				}
			else
				{
					var ans = window.confirm("This action will send an email/fax to the interpreter.\nClick Cancel to stop.");
					if (ans)
					{
						document.frmConfirm.action = "email.asp?sino=1&emailadd='" + tmpemail + "' &HID=" + <%=Request("ID")%>;
						document.frmConfirm.submit();
					}
				}
		}
		function EditMe()
		{
			document.frmConfirm.action = "main.asp?ID=" + <%=Request("ID")%>;
			document.frmConfirm.submit();
		}
		function EditMe2()
		{
			document.frmConfirm.action = "mainbill.asp?ID=" + <%=Request("ID")%>;
			document.frmConfirm.submit();
		}
		function AssignMe()
		{
			document.frmConfirm.action = "assign.asp?ID=" + <%=Request("ID")%>;
			document.frmConfirm.submit();
		}
		function PopMe(zzz)
		{
			if (zzz !== undefined)
				{
				newwindow = window.open('print.asp?ID=' + zzz,'name','height=1056,width=816,scrollbars=1,directories=0,status=0,toolbar=0,resizable=1');
				if (window.focus) {newwindow.focus()}
				}
		}
		function CalendarView(strDate)
		{
			document.frmConfirm.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmConfirm.submit();
		}
		function CopyMe()
		{
			document.frmConfirm.action = "main.asp?Clone=" + <%=Request("ID")%>;
			document.frmConfirm.submit();
		}
		//function PopMe2(instzip, intrzip)
		//{
		//	if (instzip == "" || intrzip == "")
		//	{
		//		alert("Error: Institution's zip code and/or Interpreter's zip code is blank.")
		//		return;
		//	}
		//	else
		//	{
		//		var zip1 = instzip; 
		//		var zip2 = intrzip;
		//		var zipus = zip1 + "|" + zip2
		//		//alert(zipus);
		//		newwindow2 = window.open('zipcalc.asp?zipus=' + zipus,'name','height=150,width=400,scrollbars=0,directories=0,status=1,toolbar=0,resizable=0');
		//		if (window.focus) {newwindow2.focus()}
		//	}
		//}
		function PopMe2(xxx)
		{
			//if (instzip == "" || intrzip == "")
			//{
			//	alert("Error: Institution's zip code and/or Interpreter's zip code is blank.")
			//	return;
			//}
			//else
			//{
			//	var zip1 = instzip; 
			//	var zip2 = intrzip;
			//	var zipus = zip1 + "|" + zip2
				//alert(zipus);
				newwindow2 = window.open('travel.asp?ReqID=' + xxx,'name','height=500,width=800,scrollbars=1,directories=0,status=1,toolbar=0,resizable=0');
				if (window.focus) {newwindow2.focus()}
			//}
		}
		-->
		</script>
		<body onload='PopMe(<%=Request("PID")%>);'>
			<form method='post' name='frmConfirm'>
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
									<td class='title' colspan='10' align='center'><nobr>Request Confirmation</td>
								</tr>
								<tr>
									<td align='center' colspan='10' class='RemME'>
										<%=tmpSentReq%><br>
										<%=tmpSentIntr%><br>
										<%=tmpPrint%>
									</td>
								</tr>
								<tr>
									<td align='center' colspan='10'><span class='error'><%=Session("MSG")%></span></td>
								</tr>
								<tr>
									<td class='header' colspan='10'><nobr>Contact Information </td>
								</tr>
								<tr>
									<td align='right'>Request ID:</td>
									<td class='confirm' width='300px'><%=Request("ID")%>&nbsp;<%=tmpEmer%></td>
									<input type='hidden' name='HID' value='<%=Request("ID")%>'>
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
								<tr>
									<td align='right' width='15%'>Rate:</td>
									<td class='confirm'><%=tmpInstRate%></td>
								</tr>
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
									<td class='confirm'>
										<%=tmpName%>
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
									<td align='right'>Client Phone:</td>
									<td class='confirm'><%=tmpFon%></td>
								</tr>
								<tr>
									<td align='right'>Client Alter. Phone:</td>
									<td class='confirm'><%=tmpAFon%></td>
								</tr>
								<tr>
									<td align='right'>Directions / Landmarks:</td>
									<td class='confirm'><%=tmpdir%></td>
								</tr>
								<tr>
									<td align='right'>Special Circumstances:</td>
									<td class='confirm'><%=tmpSC%></td>
								</tr>
								<tr>
									<td align='right'>DOB:</td>
									<td class='confirm'><%=tmpDOB%></td>
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
									<td align='right'>Appointment Location:</td>
									<td class='confirm'><%=tmpAppLoc%></td>
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
									<td align='right' width='15%'>E-Mail:</td>
									<td class='confirm'><%=tmpIntrEmail%></td>
								</tr>
								<tr>
									<td align='right' width='15%'>Home Phone:</td>
									<td class='confirm'><%=tmpIntrP1%></td>
								</tr>
								<tr>
									<td align='right' width='15%'>Mobile Phone:</td>
									<td class='confirm'><%=tmpIntrP2%></td>
								</tr>
								<tr>
									<td align='right' width='15%'>Fax:</td>
									<td class='confirm'><%=tmpIntrFax%></td>
								</tr>
								<tr>
									<td align='right'>Address:</td>
									<td class='confirm'><%=tmpIntrAdd%></td>
								</tr>
								<tr>
									<td align='right'>Rate:</td>
									<td class='confirm'><%=tmpIntrRate%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Comment:</td>
									<td class='confirm'><%=tmpCom%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' class='header'><nobr>Other Information</td>
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
								<tr>
									<td colspan='10' align='center' height='100px' valign='bottom'>
										<input class='btn' type='button' style='width: 125px;' value='View in Calendar' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='calendarview2.asp?appdate=<%=tmpAppDate%>'">
										<input class='btn' type='button' style='width: 125px;' value='Print' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='PopMe(<%=Request("ID")%>);'>
										<input class='btn' type='button' style='width: 125px;' value='Edit' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='EditMe();'>
										<input class='btn' type='button' style='width: 125px;' value='Clone Appt.' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'"  onclick='CopyMe();'>
										<input class='btn' type='button' style='width: 125px;' value="Zip Code Calculator" name="zipcalc"
											onclick='PopMe2(<%=Request("ID")%>);' title='Zip code calculator' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
									</td>
								</tr>
								<tr>
									<td colspan='10' align='center' valign='bottom'>
										<input class='btn' type='button' style='width: 158px;' value='Send to Requesting Person' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="chkEmail('<%=Pcon%>');">
										<input class='btn' type='button' style='width: 156px;' value='Send to Interpreter' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="chkEmail2('<%=PconIntr%>');">
										<input type='hidden' name='txtInstZip' value='<%=tmpZipInst%>'>
										<input type='hidden' name='txtIntrZip' value='<%=tmpZipIntr%>'>
										<input class='btn' type='button' style='width: 156px;' value='Assign Interpreter' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'"  onclick='AssignMe();'>
										<input class='btn' type='button' style='width: 157px;' value='Service Verification' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'"  onclick='EditMe2();'>
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
	tmpMSG = Session("MSG")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>