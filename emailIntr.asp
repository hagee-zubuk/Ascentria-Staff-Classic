<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
Function Avail(myID, myTime)
	Avail = False
	Set rsAvail = Server.CreateObject("ADODB.RecordSet")
	sqlAvail = "SELECT * FROM Avail_T WHERE intrID = " & myID & " AND Avail = '" & myTime & "'"
	rsAvail.Open sqlAvail, g_strCONN, 3, 1
	If Not rsAvail.EOF Then Avail = True
	rsAvail.Close
	set rsAvail = Nothing
	If Avail Then Exit Function
	Set rsAvail2 = Server.CreateObject("ADODB.RecordSet")
	sqlAvail2 = "SELECT * FROM Avail_T WHERE IntrID = " & myID
	rsAvail2.Open sqlAvail2, g_strCONN, 3, 1
	If rsAvail2.EOF Then Avail = True
	rsAvail2.Close
	set rsAvail2 = Nothing
End Function
Function GetEmail(xxx)
	GetEmail = ""
	Set rsEm = Server.CreateObject("ADODB.RecordSet")
	sqlEm = "SELECT [e-mail] FROM interpreter_T WHERE [index] = " & xxx
	rsEm.Open sqlEm, g_strCONN, 1, 3
	If Not rsEm.EOF Then
		GetEmail = rsEm("e-mail")
	End If
	rsEm.Close
	Set rsEm = Nothing
End Function
If Request("mail") = 1 Then 'Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	'GET EMAIL INFO
	Set rsReq = Server.CreateObject("ADODB.REcordSet")
	sqlReq = "SELECT * FROM request_T WHERE [index]= " & Request("ID")
	rsReq.Open sqlReq, g_strCONN, 1, 3
	If Not rsReq.EOF Then
		appdate = rsReq("appdate")
		apptime = CTime(rsReq("apptimefrom")) & " - " & CTime(rsReq("apptimeto"))
		appCity = GetCity(rsReq("DeptID"))
		If rsReq("cliadd") Then appCity = rsReq("Ccity")
		IntrLang = Ucase(GetLang(rsReq("langID")))
		tmpInst = rsReq("instID")
		tmpDept = rsReq("deptID")
	End If
	rsReq.Close
	Set rsReq = Nothing
	'SEND EMAIL
	'on error resume next
	Set mlMail = CreateObject("CDO.Message")
	'mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	'mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
	'mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 26
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")= 2
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.socketlabs.com"
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 2525
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "server3874"
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "UO2CUSxat9ZmzYD7jkTB"
	mlMail.Configuration.Fields.Update
	mlMail.To = GetEmail(Request("selIntr"))
	mlMail.Bcc = "sysdump1@ascentria.org"
	mlMail.From = "language.services@thelanguagebank.org"
	mlMail.Subject= "Appointment on " & appDate & " at " & apptime & " in " & appCity
	strMSG = "Are you available to do this appointment?<br><br>"& _
		"If you accept this appointment this is the amount you will be reimbursed for mileage and travel time. " & _
		"Payable travel time is " & Z_Czero(Request("txtTravel")) & " hrs. " & _
		"and payable mileage is " & Z_Czero(Request("txtMile")) & "  miles.<br><br>"& _
		"Please reply to this email or contact " & Request.Cookies("LBUsrName") & " of LanguageBank.<br><br>"& _
		"Thank you.<BR><BR><BR>" & _
		"<font color='#FFFFFF'>" & Request("adr1") & "|" & Request("adr2") & "|" & Request("zip1") & "|" & Request("zip2") & "</font>"
	mlMail.HTMLBody = "<html><body>" & vbCrLf & strMSG & vbCrLf & "</body></html>"
	mlMail.Send
	set mlMail=nothing
	'save to notes
	IntrName = GetIntr2(Request("selIntr"))
	Set rsNotes = Server.CreateObject("ADODB.RecordSet")
	sqlNotes = "SELECT LBComment FROM request_T WHERE [index] = " & Request("ReqID")
	rsNotes.Open sqlNotes, g_StrCONN, 1, 3
	If Not rsNotes.EOF Then
		rsNotes("LBComment") = rsNotes("LBComment") & vbCrlF & "Email sent to " & IntrName & " on " & now & " for availability"
		rsNotes.Update
	End If
	rsNotes.CLose
	Set rsNotes = Nothing
	If SaveHist(Request("ReqID"), "emailIntr.asp") Then
		 'Session("MSG") = "HIST SAVED"
		End If
	Session("MSG") = "Email Sent."
End If
	'PREPARE EMAIL	
	strAvail = 0
	Set rsReq = Server.CreateObject("ADODB.REcordSet")
	sqlReq = "SELECT * FROM request_T WHERE [index] = " & Request("ID")
	rsReq.Open sqlReq, g_strCONN, 1, 3
	If Not rsReq.EOF Then
		appdate = rsReq("appdate")
		apptime = rsReq("apptimefrom") & " - " & rsReq("apptimeto")
		tmpAppTFrom = rsReq("apptimefrom")
		tmpAppTTo = rsReq("apptimeto")
		appCity = GetCity(rsReq("DeptID"))
		If rsReq("cliadd") Then appCity = rsReq("Ccity")
		IntrLang = Ucase(GetLang(rsReq("langID")))
		tmpAvail = Weekday(appdate) & "," & Hour(tmpAppTFrom)
		tmpInst = rsReq("instID")
		tmpDept = rsReq("deptID")
	End If
	rsReq.Close
	Set rsReq = Nothing
	strSubj = "Appointment on " & appDate & " at " & apptime & " in " & appCity
	strMSG = "Are you available to do this appointment?" & vbCrlf & vbCrlf & _
		"Please reply to this email or contact " & Request.Cookies("LBUsrName") & " of LanguageBank." & vbCrlf & vbCrlf & _
		"If you accept this appointment this is the amount you will be reimbursed for mileage and travel time." & vbCrlf & _
		"Payable travel time is " & Z_Czero(Request("txtTravel")) & " hrs. and payable mileage is " & Z_Czero(Request("txtMile")) & " miles." & vbCrlf & vbCrlf & _
		"Thank you."
	'INTERPRETER LIST
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "SELECT * FROM interpreter_T WHERE (Upper(Language1) = '" & IntrLang & "' OR Upper(Language2) = '" & IntrLang & "' OR Upper(Language3) = '" & IntrLang & _
		"' OR Upper(Language4) = '" & IntrLang & "' OR Upper(Language5) = '" & IntrLang & "') AND Active = 1 AND [e-mail] <> '' ORDER BY [Last Name], [First Name]"
	rsIntr.Open sqlIntr, g_strCONN, 1, 3
	Do Until rsIntr.EOF
		'include vacation
		myIntr = ""
		mark = 0
		If cint(Request("selIntr")) = rsIntr("index") Then myIntr = "selected"
		tmpIntr = cint(Request("selIntr"))
		tmpIntrName = rsIntr("Last Name") & ", " & rsIntr("First Name")
		'If OnVacation(rsIntr("index"), appdate) = False Then 'If isNull(rsIntrLang("vacfrom")) Then
			If tmpIntr = rsIntr("index") Or (Avail(rsIntr("index"), tmpAvail) And NotRestrict(rsIntr("index"), tmpInst, tmpDept)) Then
				mark = 1
				rest = ""
				If NotRestrict(rsIntr("index"), tmpInst, tmpDept) = false Then rest = " (restricted)"
				strIntr = strIntr	& "<option " & myIntr & " value='" & rsIntr("Index") & "'>" & rsIntr("last name") & ", " & rsIntr("first name") & rest & "</option>" & vbCrlf
				tmpIntrName = CleanMe(rsIntr("last name")) & ", " & CleanMe(rsIntr("first name"))
				strAvail = SkedCheck(rsIntr("Index"), Request("ID"), appdate, tmpAppTFrom, tmpAppTTo)
			End If
		'End If
		If mark = 0 And NotRestrict(rsIntr("index"), tmpInst, tmpDept) Then 'And OnVacation(rsIntr("index"), appdate) = False Then
			strIntr2 = strIntr2 & "<option value='" & rsIntr("index") & "' " & IntrSel & ">" & tmpIntrName & "</option>" & vbCrlf
			strAvail = SkedCheck(rsIntr("Index"), Request("ID"), appdate, tmpAppTFrom, tmpAppTTo)
		End If
		''''OLD
		'If isNull(rsIntr("vacfrom")) Then
		'	If Avail(rsIntr("index"), tmpAvail) Then
		'		mark = 1
		'		strIntr = strIntr & "<option value='" & rsIntr("index") & "' " & myIntr & ">" & IntrName & "</option>" & vbCrlf
		'		'strAvail = SkedCheck(rsIntr("Index"), Request("ID"), appdate, tmpAppTFrom, tmpAppTTo)
		'	End If
		'Else
		'	If Not (appdate >= rsIntr("vacfrom") And appdate <= rsIntr("vacto")) Then
		'		If Avail(rsIntr("index"), tmpAvail) Then
		'			mark = 1
		'			strIntr = strIntr & "<option value='" & rsIntr("index") & "' " & myIntr & ">" & IntrName & "</option>" & vbCrlf
		'			'strAvail = SkedCheck(rsIntr("Index"), Request("ID"), appdate, tmpAppTFrom, tmpAppTTo)
		'		End If
		'	End If
		'End If
		'If mark = 0 Then
		'	strIntr2 = strIntr2 & "<option value='" & rsIntr("index") & "' " & myIntr & ">" & IntrName & "</option>" & vbCrlf
			'strAvail = SkedCheck(rsIntr("Index"), Request("ID"), appdate, tmpAppTFrom, tmpAppTTo)
		'End If
		strJScript2 = strJScript2 & "if (Intr == " & rsIntr("Index") & ") {" & vbCrLf & _
			"document.frmEmail.chkAvail.value = " & SkedCheck(rsIntr("Index"), Request("ID"), appdate, tmpAppTFrom, tmpAppTTo) &"}; " & vbCrLf & _
		rsIntr.MoveNext
	Loop
	rsIntr.Close
	Set rsIntr = Nothing
	If strIntr2 <> "" Then strIntr = strIntr & "<option value='0'>-----</option>" & vbCrlf & strIntr2
	
'End If
%>
<!-- #include file="_closeSQL.asp" -->
<html>
	<head>
		<title>Email Interpreter</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
			function sendMe()
			{
				if (document.frmEmail.selIntr.value == 0)
				{
					alert("ERROR: Please select an interpreter.")
					return;
				}
				if (document.frmEmail.chkAvail.value == 1) {
					var ans = window.confirm("WARNING: Interpreter already has an appointment for this date and time range.\nPlease check the calendar. \nClick OK to override. \nClick Cancel to stop."); 
					if (ans) {
						document.frmEmail.mail.value = 1;
						document.frmEmail.action = "emailIntr.asp";
						document.frmEmail.submit();	
					}
					else {
						return;
					}
				}
				else {
					document.frmEmail.mail.value = 1;
					document.frmEmail.action = "emailIntr.asp";
					document.frmEmail.submit();	
				}
			}
			function GetMile(xxx)
			{
				if (document.getElementById("btnSend").value == 'Send') {
					document.getElementById("btnSend").value = "Please Wait..."
					document.getElementById("btnSend").disabled = true
					document.frmEmail.action = "Intrmile.asp?selIntr=" + xxx + "&RID=" + <%=Request("ID")%>;
					document.frmEmail.submit();	
				}
			}
			function IntrInfo(Intr) {
				if (Intr == -1 || Intr == 0) {
					document.frmEmail.chkAvail.value = 0;
				}
				<%=strJScript2%>
			}
		</script>
	</head>
	<body onload='IntrInfo(<%=tmpIntr%>);'>
		<form name='frmEmail' method='post'>
			<table cellpadding='1' cellspacing='0' border='0'>
				<tr><td>&nbsp;
					<input name="mail" value="" type="hidden" />
					</td></tr>
				<tr>
					<td align='right'>&nbsp;Interpreter:</td>
					<td align='left'>&nbsp;
						<select class='seltxt' name='selIntr' style='width: 150px;' onchange='GetMile(this.value); IntrInfo(this.value);' onfocus='IntrInfo(this.value);'>
							<option value='0'>&nbsp;</option>
							<%=strIntr%>
						</select>
					</td>
				</tr>
				<tr>
					<td align='right'>Subject:</td>
					<td align='left'>&nbsp;<b><%=strSubj%></b></td>
				</tr>
				<tr>
					<td align='right' valign='top'>Message:</td>
					<td align='left'>&nbsp;
						<textarea name='txtMSG' readonly cols='48' rows='10'>
							<%=strMSG%>
						</textarea>
					</td>
				</tr>
				<tr>
					<td align='right' colspan='2'>
						<input class='btn' type='button' name="btnSend" id="btnSend" value='Send' style='width: 100px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='sendMe();'>
						<input class='btn' type='button' value='Close' style='width: 100px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='window.close();'>
						<input type='hidden' name='ReqID' value='<%=Request("ID")%>'>
						<input type="hidden" name="id" value='<%=Request("ID")%>' />
						<input type="hidden" name="txtTravel" value='<%=Request("txtTravel")%>' />
						<input type="hidden" name="txtMile" value='<%=Request("txtMile")%>' />
						<input type='hidden' name='adr1'  value='<%=Request("adr1")%>'>
						<input type='hidden' name='adr2'  value='<%=Request("adr2")%>'>
						<input type='hidden' name='zip1'  value='<%=Request("zip1")%>'>
						<input type='hidden' name='zip2'  value='<%=Request("zip2")%>'>
						<input type='hidden' name='chkAvail'>	
					</td>
				</tr>
			</table>
		</form>
	</body>
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